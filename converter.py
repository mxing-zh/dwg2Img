from __future__ import annotations

from concurrent.futures import FIRST_COMPLETED, Future, ProcessPoolExecutor, ThreadPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path
import gc
import os
import shutil
import subprocess
import tempfile
from typing import Iterable, List


SUPPORTED_IMAGE_FORMATS = {"png", "jpg", "jpeg"}
LAYOUT_MODELS = {"auto", "model", "layout"}
COLOR_MODES = {"bw", "original"}


def auto_workers() -> int:
    cpu = os.cpu_count() or 1
    if cpu <= 2:
        return 1
    return max(1, cpu - 1)


def auto_oda_workers(render_workers: int) -> int:
    # ODA conversion is I/O + external process heavy; keep conservative.
    return max(1, min(2, render_workers // 2 or 1))


@dataclass
class ConvertConfig:
    input_root: Path
    output_root: Path
    image_format: str = "png"
    dpi: int = 96
    mirror_structure: bool = True
    oda_converter: Path | None = None
    layout_mode: str = "auto"
    preferred_layout: str | None = None
    color_mode: str = "bw"
    max_workers: int = 0
    oda_workers: int = 0
    oda_batch_size: int = 200

    def normalized_format(self) -> str:
        fmt = self.image_format.lower().strip(".")
        if fmt == "jpeg":
            fmt = "jpg"
        if fmt not in SUPPORTED_IMAGE_FORMATS:
            raise ValueError(f"Unsupported image format: {self.image_format}")
        return fmt

    def normalized_layout_mode(self) -> str:
        mode = self.layout_mode.lower().strip()
        if mode not in LAYOUT_MODELS:
            raise ValueError(f"Unsupported layout mode: {self.layout_mode}")
        return mode

    def normalized_color_mode(self) -> str:
        mode = self.color_mode.lower().strip()
        if mode not in COLOR_MODES:
            raise ValueError(f"Unsupported color mode: {self.color_mode}")
        return mode

    def normalized_workers(self) -> int:
        if self.max_workers <= 0:
            return auto_workers()
        return max(1, self.max_workers)

    def normalized_oda_workers(self, render_workers: int) -> int:
        if self.oda_workers <= 0:
            return auto_oda_workers(render_workers)
        return max(1, self.oda_workers)

    def normalized_oda_batch_size(self) -> int:
        return max(20, self.oda_batch_size)


def discover_dwgs(input_root: Path) -> List[Path]:
    return sorted(p for p in input_root.rglob("*.dwg") if p.is_file())


def resolve_output_path(dwg_file: Path, config: ConvertConfig) -> Path:
    fmt = config.normalized_format()
    if config.mirror_structure:
        rel = dwg_file.relative_to(config.input_root)
        destination = config.output_root / rel
        destination.parent.mkdir(parents=True, exist_ok=True)
        return destination.with_suffix(f".{fmt}")

    config.output_root.mkdir(parents=True, exist_ok=True)
    return (config.output_root / dwg_file.name).with_suffix(f".{fmt}")


def _run_oda_converter(input_root: Path, dxf_root: Path, oda_converter: Path) -> None:
    if not oda_converter.exists():
        raise FileNotFoundError(f"ODA converter not found: {oda_converter}")

    cmd = [
        str(oda_converter),
        str(input_root),
        str(dxf_root),
        "ACAD2018",
        "DXF",
        "1",
        "1",
        "*.dwg",
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(
            "ODA conversion failed.\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout: {proc.stdout}\n"
            f"stderr: {proc.stderr}"
        )


def _pick_layout(doc, mode: str, preferred_layout: str | None) -> tuple[object, str]:
    if mode == "model":
        return doc.modelspace(), "Model"

    preferred = preferred_layout.strip().lower() if preferred_layout else ""
    paperspaces = [layout for layout in doc.layouts if layout.name.lower() != "model"]

    def has_renderable_entities(layout) -> bool:
        return len(layout) > 0

    if mode == "layout":
        if preferred:
            for layout in paperspaces:
                if layout.name.lower() == preferred:
                    return layout, layout.name
        for layout in paperspaces:
            if has_renderable_entities(layout):
                return layout, layout.name
        return doc.modelspace(), "Model"

    if preferred:
        for layout in paperspaces:
            if layout.name.lower() == preferred:
                return layout, layout.name

    for layout in paperspaces:
        if has_renderable_entities(layout):
            return layout, layout.name
    return doc.modelspace(), "Model"


def _layout_paper_inches(layout) -> tuple[float, float] | None:
    if layout.name.lower() == "model":
        return None

    width = float(getattr(layout.dxf, "paper_width", 0.0) or 0.0)
    height = float(getattr(layout.dxf, "paper_height", 0.0) or 0.0)
    if width <= 0 or height <= 0:
        return None

    units = int(getattr(layout.dxf, "plot_paper_units", 0) or 0)
    if units == 1:
        return width / 25.4, height / 25.4
    if units == 2:
        return width / 96.0, height / 96.0
    return width, height


def _safe_bbox_extents(layout):
    import ezdxf
    from ezdxf import bbox

    try:
        ext = bbox.extents(layout, fast=True)
    except ezdxf.DXFError:
        return None
    if ext is None:
        return None
    return ext


def _safe_bbox_size(layout) -> tuple[float, float] | None:
    ext = _safe_bbox_extents(layout)
    if ext is None:
        return None

    width = float(ext.size.x)
    height = float(ext.size.y)
    if width <= 0 or height <= 0:
        return None
    return width, height


def _figure_size_inches(layout) -> tuple[float, float]:
    paper_size = _layout_paper_inches(layout)
    if paper_size:
        return paper_size

    bbox_size = _safe_bbox_size(layout)
    if bbox_size:
        width, height = bbox_size
        longest = max(width, height)
        if longest <= 0:
            return 12.0, 8.0
        scale = 12.0 / longest
        return max(width * scale, 2.0), max(height * scale, 2.0)

    return 12.0, 8.0


def _normalize_image_output(
    image_file: Path, image_format: str, dpi: int, expected_size: tuple[int, int] | None = None
) -> tuple[int, int]:
    from PIL import Image

    fmt = image_format.upper()
    if fmt == "JPG":
        fmt = "JPEG"

    with Image.open(image_file) as im:
        rgb = im.convert("RGB")
        if expected_size and rgb.size != expected_size:
            rgb = rgb.resize(expected_size, Image.Resampling.LANCZOS)
        if fmt == "JPEG":
            rgb.save(image_file, format=fmt, quality=95, subsampling=0, dpi=(dpi, dpi))
        else:
            rgb.save(image_file, format=fmt, dpi=(dpi, dpi))
        return rgb.size


def _render_dxf_to_image(
    dxf_file: Path,
    image_file: Path,
    image_format: str,
    dpi: int,
    layout_mode: str,
    preferred_layout: str | None,
    color_mode: str,
) -> tuple[str, tuple[int, int]]:
    import matplotlib.pyplot as plt
    import ezdxf
    from ezdxf.addons.drawing import Frontend, RenderContext
    from ezdxf.addons.drawing.config import (
        BackgroundPolicy,
        ColorPolicy,
        Configuration,
        ProxyGraphicPolicy,
    )
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

    doc = ezdxf.readfile(dxf_file)
    target_layout, layout_name = _pick_layout(doc, layout_mode, preferred_layout)

    fig_w, fig_h = _figure_size_inches(target_layout)
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")

    ctx = RenderContext(doc)
    out = MatplotlibBackend(ax)
    color_policy = ColorPolicy.BLACK if color_mode == "bw" else ColorPolicy.COLOR
    draw_config = Configuration.defaults().with_changes(
        color_policy=color_policy,
        background_policy=BackgroundPolicy.WHITE,
        proxy_graphic_policy=ProxyGraphicPolicy.PREFER,
    )
    Frontend(ctx, out, config=draw_config).draw_layout(target_layout, finalize=True)

    ext = _safe_bbox_extents(target_layout)
    if ext is not None and ext.size.x > 0 and ext.size.y > 0:
        pad_x = ext.size.x * 0.02
        pad_y = ext.size.y * 0.02
        ax.set_xlim(ext.extmin.x - pad_x, ext.extmax.x + pad_x)
        ax.set_ylim(ext.extmin.y - pad_y, ext.extmax.y + pad_y)
        ax.margins(0)

    fig.savefig(image_file, dpi=dpi, facecolor="white", transparent=False)
    actual_inches = fig.get_size_inches()
    width_px = int(round(actual_inches[0] * dpi))
    height_px = int(round(actual_inches[1] * dpi))
    plt.close(fig)
    plt.close("all")

    final_size = _normalize_image_output(
        image_file,
        image_format,
        dpi,
        expected_size=(max(width_px, 1), max(height_px, 1)),
    )

    del out, ctx, doc
    gc.collect()
    return layout_name, final_size


def _prepare_batch_inputs(files: list[Path], input_root: Path, staged_input_root: Path) -> None:
    for src in files:
        rel = src.relative_to(input_root)
        dst = staged_input_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            os.link(src, dst)
        except OSError:
            shutil.copy2(src, dst)


def _convert_batch_to_dxf(
    batch_id: int,
    files: list[Path],
    input_root: Path,
    oda_converter: Path,
    work_root: Path,
) -> tuple[int, Path, list[Path]]:
    staged_input_root = work_root / f"batch_{batch_id:05d}_dwg"
    staged_dxf_root = work_root / f"batch_{batch_id:05d}_dxf"
    staged_input_root.mkdir(parents=True, exist_ok=True)
    staged_dxf_root.mkdir(parents=True, exist_ok=True)
    _prepare_batch_inputs(files, input_root, staged_input_root)
    _run_oda_converter(staged_input_root, staged_dxf_root, oda_converter)
    return batch_id, staged_dxf_root, files


def _render_worker(task: tuple[Path, Path, str, int, str, str | None, str]) -> tuple[bool, str]:
    dxf_file, output_file, fmt, dpi, layout_mode, preferred_layout, color_mode = task
    try:
        output_file.parent.mkdir(parents=True, exist_ok=True)
        layout_name, size_px = _render_dxf_to_image(
            dxf_file=dxf_file,
            image_file=output_file,
            image_format=fmt,
            dpi=dpi,
            layout_mode=layout_mode,
            preferred_layout=preferred_layout,
            color_mode=color_mode,
        )
        return True, f"{dxf_file.name} -> {output_file.name} ({layout_name}, {size_px[0]}x{size_px[1]}px)"
    except Exception as exc:
        return False, f"{dxf_file}: {exc}"


def _chunked(items: list[Path], size: int) -> list[list[Path]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def batch_convert(config: ConvertConfig) -> Iterable[str]:
    input_root = config.input_root.resolve()
    if not input_root.exists():
        raise FileNotFoundError(f"Input directory not found: {input_root}")

    layout_mode = config.normalized_layout_mode()
    color_mode = config.normalized_color_mode()
    image_format = config.normalized_format()
    render_workers = config.normalized_workers()
    oda_workers = config.normalized_oda_workers(render_workers)
    batch_size = config.normalized_oda_batch_size()

    dwg_files = discover_dwgs(input_root)
    if not dwg_files:
        yield "未找到 DWG 文件。"
        return

    converter_path = config.oda_converter
    if converter_path is None:
        raise ValueError("必须提供 ODA File Converter 可执行文件路径。")

    total = len(dwg_files)
    batches = _chunked(dwg_files, batch_size)
    yield f"任务总数: {total}，渲染并发: {render_workers}，ODA并发: {oda_workers}，批大小: {batch_size}"

    converted = 0
    failed = 0
    skipped = 0
    processed = 0

    with tempfile.TemporaryDirectory(prefix="dwg2img_work_") as tmp:
        work_root = Path(tmp)

        with ThreadPoolExecutor(max_workers=oda_workers) as oda_pool, ProcessPoolExecutor(
            max_workers=render_workers,
            max_tasks_per_child=25,
        ) as render_pool:
            pending_oda: dict[Future, int] = {}
            queued = 0
            for batch_id, files in enumerate(batches, 1):
                future = oda_pool.submit(
                    _convert_batch_to_dxf,
                    batch_id,
                    files,
                    input_root,
                    converter_path,
                    work_root,
                )
                pending_oda[future] = batch_id
                queued += 1

            while pending_oda:
                done, _ = wait(pending_oda.keys(), return_when=FIRST_COMPLETED)
                for fut in done:
                    batch_id = pending_oda.pop(fut)
                    try:
                        _, dxf_root, batch_files = fut.result()
                    except Exception as exc:
                        failed += len(batches[batch_id - 1])
                        processed += len(batches[batch_id - 1])
                        yield f"批次 {batch_id}/{queued} ODA 转换失败，已跳过该批: {exc}"
                        yield f"进度: {processed}/{total} ({processed / total:.1%})"
                        continue

                    tasks = []
                    for dwg_file in batch_files:
                        rel = dwg_file.relative_to(input_root)
                        dxf_file = (dxf_root / rel).with_suffix(".dxf")
                        if not dxf_file.exists():
                            skipped += 1
                            processed += 1
                            continue

                        output_file = resolve_output_path(dwg_file, config)
                        tasks.append((dxf_file, output_file, image_format, config.dpi, layout_mode, config.preferred_layout, color_mode))

                    if not tasks:
                        yield f"批次 {batch_id}/{queued} 无可渲染DXF。"
                        yield f"进度: {processed}/{total} ({processed / total:.1%})"
                        shutil.rmtree(dxf_root, ignore_errors=True)
                        continue

                    futures = [render_pool.submit(_render_worker, t) for t in tasks]
                    for rf in futures:
                        ok, _ = rf.result()
                        processed += 1
                        if ok:
                            converted += 1
                        else:
                            failed += 1
                        if processed % 10 == 0 or processed == total:
                            yield f"进度: {processed}/{total} ({processed / total:.1%})"

                    shutil.rmtree(dxf_root, ignore_errors=True)

    yield f"完成: 成功 {converted}，失败 {failed}，跳过 {skipped}，总计 {total}。"
