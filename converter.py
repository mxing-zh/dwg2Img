from __future__ import annotations

from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path
import gc
import multiprocessing as mp
import os
import subprocess
import tempfile
import time
from typing import Iterable, List


SUPPORTED_IMAGE_FORMATS = {"png", "jpg", "jpeg"}
LAYOUT_MODELS = {"auto", "model", "layout"}
COLOR_MODES = {"bw", "original"}


def auto_workers() -> int:
    cpu = os.cpu_count() or 1
    # 保守策略：避免一次拉起过多渲染进程导致窗口/内存抖动
    return max(1, min(4, cpu - 1 if cpu > 1 else 1))


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


def _run_oda_converter_stream(
    input_root: Path,
    dxf_root: Path,
    oda_converter: Path,
) -> Iterable[str]:
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

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )

    if proc.stdout is not None:
        for line in proc.stdout:
            msg = line.strip()
            if msg:
                yield msg

    code = proc.wait()
    if code != 0:
        raise RuntimeError(f"ODA conversion failed with exit code {code}: {' '.join(cmd)}")


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
    import matplotlib

    matplotlib.use("Agg", force=True)
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


def _run_render_tasks_with_recovery(
    tasks: list[tuple[Path, Path, str, int, str, str | None, str]],
    workers: int,
    total: int,
    processed: int,
) -> Iterable[tuple[str, int, int, list[str]]]:
    """Run render tasks with stall recovery.

    Yields events:
      - ("progress", processed, converted, fail_examples)
      - ("failed", processed, converted, fail_examples)
      - ("warn", processed, converted, fail_examples)
    """

    queue = list(tasks)
    converted = 0
    failed = 0
    fail_examples: list[str] = []
    retries: dict[str, int] = {}

    max_retries = 1
    inflight_limit = max(workers * 4, workers)
    stall_timeout_s = 90

    while queue:
        ctx = mp.get_context("spawn")
        inflight: dict[object, tuple[Path, Path, str, int, str, str | None, str]] = {}
        q_idx = 0
        stalled = False

        # NOTE: do not set max_tasks_per_child here; on some Windows/threaded launch paths
        # worker recycle can stall around workers * max_tasks_per_child boundaries (e.g. 4*30=120).
        with ProcessPoolExecutor(max_workers=workers, mp_context=ctx) as pool:
            while q_idx < len(queue) or inflight:
                while q_idx < len(queue) and len(inflight) < inflight_limit:
                    task = queue[q_idx]
                    q_idx += 1
                    fut = pool.submit(_render_worker, task)
                    inflight[fut] = task

                if not inflight:
                    break

                done, _ = wait(inflight.keys(), timeout=stall_timeout_s, return_when=FIRST_COMPLETED)
                if not done:
                    stalled = True
                    break

                for fut in done:
                    task = inflight.pop(fut)
                    try:
                        ok, detail = fut.result()
                    except Exception as exc:
                        ok = False
                        detail = f"{task[0]}: worker crash: {exc}"

                    processed += 1
                    if ok:
                        converted += 1
                        yield ("progress", processed, converted, fail_examples)
                    else:
                        failed += 1
                        if len(fail_examples) < 5:
                            fail_examples.append(detail)
                        yield ("failed", processed, converted, fail_examples)

            if stalled:
                remaining = [inflight[f] for f in inflight]
                remaining.extend(queue[q_idx:])
                queue = []
                for task in remaining:
                    key = str(task[0])
                    retried = retries.get(key, 0)
                    if retried < max_retries:
                        retries[key] = retried + 1
                        queue.append(task)
                    else:
                        processed += 1
                        failed += 1
                        msg = f"{task[0]}: timeout/stall after retry"
                        if len(fail_examples) < 5:
                            fail_examples.append(msg)
                yield ("warn", processed, converted, fail_examples)
                pool.shutdown(wait=False, cancel_futures=True)
            else:
                queue = []

    yield ("done", processed, converted, fail_examples)


def batch_convert(config: ConvertConfig) -> Iterable[str]:
    input_root = config.input_root.resolve()
    if not input_root.exists():
        raise FileNotFoundError(f"Input directory not found: {input_root}")

    layout_mode = config.normalized_layout_mode()
    color_mode = config.normalized_color_mode()
    image_format = config.normalized_format()
    workers = config.normalized_workers()

    dwg_files = discover_dwgs(input_root)
    if not dwg_files:
        yield "未找到 DWG 文件。"
        return

    converter_path = config.oda_converter
    if converter_path is None:
        raise ValueError("必须提供 ODA File Converter 可执行文件路径。")

    total = len(dwg_files)
    yield f"任务开始: 总数 {total}，渲染并发 {workers}，ODA转换 单进程"

    stage_all_start = time.perf_counter()
    stage_oda_start = time.perf_counter()

    with tempfile.TemporaryDirectory(prefix="dwg2img_dxf_") as tmp:
        dxf_root = Path(tmp)
        yield "开始 ODA 单进程转换 DWG -> DXF..."
        for oda_msg in _run_oda_converter_stream(input_root, dxf_root, converter_path):
            if any(token in oda_msg.lower() for token in ("%", "progress", "processing", "converting")):
                yield f"[ODA] {oda_msg}"
        oda_elapsed = time.perf_counter() - stage_oda_start
        yield f"ODA 转换完成，耗时 {oda_elapsed:.1f}s，开始并发渲染图片..."

        tasks: list[tuple[Path, Path, str, int, str, str | None, str]] = []
        skipped = 0
        for dwg_file in dwg_files:
            rel = dwg_file.relative_to(input_root)
            dxf_file = (dxf_root / rel).with_suffix(".dxf")
            if not dxf_file.exists():
                skipped += 1
                continue
            output_file = resolve_output_path(dwg_file, config)
            tasks.append((dxf_file, output_file, image_format, config.dpi, layout_mode, config.preferred_layout, color_mode))

        processed = skipped
        converted = 0
        failed = 0
        fail_examples: list[str] = []

        if skipped:
            yield f"跳过 {skipped} 个文件（未找到对应 DXF）。"

        render_total = len(tasks)
        yield f"渲染阶段开始：待渲染 {render_total}，并发 {workers}"
        stage_render_start = time.perf_counter()

        for event, processed, converted, samples in _run_render_tasks_with_recovery(tasks, workers, total, processed):
            if event == "warn":
                yield "检测到渲染子进程长时间无响应，已自动重试剩余任务。"
            if event in {"progress", "failed", "warn"} and (processed % 5 == 0 or processed == total):
                yield f"进度: {processed}/{total} ({processed / total:.1%})"

        failed = max(total - skipped - converted, 0)
        fail_examples = samples

        render_elapsed = time.perf_counter() - stage_render_start
        yield f"渲染阶段结束：成功 {converted}，失败 {failed}，耗时 {render_elapsed:.1f}s"
        for msg in fail_examples:
            yield f"[失败样例] {msg}"

    all_elapsed = time.perf_counter() - stage_all_start
    yield f"任务结束: 成功 {converted}，失败 {failed}，跳过 {skipped}，总计 {total}，总耗时 {all_elapsed:.1f}s。"
