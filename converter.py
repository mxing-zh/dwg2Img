from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import subprocess
import tempfile
from typing import Iterable, List


SUPPORTED_IMAGE_FORMATS = {"png", "jpg", "jpeg"}
LAYOUT_MODELS = {"auto", "model", "layout"}


@dataclass
class ConvertConfig:
    input_root: Path
    output_root: Path
    image_format: str = "png"
    dpi: int = 96
    mirror_structure: bool = True
    single_output_dir: Path | None = None
    oda_converter: Path | None = None
    layout_mode: str = "auto"
    preferred_layout: str | None = None

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


def discover_dwgs(input_root: Path) -> List[Path]:
    return sorted(p for p in input_root.rglob("*.dwg") if p.is_file())


def resolve_output_path(dwg_file: Path, config: ConvertConfig) -> Path:
    fmt = config.normalized_format()
    if config.single_output_dir:
        config.single_output_dir.mkdir(parents=True, exist_ok=True)
        return config.single_output_dir / f"{dwg_file.stem}.{fmt}"

    rel = dwg_file.relative_to(config.input_root)
    destination = config.output_root / rel
    destination.parent.mkdir(parents=True, exist_ok=True)
    return destination.with_suffix(f".{fmt}")


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

    if mode == "layout":
        if preferred:
            for layout in paperspaces:
                if layout.name.lower() == preferred:
                    return layout, layout.name
        for layout in paperspaces:
            if len(layout):
                return layout, layout.name
        return doc.modelspace(), "Model"

    if preferred:
        for layout in paperspaces:
            if layout.name.lower() == preferred:
                return layout, layout.name

    for layout in paperspaces:
        if len(layout):
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


def _safe_bbox_size(layout) -> tuple[float, float] | None:
    import ezdxf
    from ezdxf import bbox

    try:
        ext = bbox.extents(layout, fast=True)
    except ezdxf.DXFError:
        return None
    if ext is None:
        return None

    width = float(ext.size.x)
    height = float(ext.size.y)
    if width <= 0 or height <= 0:
        return None
    return width, height


def _figure_size_inches(layout, dpi: int) -> tuple[float, float]:
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


def _normalize_image_output(image_file: Path, image_format: str, dpi: int) -> None:
    from PIL import Image

    fmt = image_format.upper()
    if fmt == "JPG":
        fmt = "JPEG"

    with Image.open(image_file) as im:
        rgb = im.convert("RGB")
        if fmt == "JPEG":
            rgb.save(image_file, format=fmt, quality=95, subsampling=0, dpi=(dpi, dpi))
        else:
            rgb.save(image_file, format=fmt, dpi=(dpi, dpi))


def _render_dxf_to_image(
    dxf_file: Path,
    image_file: Path,
    image_format: str,
    dpi: int,
    layout_mode: str,
    preferred_layout: str | None,
) -> tuple[str, tuple[int, int]]:
    import matplotlib.pyplot as plt
    import ezdxf
    from ezdxf.addons.drawing import Frontend, RenderContext
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend

    doc = ezdxf.readfile(dxf_file)
    target_layout, layout_name = _pick_layout(doc, layout_mode, preferred_layout)

    fig_w, fig_h = _figure_size_inches(target_layout, dpi)
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")

    ctx = RenderContext(doc)
    out = MatplotlibBackend(ax)
    Frontend(ctx, out).draw_layout(target_layout, finalize=True)
    fig.savefig(image_file, dpi=dpi, facecolor="white", transparent=False)
    width_px = int(round(fig_w * dpi))
    height_px = int(round(fig_h * dpi))
    plt.close(fig)

    _normalize_image_output(image_file, image_format, dpi)
    return layout_name, (width_px, height_px)


def batch_convert(config: ConvertConfig) -> Iterable[str]:
    input_root = config.input_root.resolve()
    if not input_root.exists():
        raise FileNotFoundError(f"Input directory not found: {input_root}")

    layout_mode = config.normalized_layout_mode()

    dwg_files = discover_dwgs(input_root)
    if not dwg_files:
        yield "未找到 DWG 文件。"
        return

    yield f"共找到 {len(dwg_files)} 个 DWG 文件。"

    converter_path = config.oda_converter
    if converter_path is None:
        raise ValueError("必须提供 ODA File Converter 可执行文件路径。")

    with tempfile.TemporaryDirectory(prefix="dwg2img_dxf_") as tmp:
        dxf_root = Path(tmp)
        yield "开始将 DWG 批量转换为 DXF..."
        _run_oda_converter(input_root, dxf_root, converter_path)
        yield "DWG -> DXF 完成，开始渲染图片..."

        converted = 0
        for index, dwg_file in enumerate(dwg_files, 1):
            rel = dwg_file.relative_to(input_root)
            dxf_file = (dxf_root / rel).with_suffix(".dxf")
            if not dxf_file.exists():
                yield f"[{index}/{len(dwg_files)}] 跳过（未找到DXF）: {rel}"
                continue

            output_file = resolve_output_path(dwg_file, config)
            layout_name, size_px = _render_dxf_to_image(
                dxf_file=dxf_file,
                image_file=output_file,
                image_format=config.normalized_format(),
                dpi=config.dpi,
                layout_mode=layout_mode,
                preferred_layout=config.preferred_layout,
            )
            converted += 1
            yield (
                f"[{index}/{len(dwg_files)}] 完成({layout_name}, {size_px[0]}x{size_px[1]}px, "
                f"{config.dpi}dpi, 24-bit): {rel} -> {output_file}"
            )

    yield f"完成，成功输出 {converted}/{len(dwg_files)} 张图片。"
