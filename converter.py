from __future__ import annotations

import copy
import ctypes
from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path
import gc
import math
import multiprocessing as mp
import os
import subprocess
import tempfile
import time
from typing import Iterable, List


SUPPORTED_IMAGE_FORMATS = {"png", "jpg", "jpeg"}
LAYOUT_MODELS = {"auto", "model", "layout"}
COLOR_MODES = {"bw", "original"}
CJK_FONT_CANDIDATES = (
    "simhei.ttf",
    "msyh.ttc",
    "simsun.ttc",
    "NotoSansSC-VF.ttf",
    "NotoSerifSC-VF.ttf",
)
CJK_STYLE_KEYWORDS = ("黑体", "宋体", "仿宋", "楷体", "等线", "中文")
CJK_BIGFONT_MARKERS = ("hztxt", "hz", "ht.shx", "fs.shx", "khz", "gbcbig")
_CJK_FONT_CACHE: str | None | bool = False


def _available_memory_gb() -> float | None:
    if os.name == "nt":
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong),
                ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]

        status = MEMORYSTATUSEX()
        status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
            return status.ullAvailPhys / (1024 ** 3)
        return None

    if hasattr(os, "sysconf"):
        try:
            page_size = os.sysconf("SC_PAGE_SIZE")
            available_pages = os.sysconf("SC_AVPHYS_PAGES")
            return (page_size * available_pages) / (1024 ** 3)
        except (ValueError, OSError, AttributeError):
            return None
    return None


def auto_workers_details() -> tuple[int, str]:
    cpu = os.cpu_count() or 1
    cpu_limit = max(1, cpu - 1 if cpu > 1 else 1)

    available_gb = _available_memory_gb()
    if available_gb is None:
        workers = max(1, min(8, cpu_limit))
        return workers, f"建议 {workers}，按 CPU 核心数估算"

    reserved_gb = 1.0
    per_worker_gb = 0.6
    memory_limit = max(1, int((available_gb - reserved_gb) // per_worker_gb))
    workers = max(1, min(cpu_limit, memory_limit, 12))
    hint = (
        f"建议 {workers}，CPU {cpu} 核，可用内存约 {available_gb:.1f} GB，"
        f"按每个渲染进程约 {per_worker_gb:.1f} GB 估算"
    )
    return workers, hint


def auto_workers() -> int:
    return auto_workers_details()[0]


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
    cluster_gap_scale: float = 8.0
    view_padding_ratio: float = 0.10

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

    def normalized_cluster_gap_scale(self) -> float:
        if self.cluster_gap_scale <= 0:
            raise ValueError("Cluster gap scale must be greater than 0")
        return float(self.cluster_gap_scale)

    def normalized_view_padding_ratio(self) -> float:
        if self.view_padding_ratio < 0:
            raise ValueError("View padding ratio must be greater than or equal to 0")
        return float(self.view_padding_ratio)


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

    def has_substantive_paperspace_entities(layout) -> bool:
        return any(entity.dxftype() != "VIEWPORT" for entity in layout)

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
        if has_substantive_paperspace_entities(layout):
            return layout, layout.name
    if has_renderable_entities(doc.modelspace()):
        return doc.modelspace(), "Model"
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


def _layout_by_name(doc, layout_name: str):
    if layout_name.lower() == "model":
        return doc.modelspace()

    for layout in doc.layouts:
        if layout.name == layout_name:
            return layout
    return doc.modelspace()


def _detect_cjk_font_file() -> str | None:
    global _CJK_FONT_CACHE
    if _CJK_FONT_CACHE is not False:
        return _CJK_FONT_CACHE or None

    try:
        from ezdxf.fonts import fonts
    except Exception:
        _CJK_FONT_CACHE = None
        return None

    for candidate in CJK_FONT_CANDIDATES:
        try:
            if fonts.font_manager.has_font(candidate):
                _CJK_FONT_CACHE = candidate
                return candidate
        except Exception:
            continue

    _CJK_FONT_CACHE = None
    return None


def _needs_cjk_font_fallback(text_style) -> bool:
    style_name = str(getattr(text_style.dxf, "name", "") or "").lower()
    font_name = str(getattr(text_style.dxf, "font", "") or "").lower()
    bigfont_name = str(getattr(text_style.dxf, "bigfont", "") or "").lower()

    if any(keyword in style_name for keyword in CJK_STYLE_KEYWORDS):
        return True
    if any(marker in bigfont_name for marker in CJK_BIGFONT_MARKERS):
        return True
    if bigfont_name and font_name.endswith(".shx"):
        return True
    return False


def _apply_cjk_font_fallbacks(doc) -> None:
    cjk_font = _detect_cjk_font_file()
    if not cjk_font:
        return

    for text_style in doc.styles:
        if _needs_cjk_font_fallback(text_style):
            text_style.dxf.font = cjk_font


def _bbox_to_rect(ext) -> tuple[float, float, float, float] | None:
    if ext is None:
        return None
    return (
        float(ext.extmin.x),
        float(ext.extmin.y),
        float(ext.extmax.x),
        float(ext.extmax.y),
    )


def _rect_size(rect: tuple[float, float, float, float] | None) -> tuple[float, float] | None:
    if rect is None:
        return None

    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    if width <= 0 or height <= 0:
        return None
    return width, height


def _rect_union(rects: list[tuple[float, float, float, float]]) -> tuple[float, float, float, float] | None:
    if not rects:
        return None

    min_x = min(rect[0] for rect in rects)
    min_y = min(rect[1] for rect in rects)
    max_x = max(rect[2] for rect in rects)
    max_y = max(rect[3] for rect in rects)
    return min_x, min_y, max_x, max_y


def _rect_gap(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    dx = max(0.0, a[0] - b[2], b[0] - a[2])
    dy = max(0.0, a[1] - b[3], b[1] - a[3])
    return max(dx, dy)


def _rect_diagonal(rect: tuple[float, float, float, float]) -> float:
    size = _rect_size(rect)
    if size is None:
        return 0.0
    return math.hypot(size[0], size[1])


def _rect_longest_side(rect: tuple[float, float, float, float]) -> float:
    size = _rect_size(rect)
    if size is None:
        return 0.0
    return max(size)


def _safe_entity_rect(entity) -> tuple[float, float, float, float] | None:
    import ezdxf
    from ezdxf import bbox

    try:
        ext = bbox.extents([entity], fast=True)
    except ezdxf.DXFError:
        return None
    return _bbox_to_rect(ext)


def _collect_entity_rects(layout) -> list[tuple[str, tuple[float, float, float, float]]]:
    rects: list[tuple[str, tuple[float, float, float, float]]] = []
    for entity in layout:
        rect = _safe_entity_rect(entity)
        if rect is None:
            continue
        rects.append((str(entity.dxf.handle), rect))
    return rects


def _cluster_entity_rects(
    entity_rects: list[tuple[str, tuple[float, float, float, float]]],
    gap_scale: float,
) -> list[dict[str, object]]:
    clusters: list[dict[str, object]] = []
    pending = set(range(len(entity_rects)))

    while pending:
        seed = max(pending, key=lambda idx: _rect_longest_side(entity_rects[idx][1]))
        pending.remove(seed)

        seed_handle, seed_rect = entity_rects[seed]
        handles: list[str] = []
        rects: list[tuple[float, float, float, float]] = []
        cluster_rect = seed_rect
        handles.append(seed_handle)
        rects.append(seed_rect)

        changed = True
        while changed:
            changed = False
            dynamic_gap = max(_rect_longest_side(cluster_rect), _rect_longest_side(seed_rect), 1.0) * gap_scale
            for other_idx in list(pending):
                handle, other_rect = entity_rects[other_idx]
                if _rect_gap(cluster_rect, other_rect) <= dynamic_gap:
                    pending.remove(other_idx)
                    handles.append(handle)
                    rects.append(other_rect)
                    cluster_rect = _rect_union(rects) or cluster_rect
                    changed = True

        bbox_rect = _rect_union(rects)
        if bbox_rect is None:
            continue

        clusters.append(
            {
                "handles": handles,
                "bbox": bbox_rect,
                "count": len(handles),
                "score": max(_rect_diagonal(bbox_rect), 1.0) * len(handles),
            }
        )

    clusters.sort(key=lambda item: (float(item["score"]), int(item["count"])), reverse=True)
    return clusters


def _pick_focus_clusters(clusters: list[dict[str, object]], cluster_gap_scale: float) -> list[dict[str, object]]:
    if len(clusters) <= 1:
        return clusters

    best = clusters[0]
    best_count = int(best["count"])
    best_score = float(best["score"])
    best_bbox = best["bbox"]
    proximity_limit = max(_rect_longest_side(best_bbox), 1.0) * max(cluster_gap_scale, 1.0)

    kept = [best]

    for cluster in clusters[1:]:
        cluster_count = int(cluster["count"])
        cluster_score = float(cluster["score"])
        gap_to_best = _rect_gap(best_bbox, cluster["bbox"])
        is_significant = cluster_count >= max(2, int(best_count * 0.15)) or cluster_score >= best_score * 0.1
        if gap_to_best <= proximity_limit and is_significant:
            kept.append(cluster)

    return kept


def _prepare_render_layout(doc, layout_name: str, cluster_gap_scale: float):
    source_layout = _layout_by_name(doc, layout_name)
    entity_rects = _collect_entity_rects(source_layout)
    if not entity_rects:
        return doc, source_layout, None

    clusters = _cluster_entity_rects(entity_rects, cluster_gap_scale)
    focus_clusters = _pick_focus_clusters(clusters, cluster_gap_scale)
    focus_rect = _rect_union([cluster["bbox"] for cluster in focus_clusters if cluster.get("bbox") is not None])
    if focus_rect is None:
        return doc, source_layout, None

    kept_handles = {
        handle
        for cluster in focus_clusters
        for handle in cluster["handles"]
    }
    max_abs_coord = max(abs(value) for value in focus_rect)
    needs_filtering = len(kept_handles) < len(entity_rects)
    needs_translation = max_abs_coord > 1_000_000

    if not needs_filtering and not needs_translation:
        return doc, source_layout, focus_rect

    render_doc = copy.deepcopy(doc)
    render_layout = _layout_by_name(render_doc, layout_name)

    if needs_filtering:
        for entity in list(render_layout):
            if str(entity.dxf.handle) not in kept_handles:
                render_layout.delete_entity(entity)

    if needs_translation:
        shift_x = -focus_rect[0]
        shift_y = -focus_rect[1]
        for entity in render_layout:
            try:
                entity.translate(shift_x, shift_y, 0)
            except Exception:
                continue
        width = focus_rect[2] - focus_rect[0]
        height = focus_rect[3] - focus_rect[1]
        focus_rect = (0.0, 0.0, width, height)

    return render_doc, render_layout, focus_rect


def _figure_size_inches(layout, focus_rect: tuple[float, float, float, float] | None = None) -> tuple[float, float]:
    paper_size = _layout_paper_inches(layout)
    if paper_size:
        return paper_size

    bbox_size = _rect_size(focus_rect) or _safe_bbox_size(layout)
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
    cluster_gap_scale: float,
    view_padding_ratio: float,
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
    render_doc, render_layout, focus_rect = _prepare_render_layout(doc, layout_name, cluster_gap_scale)
    _apply_cjk_font_fallbacks(render_doc)

    fig_w, fig_h = _figure_size_inches(render_layout, focus_rect)
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_aspect("equal", adjustable="box")
    ax.axis("off")

    ctx = RenderContext(render_doc)
    out = MatplotlibBackend(ax)
    color_policy = ColorPolicy.BLACK if color_mode == "bw" else ColorPolicy.COLOR
    draw_config = Configuration.defaults().with_changes(
        color_policy=color_policy,
        background_policy=BackgroundPolicy.WHITE,
        proxy_graphic_policy=ProxyGraphicPolicy.PREFER,
    )
    Frontend(ctx, out, config=draw_config).draw_layout(render_layout, finalize=True)

    view_rect = focus_rect or _bbox_to_rect(_safe_bbox_extents(render_layout))
    view_size = _rect_size(view_rect)
    if view_rect is not None and view_size is not None:
        pad_x = view_size[0] * view_padding_ratio
        pad_y = view_size[1] * view_padding_ratio
        ax.set_xlim(view_rect[0] - pad_x, view_rect[2] + pad_x)
        ax.set_ylim(view_rect[1] - pad_y, view_rect[3] + pad_y)
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

    del out, ctx, target_layout, render_doc, doc
    gc.collect()
    return layout_name, final_size


def _render_worker(task: tuple[Path, Path, str, int, str, str | None, str, float, float]) -> tuple[bool, str]:
    dxf_file, output_file, fmt, dpi, layout_mode, preferred_layout, color_mode, cluster_gap_scale, view_padding_ratio = task
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
            cluster_gap_scale=cluster_gap_scale,
            view_padding_ratio=view_padding_ratio,
        )
        return True, f"{dxf_file.name} -> {output_file.name} ({layout_name}, {size_px[0]}x{size_px[1]}px)"
    except Exception as exc:
        return False, f"{dxf_file}: {exc}"


def _run_render_tasks_with_recovery(
    tasks: list[tuple[Path, Path, str, int, str, str | None, str, float, float]],
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
        inflight: dict[object, tuple[Path, Path, str, int, str, str | None, str, float, float]] = {}
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
    cluster_gap_scale = config.normalized_cluster_gap_scale()
    view_padding_ratio = config.normalized_view_padding_ratio()

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

        tasks: list[tuple[Path, Path, str, int, str, str | None, str, float, float]] = []
        skipped = 0
        for dwg_file in dwg_files:
            rel = dwg_file.relative_to(input_root)
            dxf_file = (dxf_root / rel).with_suffix(".dxf")
            if not dxf_file.exists():
                skipped += 1
                continue
            output_file = resolve_output_path(dwg_file, config)
            tasks.append(
                (
                    dxf_file,
                    output_file,
                    image_format,
                    config.dpi,
                    layout_mode,
                    config.preferred_layout,
                    color_mode,
                    cluster_gap_scale,
                    view_padding_ratio,
                )
            )

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
            if event in {"progress", "failed", "warn"} and (processed % 200 == 0 or processed == total):
                yield f"进度: {processed}/{total} ({processed / total:.1%})"

        failed = max(total - skipped - converted, 0)
        fail_examples = samples

        render_elapsed = time.perf_counter() - stage_render_start
        yield f"渲染阶段结束：成功 {converted}，失败 {failed}，耗时 {render_elapsed:.1f}s"
        for msg in fail_examples:
            yield f"[失败样例] {msg}"

    all_elapsed = time.perf_counter() - stage_all_start
    yield f"任务结束: 成功 {converted}，失败 {failed}，跳过 {skipped}，总计 {total}，总耗时 {all_elapsed:.1f}s。"
