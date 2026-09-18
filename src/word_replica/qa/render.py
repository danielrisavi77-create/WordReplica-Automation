from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image, ImageChops, ImageFilter


ANTIALIASING_BLUR_RADIUS = 1.0
LEGACY_ANTIALIASING_CHANGED_PIXEL_ALLOWANCE = 0.03
LEGACY_ANTIALIASING_MAE_ALLOWANCE = 1.0
BLURRED_ANTIALIASING_CHANGED_PIXEL_ALLOWANCE = 0.04
BLURRED_ANTIALIASING_MAE_ALLOWANCE = 1.0


@dataclass(slots=True)
class VisualMetric:
    same_dimensions: bool
    changed_pixel_ratio: float
    mean_absolute_error: float
    source_size: tuple[int, int]
    rebuilt_size: tuple[int, int]
    blurred_mean_absolute_error: float | None = None


@dataclass(slots=True)
class RenderQaResult:
    available: bool
    within_tolerance: bool
    page_count_match: bool
    source_page_count: int
    rebuilt_page_count: int
    metrics: list[VisualMetric] = field(default_factory=list)
    diff_images: list[Path] = field(default_factory=list)
    message: str | None = None


def visual_metric_acceptance_mode(
    metric: VisualMetric,
    *,
    changed_pixel_tolerance: float,
    mae_tolerance: float,
) -> str | None:
    if not metric.same_dimensions:
        return None
    if (
        metric.changed_pixel_ratio <= changed_pixel_tolerance
        and metric.mean_absolute_error <= mae_tolerance
    ):
        return "strict"
    if (
        metric.changed_pixel_ratio
        <= max(changed_pixel_tolerance, LEGACY_ANTIALIASING_CHANGED_PIXEL_ALLOWANCE)
        and metric.mean_absolute_error
        <= max(mae_tolerance, LEGACY_ANTIALIASING_MAE_ALLOWANCE)
    ):
        return "legacy_antialiasing"
    if (
        metric.blurred_mean_absolute_error is not None
        and metric.changed_pixel_ratio
        <= max(changed_pixel_tolerance, BLURRED_ANTIALIASING_CHANGED_PIXEL_ALLOWANCE)
        and metric.blurred_mean_absolute_error
        <= max(mae_tolerance, BLURRED_ANTIALIASING_MAE_ALLOWANCE)
    ):
        return "blurred_antialiasing"
    return None


def rasterize_pdf(pdf_path: Path, out_dir: Path, dpi: int = 144) -> list[Path]:
    import pypdfium2 as pdfium

    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = pdfium.PdfDocument(str(pdf_path))
    scale = dpi / 72.0
    pages: list[Path] = []
    try:
        for index in range(len(pdf)):
            image = pdf[index].render(scale=scale).to_pil().convert("RGB")
            path = out_dir / f"page_{index + 1:04d}.png"
            image.save(path)
            pages.append(path)
    finally:
        pdf.close()
    return pages


def compare_page_images(a: Path, b: Path) -> VisualMetric:
    with Image.open(a) as source_image, Image.open(b) as rebuilt_image:
        ia = source_image.convert("RGB")
        ib = rebuilt_image.convert("RGB")
        if ia.size != ib.size:
            return VisualMetric(False, 1.0, 255.0, ia.size, ib.size)
        diff = ImageChops.difference(ia, ib)
        hist = diff.histogram()
        pixels = ia.width * ia.height
        changed_channels = sum(
            count
            for channel in range(3)
            for value in range(1, 256)
            for count in (hist[channel * 256 + value],)
        )
        changed_ratio = min(1.0, changed_channels / max(1, pixels * 3))
        mae = sum(
            value * hist[channel * 256 + value]
            for channel in range(3)
            for value in range(256)
        ) / max(1, pixels * 3)
        blurred_diff = ImageChops.difference(
            ia.filter(ImageFilter.GaussianBlur(radius=ANTIALIASING_BLUR_RADIUS)),
            ib.filter(ImageFilter.GaussianBlur(radius=ANTIALIASING_BLUR_RADIUS)),
        )
        blurred_hist = blurred_diff.histogram()
        blurred_mae = sum(
            value * blurred_hist[channel * 256 + value]
            for channel in range(3)
            for value in range(256)
        ) / max(1, pixels * 3)
        return VisualMetric(
            True,
            changed_ratio,
            mae,
            ia.size,
            ib.size,
            blurred_mean_absolute_error=blurred_mae,
        )


def _write_diff_image(a: Path, b: Path, destination: Path) -> Path | None:
    with Image.open(a) as source_image, Image.open(b) as rebuilt_image:
        ia = source_image.convert("RGB")
        ib = rebuilt_image.convert("RGB")
        if ia.size != ib.size:
            return None
        destination.parent.mkdir(parents=True, exist_ok=True)
        ImageChops.difference(ia, ib).save(destination)
        return destination


def compare_pdfs(
    source_pdf: Path,
    rebuilt_pdf: Path,
    qa_dir: Path,
    dpi: int = 144,
    *,
    changed_pixel_tolerance: float = 0.001,
    mae_tolerance: float = 0.25,
) -> RenderQaResult:
    source_dir = qa_dir / "l4_source_pages"
    rebuilt_dir = qa_dir / "l4_rebuilt_pages"
    diff_dir = qa_dir / "l4_diffs"
    source_pages = rasterize_pdf(source_pdf, source_dir, dpi=dpi)
    rebuilt_pages = rasterize_pdf(rebuilt_pdf, rebuilt_dir, dpi=dpi)
    page_count_match = len(source_pages) == len(rebuilt_pages)
    metrics: list[VisualMetric] = []
    diff_images: list[Path] = []
    for index, (source_page, rebuilt_page) in enumerate(zip(source_pages, rebuilt_pages), start=1):
        metric = compare_page_images(source_page, rebuilt_page)
        metrics.append(metric)
        if metric.changed_pixel_ratio > 0 or not metric.same_dimensions:
            diff_path = _write_diff_image(source_page, rebuilt_page, diff_dir / f"page_{index:04d}_diff.png")
            if diff_path is not None:
                diff_images.append(diff_path)
    within_tolerance = page_count_match and all(
        visual_metric_acceptance_mode(
            metric,
            changed_pixel_tolerance=changed_pixel_tolerance,
            mae_tolerance=mae_tolerance,
        ) is not None for metric in metrics
    )
    return RenderQaResult(
        available=True,
        within_tolerance=within_tolerance,
        page_count_match=page_count_match,
        source_page_count=len(source_pages),
        rebuilt_page_count=len(rebuilt_pages),
        metrics=metrics,
        diff_images=diff_images,
    )
