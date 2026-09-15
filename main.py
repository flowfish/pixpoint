from __future__ import annotations

import json
import math
import sys
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from time import perf_counter
from typing import Callable, Optional

from PIL import Image, ImageOps
from PyQt5.QtCore import QPoint, QPointF, QRect, QRectF, QSize, Qt, pyqtSignal
from PyQt5.QtGui import (
    QBrush,
    QColor,
    QFont,
    QIcon,
    QImage,
    QImageReader,
    QKeySequence,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
)
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAction,
    QApplication,
    QColorDialog,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGraphicsItem,
    QGraphicsPixmapItem,
    QGraphicsScene,
    QGraphicsView,
    QHBoxLayout,
    QInputDialog,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressDialog,
    QShortcut,
    QSpinBox,
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


if hasattr(QImageReader, "setAllocationLimit"):
    QImageReader.setAllocationLimit(1024)


MAX_EXCEL_EXPORT_DIMENSION = 2560
MAX_EXCEL_EXPORT_PIXELS = 4_194_304
DEFAULT_EXCEL_EXPORT_WIDTH = 100
DEFAULT_EXCEL_CELL_SIZE = 6
EXCEL_FORMAT_COLOR_LIMIT = 32_768
EXCEL_COLOR_BUCKET_STEP = 8
LARGE_EXCEL_EXPORT_WARNING_PIXELS = 500_000
RECOMMENDED_EXCEL_EXPORT_PIXELS = 1_600_000
INVALID_WORKSHEET_CHARS = set("[]:*?/\\")


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


def _resample_filter() -> int:
    if hasattr(Image, "Resampling"):
        return Image.Resampling.LANCZOS
    return Image.LANCZOS


def compute_scaled_height(source_width: int, source_height: int, target_width: int) -> int:
    if source_width <= 0 or source_height <= 0 or target_width <= 0:
        raise ValueError("Image dimensions and target width must be positive.")
    return max(1, round(source_height * target_width / source_width))


def compute_max_excel_width(
    source_width: int,
    source_height: int,
    max_dimension: int = MAX_EXCEL_EXPORT_DIMENSION,
    max_pixels: int = MAX_EXCEL_EXPORT_PIXELS,
) -> int:
    valid_width = 0
    for width in range(1, max_dimension + 1):
        height = compute_scaled_height(source_width, source_height, width)
        if height <= max_dimension and width * height <= max_pixels:
            valid_width = width
    return valid_width


def compute_recommended_excel_width(
    source_width: int,
    source_height: int,
    max_dimension: int = MAX_EXCEL_EXPORT_DIMENSION,
    max_pixels: int = MAX_EXCEL_EXPORT_PIXELS,
    recommended_pixels: int = RECOMMENDED_EXCEL_EXPORT_PIXELS,
) -> int:
    max_width = compute_max_excel_width(source_width, source_height, max_dimension, max_pixels)
    if max_width < 1:
        return 0

    original_pixels = source_width * source_height
    if source_width <= max_width and original_pixels <= recommended_pixels:
        return source_width

    aspect_ratio = source_width / source_height
    suggested_width = max(1, int(round(math.sqrt(recommended_pixels * aspect_ratio))))
    suggested_width = min(suggested_width, source_width, max_width)

    while suggested_width > 1:
        suggested_height = compute_scaled_height(source_width, source_height, suggested_width)
        if suggested_height <= max_dimension and suggested_width * suggested_height <= max_pixels:
            return suggested_width
        suggested_width -= 1
    return 1


def sanitize_worksheet_name(name: str, fallback: str = "") -> str:
    cleaned = "".join("_" if char in INVALID_WORKSHEET_CHARS else char for char in name).strip()
    cleaned = cleaned.strip("'")
    if not cleaned:
        cleaned = fallback.strip() if fallback else datetime.now().strftime("%Y%m%d")
    cleaned = cleaned.strip("'")
    return cleaned[:31] or datetime.now().strftime("%Y%m%d")


def unique_color_count(image: Image.Image, max_colors: int) -> Optional[int]:
    colors = image.getcolors(max_colors + 1)
    if colors is None:
        return None
    return len(colors)


def reduce_excel_color_space(image: Image.Image, bucket_step: int = EXCEL_COLOR_BUCKET_STEP) -> Image.Image:
    return image.point(lambda value: min(255, (value // bucket_step) * bucket_step)).convert("RGB")


def format_duration(seconds: float) -> str:
    seconds = max(0, int(round(seconds)))
    minutes, secs = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours:d}h {minutes:02d}m {secs:02d}s"
    if minutes:
        return f"{minutes:d}m {secs:02d}s"
    return f"{secs:d}s"


def export_image_to_excel(
    image_path: Path,
    output_path: Path,
    target_width: int,
    worksheet_name: str,
    cell_size_pixels: int = DEFAULT_EXCEL_CELL_SIZE,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> tuple[int, int]:
    if xlsxwriter is None:
        raise RuntimeError("XlsxWriter is not installed.")

    image_path = Path(image_path)
    output_path = Path(output_path)
    if not image_path.exists():
        raise FileNotFoundError(image_path)

    with Image.open(image_path) as source_image:
        source_image = ImageOps.exif_transpose(source_image).convert("RGB")
        source_width, source_height = source_image.size
        if source_width <= 0 or source_height <= 0:
            raise ValueError("Invalid image size.")

        target_height = compute_scaled_height(source_width, source_height, target_width)
        if (
            target_width > MAX_EXCEL_EXPORT_DIMENSION
            or target_height > MAX_EXCEL_EXPORT_DIMENSION
            or target_width * target_height > MAX_EXCEL_EXPORT_PIXELS
        ):
            raise ValueError(
                "Target size exceeds the Excel export limit "
                f"({MAX_EXCEL_EXPORT_DIMENSION}x{MAX_EXCEL_EXPORT_DIMENSION}, "
                f"max {MAX_EXCEL_EXPORT_PIXELS} pixels)."
            )

        resized = source_image.resize((target_width, target_height), _resample_filter())
        color_count = unique_color_count(resized, EXCEL_FORMAT_COLOR_LIMIT)
        if color_count is None or color_count > EXCEL_FORMAT_COLOR_LIMIT:
            resized = reduce_excel_color_space(resized)
        pixels = resized.load()

    workbook = xlsxwriter.Workbook(str(output_path), {"constant_memory": True})
    try:
        worksheet = workbook.add_worksheet(sanitize_worksheet_name(worksheet_name, image_path.stem))
        worksheet.hide_gridlines(2)
        worksheet.set_zoom(20)
        worksheet.set_column_pixels(0, target_width - 1, cell_size_pixels)

        format_cache: dict[str, object] = {}
        for row in range(target_height):
            worksheet.set_row_pixels(row, cell_size_pixels)
            for col in range(target_width):
                red, green, blue = pixels[col, row]
                color = f"#{red:02X}{green:02X}{blue:02X}"
                cell_format = format_cache.get(color)
                if cell_format is None:
                    cell_format = workbook.add_format({"bg_color": color, "pattern": 1})
                    format_cache[color] = cell_format
                worksheet.write_blank(row, col, None, cell_format)
            if progress_callback is not None and (row % 8 == 0 or row + 1 == target_height):
                progress_callback(row + 1, target_height)
    except Exception:
        workbook.close()
        if output_path.exists():
            output_path.unlink(missing_ok=True)
        raise

    workbook.close()
    return target_width, target_height


@dataclass
class PointRecord:
    index: int
    x: int
    y: int
    color: QColor
    marker: "PointMarkerItem"


class PointMarkerItem(QGraphicsItem):
    def __init__(self, index: int, color: QColor) -> None:
        super().__init__()
        self._index = index
        self._color = QColor(color)
        self._highlighted = False
        self.setFlag(QGraphicsItem.ItemIgnoresTransformations, True)
        self.setAcceptedMouseButtons(Qt.NoButton)
        self.setZValue(1000)

    def boundingRect(self) -> QRectF:
        return QRectF(-24, -24, 48, 48)

    def set_index(self, index: int) -> None:
        self._index = index
        self.update()

    def set_color(self, color: QColor) -> None:
        self._color = QColor(color)
        self.update()

    def set_highlighted(self, highlighted: bool) -> None:
        self._highlighted = highlighted
        self.update()

    def paint(self, painter: QPainter, option, widget=None) -> None:
        painter.setRenderHint(QPainter.Antialiasing, True)

        fill = QColor(self._color)
        fill.setAlpha(190 if self._highlighted else 150)
        ring_pen = QPen(QColor("#ffffff") if self._highlighted else QColor("#111111"))
        ring_pen.setWidth(2 if self._highlighted else 1)

        painter.setPen(ring_pen)
        painter.setBrush(fill)
        painter.drawEllipse(QRectF(-11, -11, 22, 22))

        # Punch a small transparent cross through the marker so the underlying
        # image remains readable at the exact pick position.
        painter.save()
        painter.setCompositionMode(QPainter.CompositionMode_Clear)
        clear_pen = QPen(Qt.transparent, 2, Qt.SolidLine, Qt.RoundCap)
        painter.setPen(clear_pen)
        painter.drawLine(QPointF(-4, 0), QPointF(4, 0))
        painter.drawLine(QPointF(0, -4), QPointF(0, 4))
        painter.restore()

        tag_rect = QRectF(10, -21, 22, 16)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(20, 20, 20, 220))
        painter.drawRoundedRect(tag_rect, 4, 4)

        painter.setPen(QPen(QColor("#f5f5f5"), 1))
        painter.setFont(QFont("Segoe UI", 8, QFont.Bold))
        painter.drawText(tag_rect, Qt.AlignCenter, str(self._index))


class MagnifierWidget(QWidget):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._image: Optional[QImage] = None
        self._point: Optional[QPoint] = None
        self._zoom_factor = 10
        self.setMinimumSize(220, 220)
        self.setStyleSheet("background:#1a1a1a; border:1px solid #444;")

    def set_image(self, image: Optional[QImage]) -> None:
        self._image = image
        self.update()

    def set_point(self, point: Optional[QPoint]) -> None:
        self._point = point
        self.update()

    def set_zoom_factor(self, zoom_factor: int) -> None:
        self._zoom_factor = max(2, zoom_factor)
        self.update()

    def zoom_factor(self) -> int:
        return self._zoom_factor

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#1a1a1a"))

        if self._image is None or self._image.isNull() or self._point is None:
            painter.setPen(QColor("#aaaaaa"))
            painter.drawText(self.rect(), Qt.AlignCenter, "Magnifier")
            return

        width = max(3, self.width() // self._zoom_factor)
        height = max(3, self.height() // self._zoom_factor)
        width = min(width, self._image.width())
        height = min(height, self._image.height())

        half_w = width // 2
        half_h = height // 2
        x0 = self._point.x() - half_w
        y0 = self._point.y() - half_h
        x0 = max(0, min(x0, self._image.width() - width))
        y0 = max(0, min(y0, self._image.height() - height))

        cropped = self._image.copy(QRect(x0, y0, width, height))
        scaled = cropped.scaled(self.size(), Qt.IgnoreAspectRatio, Qt.FastTransformation)
        painter.drawImage(self.rect(), scaled)

        painter.setRenderHint(QPainter.Antialiasing, False)
        painter.setPen(QPen(QColor(255, 255, 255, 60), 1))
        for x in range(0, self.width(), self._zoom_factor):
            painter.drawLine(x, 0, x, self.height())
        for y in range(0, self.height(), self._zoom_factor):
            painter.drawLine(0, y, self.width(), y)

        painter.setRenderHint(QPainter.Antialiasing, True)
        center = self.rect().center()
        painter.setPen(QPen(QColor("#ff4d4f"), 2))
        painter.drawLine(center.x() - 8, center.y(), center.x() + 8, center.y())
        painter.drawLine(center.x(), center.y() - 8, center.x(), center.y() + 8)
        painter.setBrush(QColor("#ff4d4f"))
        painter.drawEllipse(center, 3, 3)

        label_rect = QRect(8, 8, self.width() - 16, 22)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(0, 0, 0, 160))
        painter.drawRoundedRect(label_rect, 4, 4)
        painter.setPen(QColor("#ffffff"))
        painter.drawText(
            label_rect,
            Qt.AlignCenter,
            f"({self._point.x()}, {self._point.y()})  {self._zoom_factor}x",
        )


class OverviewWidget(QWidget):
    requested_center = pyqtSignal(QPointF)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._image: Optional[QImage] = None
        self._viewport_rect = QRectF()
        self._display_rect = QRect()
        self._source_pixmap = QPixmap()
        self._cached_pixmap = QPixmap()
        self._cache_size = QSize()
        self.setMinimumHeight(180)
        self.setStyleSheet("background:#161616; border:1px solid #444;")

    def set_image(self, image: Optional[QImage]) -> None:
        self._image = image
        self._source_pixmap = QPixmap.fromImage(image) if image is not None else QPixmap()
        self._cached_pixmap = QPixmap()
        self._cache_size = QSize()
        self.update()

    def set_viewport_rect(self, rect: QRectF) -> None:
        self._viewport_rect = QRectF(rect)
        self.update()

    def resizeEvent(self, event) -> None:
        self._cached_pixmap = QPixmap()
        self._cache_size = QSize()
        super().resizeEvent(event)

    def mousePressEvent(self, event) -> None:
        if self._image is None or self._display_rect.isNull():
            return
        if not self._display_rect.contains(event.pos()):
            return

        nx = (event.x() - self._display_rect.left()) / max(1, self._display_rect.width())
        ny = (event.y() - self._display_rect.top()) / max(1, self._display_rect.height())
        self.requested_center.emit(
            QPointF(nx * self._image.width(), ny * self._image.height())
        )

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#161616"))

        if self._image is None or self._image.isNull():
            painter.setPen(QColor("#aaaaaa"))
            painter.drawText(self.rect(), Qt.AlignCenter, "Overview")
            return

        inner = self.rect().adjusted(8, 8, -8, -8)
        target = self._source_pixmap.size()
        target.scale(inner.size(), Qt.KeepAspectRatio)
        left = inner.left() + (inner.width() - target.width()) // 2
        top = inner.top() + (inner.height() - target.height()) // 2
        self._display_rect = QRect(left, top, target.width(), target.height())

        if self._cache_size != target or self._cached_pixmap.isNull():
            self._cached_pixmap = self._source_pixmap.scaled(
                target, Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            self._cache_size = target

        painter.drawPixmap(self._display_rect.topLeft(), self._cached_pixmap)

        if not self._viewport_rect.isNull():
            sx = self._display_rect.width() / self._image.width()
            sy = self._display_rect.height() / self._image.height()
            vp = QRectF(
                self._display_rect.left() + self._viewport_rect.left() * sx,
                self._display_rect.top() + self._viewport_rect.top() * sy,
                self._viewport_rect.width() * sx,
                self._viewport_rect.height() * sy,
            )
            painter.setBrush(Qt.NoBrush)
            painter.setPen(QPen(QColor("#ff4d4f"), 2))
            painter.drawRect(vp)


class PointEditDialog(QDialog):
    def __init__(
        self,
        x: int,
        y: int,
        color: QColor,
        max_width: int,
        max_height: int,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._color = QColor(color)
        self.setWindowTitle("Edit Point")

        self.x_spin = QSpinBox(self)
        self.x_spin.setRange(0, max_width)
        self.x_spin.setValue(x)

        self.y_spin = QSpinBox(self)
        self.y_spin.setRange(0, max_height)
        self.y_spin.setValue(y)

        self.color_button = QPushButton(self._color.name().upper(), self)
        self.color_button.clicked.connect(self._choose_color)
        self._apply_color_button()

        layout = QFormLayout(self)
        layout.addRow("X", self.x_spin)
        layout.addRow("Y", self.y_spin)
        layout.addRow("Color", self.color_button)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def _apply_color_button(self) -> None:
        self.color_button.setText(self._color.name().upper())
        self.color_button.setStyleSheet(
            f"background:{self._color.name()}; color:#ffffff; padding:4px 8px;"
        )

    def _choose_color(self) -> None:
        color = QColorDialog.getColor(self._color, self, "Point Color")
        if color.isValid():
            self._color = color
            self._apply_color_button()

    def values(self) -> tuple[int, int, QColor]:
        return self.x_spin.value(), self.y_spin.value(), QColor(self._color)


class PixelToExcelDialog(QDialog):
    def __init__(
        self,
        image_width: int,
        image_height: int,
        sheet_name: str,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._image_width = image_width
        self._image_height = image_height
        self._max_width = compute_max_excel_width(image_width, image_height)
        self._suggested_width = compute_recommended_excel_width(image_width, image_height)
        if self._max_width < 1:
            raise ValueError("The current image cannot fit within the Excel export limit.")

        self.setWindowTitle("Export Pixel to Excel")

        self.width_spin = QSpinBox(self)
        self.width_spin.setRange(1, self._max_width)
        self.width_spin.setValue(self._suggested_width or min(DEFAULT_EXCEL_EXPORT_WIDTH, self._max_width))
        self.width_spin.setSuffix(" px")

        self.height_value = QLabel(self)
        self.pixel_value = QLabel(self)
        self.width_hint = QLabel(
            f"系统建议值：{self.width_spin.value()}。你可以直接修改这个值。\n"
            "这里的宽度表示导出后图片宽度，也就是 Excel 里大约会占用多少列；"
            "高度会按原图比例自动计算。",
            self,
        )
        self.limit_hint = QLabel(
            f"最大支持：{MAX_EXCEL_EXPORT_DIMENSION} x {MAX_EXCEL_EXPORT_DIMENSION}，"
            f"总像素不超过 {MAX_EXCEL_EXPORT_PIXELS}",
            self,
        )
        self.width_hint.setWordWrap(True)
        self.limit_hint.setWordWrap(True)
        self.width_hint.setStyleSheet("color:#cccccc;")
        self.limit_hint.setStyleSheet("color:#cccccc;")

        self.sheet_edit = QLineEdit(sheet_name, self)
        self.sheet_edit.setMaxLength(31)

        layout = QFormLayout(self)
        layout.addRow("导出宽度", self.width_spin)
        layout.addRow("参数说明", self.width_hint)
        layout.addRow("自动高度", self.height_value)
        layout.addRow("总单元格数", self.pixel_value)
        layout.addRow("工作表名称", self.sheet_edit)
        layout.addRow("导出边界", self.limit_hint)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        self.width_spin.valueChanged.connect(self._update_preview)
        self._update_preview(self.width_spin.value())

    def _update_preview(self, width: int) -> None:
        height = compute_scaled_height(self._image_width, self._image_height, width)
        self.width_hint.setText(
            f"系统建议值：{self._suggested_width}。你可以直接修改这个值。\n"
            f"当前设置表示导出宽度为 {width} 个像素，也就是 Excel 里大约 {width} 列；"
            "高度会按原图比例自动计算。"
        )
        self.height_value.setText(str(height))
        self.pixel_value.setText(f"{width} x {height} = {width * height}")

    def values(self) -> tuple[int, int, str]:
        width = self.width_spin.value()
        height = compute_scaled_height(self._image_width, self._image_height, width)
        sheet_name = sanitize_worksheet_name(
            self.sheet_edit.text(),
            datetime.now().strftime("%Y%m%d"),
        )
        return width, height, sheet_name


class ImageCanvasView(QGraphicsView):
    cursor_changed = pyqtSignal(int, int, bool)
    point_created = pyqtSignal(int, int)
    zoom_changed = pyqtSignal(float)
    viewport_changed = pyqtSignal(QRectF)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self._scene.setItemIndexMethod(QGraphicsScene.BspTreeIndex)
        self.setScene(self._scene)

        self._pixmap_item: Optional[QGraphicsPixmapItem] = None
        self._image_rect = QRectF()
        self._mouse_pos_view: Optional[QPoint] = None
        self._space_pressed = False
        self._manual_panning = False
        self._pan_button = Qt.NoButton
        self._pan_last_pos = QPoint()
        self._click_candidate = False
        self._click_origin = QPoint()
        self._grid_color = QColor(255, 255, 255, 70)
        self._crosshair_color = QColor("#ffd666")
        self._pixel_grid_threshold = 12.0
        self._smooth_threshold = 8.0
        self._min_zoom = 0.02
        self._max_zoom = 64.0

        self.setFrameShape(QFrame.NoFrame)
        self.setBackgroundBrush(QColor("#2b2b2b"))
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setDragMode(QGraphicsView.NoDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.setViewportUpdateMode(QGraphicsView.SmartViewportUpdate)
        self.setOptimizationFlags(
            QGraphicsView.DontSavePainterState | QGraphicsView.DontAdjustForAntialiasing
        )
        self.setRenderHint(QPainter.Antialiasing, False)
        self.setRenderHint(QPainter.TextAntialiasing, True)
        self.setRenderHint(QPainter.SmoothPixmapTransform, True)

        self.horizontalScrollBar().valueChanged.connect(self._emit_viewport_changed)
        self.verticalScrollBar().valueChanged.connect(self._emit_viewport_changed)
        self._update_view_cursor()

    def has_image(self) -> bool:
        return self._pixmap_item is not None

    def image_rect(self) -> QRectF:
        return QRectF(self._image_rect)

    def clear_scene(self) -> None:
        self._scene.clear()
        self._pixmap_item = None
        self._image_rect = QRectF()
        self.resetTransform()
        self._mouse_pos_view = None
        self._manual_panning = False
        self._click_candidate = False
        self._update_view_cursor()
        self.viewport().update()
        self._emit_zoom_changed()
        self._emit_viewport_changed()

    def load_image(self, image: QImage) -> None:
        self.clear_scene()
        pixmap = QPixmap.fromImage(image)
        self._pixmap_item = self._scene.addPixmap(pixmap)
        self._pixmap_item.setTransformationMode(Qt.SmoothTransformation)
        self._image_rect = QRectF(0, 0, image.width(), image.height())
        self._scene.setSceneRect(self._image_rect)
        self.fit_image()
        self.viewport().update()

    def fit_image(self) -> None:
        if not self.has_image():
            return
        self.fitInView(self._image_rect, Qt.KeepAspectRatio)
        self._update_pixmap_rendering()
        self.viewport().update()
        self._emit_zoom_changed()
        self._emit_viewport_changed()

    def current_zoom(self) -> float:
        return self.transform().m11()

    def set_guide_color(self, color: QColor) -> None:
        self._crosshair_color = QColor(color)
        self._grid_color = QColor(color)
        self._grid_color.setAlpha(70)
        self.viewport().update()

    def focus_on_point(self, x: int, y: int) -> None:
        if not self.has_image():
            return
        self.set_absolute_zoom(max(self.current_zoom(), 10.0))
        self.centerOn(QPointF(x + 0.5, y + 0.5))
        self.viewport().update()
        self._emit_viewport_changed()

    def center_on_scene_point(self, point: QPointF) -> None:
        if not self.has_image():
            return
        self.centerOn(point)
        self.viewport().update()
        self._emit_viewport_changed()

    def set_absolute_zoom(self, zoom: float) -> None:
        if not self.has_image():
            return
        zoom = max(self._min_zoom, min(self._max_zoom, zoom))
        current = self.current_zoom()
        if current <= 0 or abs(current - zoom) < 1e-6:
            return

        center = self.mapToScene(self.viewport().rect().center())
        self.scale(zoom / current, zoom / current)
        self.centerOn(center)
        self._update_pixmap_rendering()
        self.viewport().update()
        self._emit_zoom_changed()
        self._emit_viewport_changed()

    def visible_scene_rect(self) -> QRectF:
        if not self.has_image():
            return QRectF()
        return self.mapToScene(self.viewport().rect()).boundingRect()

    def _emit_zoom_changed(self) -> None:
        self.zoom_changed.emit(self.current_zoom() if self.has_image() else 1.0)

    def _emit_viewport_changed(self) -> None:
        self.viewport_changed.emit(self.visible_scene_rect())

    def _scene_pixel_at(self, view_pos: QPoint) -> Optional[QPoint]:
        if not self.has_image():
            return None
        scene_pos = self.mapToScene(view_pos)
        if not self._image_rect.contains(scene_pos):
            return None
        x = int(scene_pos.x())
        y = int(scene_pos.y())
        if 0 <= x < int(self._image_rect.width()) and 0 <= y < int(self._image_rect.height()):
            return QPoint(x, y)
        return None

    def _update_view_cursor(self) -> None:
        if not self.has_image():
            self.viewport().setCursor(Qt.ArrowCursor)
        elif self._manual_panning:
            self.viewport().setCursor(Qt.ClosedHandCursor)
        elif self._space_pressed:
            self.viewport().setCursor(Qt.OpenHandCursor)
        else:
            self.viewport().setCursor(Qt.BlankCursor)

    def _update_pixmap_rendering(self) -> None:
        if self._pixmap_item is None:
            return
        smooth = self.current_zoom() < self._smooth_threshold
        self.setRenderHint(QPainter.SmoothPixmapTransform, smooth)
        self._pixmap_item.setTransformationMode(
            Qt.SmoothTransformation if smooth else Qt.FastTransformation
        )

    def _start_manual_panning(self, button: Qt.MouseButton, pos: QPoint) -> None:
        self._manual_panning = True
        self._pan_button = button
        self._pan_last_pos = QPoint(pos)
        self._click_candidate = False
        self._update_view_cursor()

    def _stop_manual_panning(self) -> None:
        self._manual_panning = False
        self._pan_button = Qt.NoButton
        self._update_view_cursor()

    def wheelEvent(self, event) -> None:
        if not self.has_image():
            super().wheelEvent(event)
            return

        delta = event.angleDelta().y()
        if delta == 0:
            event.ignore()
            return

        factor = 1.15 ** (delta / 120.0)
        current = self.current_zoom()
        target = max(self._min_zoom, min(self._max_zoom, current * factor))
        actual_factor = target / current if current else 1.0
        if abs(actual_factor - 1.0) < 1e-6:
            event.accept()
            return

        self.scale(actual_factor, actual_factor)
        self._update_pixmap_rendering()
        self.viewport().update()
        self._emit_zoom_changed()
        self._emit_viewport_changed()
        event.accept()

    def mousePressEvent(self, event) -> None:
        self.setFocus()
        if not self.has_image():
            super().mousePressEvent(event)
            return

        if event.button() == Qt.RightButton:
            self._start_manual_panning(Qt.RightButton, event.pos())
            event.accept()
            return

        if event.button() == Qt.LeftButton and self._space_pressed:
            self._start_manual_panning(Qt.LeftButton, event.pos())
            event.accept()
            return

        if event.button() == Qt.LeftButton:
            self._click_candidate = True
            self._click_origin = QPoint(event.pos())
            event.accept()
            return

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        self._mouse_pos_view = QPoint(event.pos())

        pixel = self._scene_pixel_at(event.pos())
        if pixel is None:
            self.cursor_changed.emit(-1, -1, False)
        else:
            self.cursor_changed.emit(pixel.x(), pixel.y(), True)

        if self._manual_panning:
            delta = event.pos() - self._pan_last_pos
            self._pan_last_pos = QPoint(event.pos())
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - delta.y())
            self.viewport().update()
            event.accept()
            return

        if self._click_candidate:
            if (
                event.pos() - self._click_origin
            ).manhattanLength() > QApplication.startDragDistance():
                self._click_candidate = False

        self.viewport().update()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self._manual_panning and event.button() == self._pan_button:
            self._stop_manual_panning()
            self.viewport().update()
            event.accept()
            return

        if event.button() == Qt.LeftButton and self._click_candidate and not self._space_pressed:
            self._click_candidate = False
            pixel = self._scene_pixel_at(event.pos())
            if pixel is not None:
                self.point_created.emit(pixel.x(), pixel.y())
            event.accept()
            return

        self._click_candidate = False
        super().mouseReleaseEvent(event)

    def leaveEvent(self, event) -> None:
        self._mouse_pos_view = None
        self.cursor_changed.emit(-1, -1, False)
        self.viewport().update()
        super().leaveEvent(event)

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key_Space and not event.isAutoRepeat():
            self._space_pressed = True
            self._update_view_cursor()
            event.accept()
            return
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event) -> None:
        if event.key() == Qt.Key_Space and not event.isAutoRepeat():
            self._space_pressed = False
            if self._manual_panning and self._pan_button == Qt.LeftButton:
                self._stop_manual_panning()
            self._update_view_cursor()
            event.accept()
            return
        super().keyReleaseEvent(event)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._emit_viewport_changed()

    def drawForeground(self, painter: QPainter, rect: QRectF) -> None:
        if self.has_image() and self.current_zoom() >= self._pixel_grid_threshold:
            visible = rect.intersected(self._image_rect)
            if not visible.isEmpty():
                left = math.floor(visible.left())
                right = math.ceil(visible.right())
                top = math.floor(visible.top())
                bottom = math.ceil(visible.bottom())

                path = QPainterPath()
                for x in range(left, right + 1):
                    path.moveTo(x, top)
                    path.lineTo(x, bottom)
                for y in range(top, bottom + 1):
                    path.moveTo(left, y)
                    path.lineTo(right, y)

                painter.save()
                painter.setPen(QPen(self._grid_color, 0))
                painter.drawPath(path)
                painter.restore()

        if self._mouse_pos_view is not None:
            painter.save()
            painter.resetTransform()
            painter.setPen(QPen(self._crosshair_color, 1))
            painter.drawLine(
                0,
                self._mouse_pos_view.y(),
                self.viewport().width(),
                self._mouse_pos_view.y(),
            )
            painter.drawLine(
                self._mouse_pos_view.x(),
                0,
                self._mouse_pos_view.x(),
                self.viewport().height(),
            )
            painter.restore()


class MainWindow(QMainWindow):
    POINT_COLORS = [
        QColor("#ff6b6b"),
        QColor("#4ecdc4"),
        QColor("#ffe66d"),
        QColor("#5dade2"),
        QColor("#a569bd"),
        QColor("#f5b041"),
        QColor("#58d68d"),
        QColor("#ec7063"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("高精度图像坐标取点工具V1.2")
        self.resize(1460, 920)

        self._image: Optional[QImage] = None
        self._image_path: Optional[Path] = None
        self._points: list[PointRecord] = []
        self._guide_color = QColor("#ffd666")

        self._build_ui()
        self._build_menu()
        self._connect_signals()
        self._refresh_actions()

    def _build_ui(self) -> None:
        self.canvas = ImageCanvasView(self)
        self.canvas.set_guide_color(self._guide_color)

        self.overview = OverviewWidget(self)
        self.magnifier = MagnifierWidget(self)
        self.point_table = QTableWidget(self)
        self.point_table.setColumnCount(4)
        self.point_table.setHorizontalHeaderLabels(["Index", "X", "Y", "Color"])
        self.point_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.point_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.point_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.point_table.setAlternatingRowColors(True)
        self.point_table.setSortingEnabled(False)
        self.point_table.verticalHeader().setVisible(False)
        self.point_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.point_table.horizontalHeader().setStretchLastSection(True)

        self.edit_button = QPushButton("Edit Point", self)
        self.delete_button = QPushButton("Delete Point", self)
        self.fit_button = QPushButton("Fit Image", self)
        self.export_excel_button = QPushButton("导出 Excel 像素画", self)
        self.export_excel_button.setObjectName("primaryActionButton")
        self.export_excel_hint = QLabel(
            "把当前图片按像素映射到 Excel 单元格。\n"
            "系统会自动给出建议宽度，你也可以手动调整。",
            self,
        )
        self.export_excel_hint.setWordWrap(True)
        self.export_excel_hint.setStyleSheet("color:#cfcfcf; font-size:12px;")

        sidebar = QWidget(self)
        sidebar.setMinimumWidth(320)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(12, 12, 12, 12)
        sidebar_layout.setSpacing(12)

        overview_label = QLabel("Overview", self)
        magnifier_label = QLabel("Magnifier", self)
        points_label = QLabel("Points", self)
        for label in (overview_label, magnifier_label, points_label):
            label.setStyleSheet("font-weight:600; color:#eaeaea;")

        button_row = QHBoxLayout()
        button_row.addWidget(self.edit_button)
        button_row.addWidget(self.delete_button)

        sidebar_layout.addWidget(overview_label)
        sidebar_layout.addWidget(self.overview)
        sidebar_layout.addWidget(magnifier_label)
        sidebar_layout.addWidget(self.magnifier)
        sidebar_layout.addWidget(points_label)
        sidebar_layout.addWidget(self.point_table, 1)
        sidebar_layout.addLayout(button_row)
        sidebar_layout.addWidget(self.fit_button)
        sidebar_layout.addWidget(self.export_excel_button)
        sidebar_layout.addWidget(self.export_excel_hint)

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(self.canvas)
        splitter.addWidget(sidebar)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setSizes([1100, 340])
        self.setCentralWidget(splitter)

        self.setStyleSheet(
            """
            QMainWindow { background:#222; }
            QDialog, QMessageBox, QProgressDialog {
                background:#222;
                color:#f0f0f0;
            }
            QTableWidget {
                background:#1f1f1f;
                alternate-background-color:#262626;
                color:#f0f0f0;
                border:1px solid #444;
                gridline-color:#353535;
                selection-background-color:#3d4f6a;
                selection-color:#f0f0f0;
            }
            QHeaderView::section {
                background:#2a2a2a;
                color:#f0f0f0;
                border:1px solid #444;
                padding:4px;
            }
            QLineEdit, QSpinBox {
                background:#1f1f1f;
                color:#f0f0f0;
                border:1px solid #4a4a4a;
                padding:4px 6px;
                selection-background-color:#3d4f6a;
            }
            QDialogButtonBox QPushButton,
            QMessageBox QPushButton {
                min-width:88px;
            }
            QMessageBox QLabel,
            QProgressDialog QLabel,
            QDialog QLabel {
                color:#f0f0f0;
                background:transparent;
            }
            QTableWidget::item {
                background:#1f1f1f;
                color:#f0f0f0;
            }
            QTableWidget::item:alternate {
                background:#262626;
            }
            QTableWidget::item:selected { background:#3d4f6a; }
            QPushButton {
                background:#303030;
                color:#f0f0f0;
                border:1px solid #4a4a4a;
                padding:6px 10px;
            }
            QPushButton:hover { background:#3a3a3a; }
            QPushButton#primaryActionButton {
                background:#c65d1a;
                border:1px solid #d8782f;
                color:#ffffff;
                font-weight:600;
                padding:9px 12px;
            }
            QPushButton#primaryActionButton:hover {
                background:#da6a22;
            }
            QLabel { color:#f0f0f0; }
            """
        )

        self.image_status = QLabel("Image: -", self)
        self.zoom_status = QLabel("Zoom: -", self)
        self.cursor_status = QLabel("Current (Original): -", self)
        self.statusBar().addPermanentWidget(self.image_status)
        self.statusBar().addPermanentWidget(self.zoom_status)
        self.statusBar().addPermanentWidget(self.cursor_status, 1)

        delete_shortcut = QShortcut(QKeySequence.Delete, self.point_table)
        delete_shortcut.activated.connect(self.delete_selected_point)

    def _build_menu(self) -> None:
        menu = self.menuBar()

        file_menu = menu.addMenu("File")
        self.open_action = QAction(
            self.style().standardIcon(QStyle.SP_DialogOpenButton), "Open Image...", self
        )
        self.open_action.setShortcut("Ctrl+O")
        self.save_points_action = QAction("Save Coordinates...", self)
        self.save_points_action.setShortcut("Ctrl+S")
        self.export_excel_action = QAction("Export Pixel To Excel...", self)
        self.export_excel_action.setShortcut("Ctrl+E")
        self.exit_action = QAction("Exit", self)

        file_menu.addAction(self.open_action)
        file_menu.addAction(self.save_points_action)
        file_menu.addAction(self.export_excel_action)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_action)

        settings_menu = menu.addMenu("Settings")
        self.guide_color_action = QAction("Guide Color...", self)
        self.magnifier_zoom_action = QAction("Magnifier Zoom...", self)
        settings_menu.addAction(self.guide_color_action)
        settings_menu.addAction(self.magnifier_zoom_action)

    def _connect_signals(self) -> None:
        self.open_action.triggered.connect(self.open_image)
        self.save_points_action.triggered.connect(self.save_points)
        self.export_excel_action.triggered.connect(self.export_pixel_to_excel)
        self.exit_action.triggered.connect(self.close)
        self.guide_color_action.triggered.connect(self.choose_guide_color)
        self.magnifier_zoom_action.triggered.connect(self.choose_magnifier_zoom)

        self.canvas.cursor_changed.connect(self._handle_cursor_changed)
        self.canvas.point_created.connect(self.add_point)
        self.canvas.zoom_changed.connect(self._update_zoom_status)
        self.canvas.viewport_changed.connect(self.overview.set_viewport_rect)

        self.overview.requested_center.connect(self.canvas.center_on_scene_point)

        self.point_table.itemSelectionChanged.connect(self._sync_marker_highlight)
        self.point_table.itemSelectionChanged.connect(self._refresh_actions)
        self.point_table.itemClicked.connect(self.focus_selected_point)
        self.point_table.itemDoubleClicked.connect(self.edit_selected_point)

        self.edit_button.clicked.connect(self.edit_selected_point)
        self.delete_button.clicked.connect(self.delete_selected_point)
        self.fit_button.clicked.connect(self.canvas.fit_image)
        self.export_excel_button.clicked.connect(self.export_pixel_to_excel)

    def _refresh_actions(self) -> None:
        has_image = self._image is not None and not self._image.isNull()
        has_points = bool(self._points)
        has_selection = self._selected_point_index() is not None

        self.save_points_action.setEnabled(has_image and has_points)
        self.export_excel_action.setEnabled(has_image)
        self.export_excel_button.setEnabled(has_image)
        self.edit_button.setEnabled(has_selection)
        self.delete_button.setEnabled(has_selection)
        self.fit_button.setEnabled(has_image)

    def open_image(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Image",
            "",
            "Images (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp)",
        )
        if path:
            self.load_image(Path(path))

    def load_image(self, path: Path) -> None:
        reader = QImageReader(str(path))
        reader.setAutoTransform(True)
        image = reader.read()
        if image.isNull():
            QMessageBox.critical(
                self,
                "Open Failed",
                reader.errorString() or f"Unable to load image: {path}",
            )
            return

        self._image = image
        self._image_path = path
        self._clear_points()
        self.canvas.load_image(image)
        self.magnifier.set_image(image)
        self.magnifier.set_point(None)
        self.overview.set_image(image)
        self.overview.set_viewport_rect(self.canvas.visible_scene_rect())

        self.image_status.setText(f"Image: {image.width()} x {image.height()}")
        self.cursor_status.setText("Current (Original): -")
        self.statusBar().showMessage(f"Loaded {path.name}", 4000)
        self._refresh_actions()

    def _clear_points(self) -> None:
        for record in self._points:
            if record.marker.scene() is not None:
                record.marker.scene().removeItem(record.marker)
        self._points.clear()
        self.point_table.setRowCount(0)
        self.point_table.clearContents()
        self._refresh_actions()

    def _next_point_color(self) -> QColor:
        return QColor(self.POINT_COLORS[len(self._points) % len(self.POINT_COLORS)])

    def add_point(self, x: int, y: int) -> None:
        if self._image is None:
            return

        color = self._next_point_color()
        marker = PointMarkerItem(len(self._points) + 1, color)
        marker.setPos(QPointF(x + 0.5, y + 0.5))
        self.canvas.scene().addItem(marker)

        record = PointRecord(
            index=len(self._points) + 1,
            x=x,
            y=y,
            color=color,
            marker=marker,
        )
        self._points.append(record)
        self._rebuild_point_list(select_row=len(self._points) - 1)
        self.statusBar().showMessage(f"Point #{record.index} added at ({x}, {y})", 3000)

    def _rebuild_point_list(self, select_row: Optional[int] = None) -> None:
        self.point_table.blockSignals(True)
        self.point_table.clearContents()
        self.point_table.setRowCount(len(self._points))

        for index, record in enumerate(self._points, start=1):
            record.index = index
            record.marker.set_index(index)
            record.marker.set_color(record.color)

            values = [
                f"{index:02d}",
                str(record.x),
                str(record.y),
                record.color.name().upper(),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.UserRole, index - 1)
                if column == 3:
                    item.setForeground(QBrush(record.color))
                self.point_table.setItem(index - 1, column, item)

        if select_row is not None and 0 <= select_row < self.point_table.rowCount():
            self.point_table.selectRow(select_row)

        self.point_table.blockSignals(False)
        self._sync_marker_highlight()
        self._refresh_actions()

    def _selected_point_index(self) -> Optional[int]:
        row = self.point_table.currentRow()
        if row < 0:
            return None
        item = self.point_table.item(row, 0)
        if item is None:
            return None
        index = item.data(Qt.UserRole)
        if isinstance(index, int) and 0 <= index < len(self._points):
            return index
        return None

    def _sync_marker_highlight(self) -> None:
        selected = self._selected_point_index()
        for idx, record in enumerate(self._points):
            record.marker.set_highlighted(idx == selected)

    def focus_selected_point(self, *_args) -> None:
        index = self._selected_point_index()
        if index is None:
            return
        record = self._points[index]
        self.canvas.focus_on_point(record.x, record.y)

    def edit_selected_point(self, *_args) -> None:
        index = self._selected_point_index()
        if index is None or self._image is None:
            return

        record = self._points[index]
        dialog = PointEditDialog(
            record.x,
            record.y,
            record.color,
            self._image.width() - 1,
            self._image.height() - 1,
            self,
        )
        if dialog.exec_() != QDialog.Accepted:
            return

        x, y, color = dialog.values()
        record.x = x
        record.y = y
        record.color = color
        record.marker.setPos(QPointF(x + 0.5, y + 0.5))
        record.marker.set_color(color)
        self._rebuild_point_list(select_row=index)
        self.canvas.focus_on_point(x, y)
        self.statusBar().showMessage(f"Point #{record.index} updated", 3000)

    def delete_selected_point(self, *_args) -> None:
        index = self._selected_point_index()
        if index is None:
            return
        record = self._points.pop(index)
        if record.marker.scene() is not None:
            record.marker.scene().removeItem(record.marker)
        next_row = min(index, len(self._points) - 1) if self._points else None
        self._rebuild_point_list(select_row=next_row)
        self.statusBar().showMessage(f"Point #{record.index} deleted", 3000)

    def save_points(self) -> None:
        if self._image is None or not self._points:
            return

        default_path = (
            str(self._image_path.with_suffix(".points.json"))
            if self._image_path is not None
            else "points.json"
        )
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Coordinates",
            default_path,
            "JSON (*.json)",
        )
        if not path:
            return

        payload = {
            "image_path": str(self._image_path) if self._image_path else None,
            "image_size": {"width": self._image.width(), "height": self._image.height()},
            "points": [
                {
                    "index": record.index,
                    "x": record.x,
                    "y": record.y,
                    "color": record.color.name(),
                }
                for record in self._points
            ],
        }

        try:
            Path(path).write_text(
                json.dumps(payload, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except OSError as exc:
            QMessageBox.critical(self, "Save Failed", str(exc))
            return

        self.statusBar().showMessage(f"Coordinates saved to {Path(path).name}", 4000)

    def export_pixel_to_excel(self) -> None:
        if self._image is None or self._image_path is None:
            return
        if xlsxwriter is None:
            QMessageBox.critical(
                self,
                "Dependency Missing",
                "XlsxWriter is not installed. Please run:\n\npip install XlsxWriter",
            )
            return

        max_width = compute_max_excel_width(self._image.width(), self._image.height())
        if max_width < 1:
            QMessageBox.warning(
                self,
                "Export Blocked",
                "The current image aspect ratio is too extreme to fit within the "
                f"Excel export limit of {MAX_EXCEL_EXPORT_DIMENSION}x"
                f"{MAX_EXCEL_EXPORT_DIMENSION} while preserving aspect ratio.",
            )
            return

        sheet_name = sanitize_worksheet_name(
            self._image_path.stem,
            datetime.now().strftime("%Y%m%d"),
        )
        dialog = PixelToExcelDialog(
            self._image.width(),
            self._image.height(),
            sheet_name,
            self,
        )
        if dialog.exec_() != QDialog.Accepted:
            return

        target_width, target_height, worksheet_name = dialog.values()
        if (
            target_width > MAX_EXCEL_EXPORT_DIMENSION
            or target_height > MAX_EXCEL_EXPORT_DIMENSION
            or target_width * target_height > MAX_EXCEL_EXPORT_PIXELS
        ):
            QMessageBox.warning(
                self,
                "Export Blocked",
                "The selected size exceeds the current export boundary.",
            )
            return

        total_pixels = target_width * target_height
        if total_pixels >= LARGE_EXCEL_EXPORT_WARNING_PIXELS:
            answer = QMessageBox.question(
                self,
                "Large Excel Export",
                "The selected size is large and may take noticeable time to generate.\n\n"
                f"Target size: {target_width} x {target_height} ({total_pixels} cells)\n\n"
                "Continue exporting?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer != QMessageBox.Yes:
                return

        default_output = str(self._image_path.with_name(f"{self._image_path.stem}_pixel.xlsx"))
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Pixel To Excel",
            default_output,
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return

        output_path = Path(path)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")

        progress_dialog = QProgressDialog(
            "Preparing Excel export...\nEstimating time...",
            "Cancel",
            0,
            target_height,
            self,
        )
        progress_dialog.setWindowTitle("Export Pixel To Excel")
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.setAutoClose(True)
        progress_dialog.setValue(0)
        export_started_at = perf_counter()

        def progress_callback(done_rows: int, total_rows: int) -> None:
            elapsed = perf_counter() - export_started_at
            ratio = done_rows / total_rows if total_rows else 0.0
            percent = ratio * 100
            if done_rows > 0 and elapsed > 0:
                remaining = max(0.0, elapsed / done_rows * (total_rows - done_rows))
                eta_text = format_duration(remaining)
            else:
                eta_text = "estimating..."

            progress_dialog.setMaximum(total_rows)
            progress_dialog.setLabelText(
                f"Exporting row {done_rows}/{total_rows} ({percent:.1f}%)\n"
                f"Size: {target_width} x {target_height}\n"
                f"Elapsed: {format_duration(elapsed)} | Remaining: {eta_text}"
            )
            progress_dialog.setValue(done_rows)
            self.statusBar().showMessage(
                f"Excel export {percent:.1f}%  "
                f"Elapsed {format_duration(elapsed)}  "
                f"Remaining {eta_text}",
                0,
            )
            QApplication.processEvents()
            if progress_dialog.wasCanceled():
                raise RuntimeError("Excel export canceled.")

        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.statusBar().showMessage(
            f"Exporting {target_width} x {target_height} Excel pixel art...",
            0,
        )
        QApplication.processEvents()
        try:
            actual_width, actual_height = export_image_to_excel(
                self._image_path,
                output_path,
                target_width,
                worksheet_name,
                progress_callback=progress_callback,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", str(exc))
        else:
            self.statusBar().showMessage(
                f"Excel exported: {output_path.name} ({actual_width} x {actual_height})",
                5000,
            )
        finally:
            progress_dialog.close()
            QApplication.restoreOverrideCursor()

    def choose_guide_color(self) -> None:
        color = QColorDialog.getColor(self._guide_color, self, "Guide Color")
        if color.isValid():
            self._guide_color = color
            self.canvas.set_guide_color(color)

    def choose_magnifier_zoom(self) -> None:
        value, ok = QInputDialog.getInt(
            self,
            "Magnifier Zoom",
            "Magnifier scale:",
            value=self.magnifier.zoom_factor(),
            min=2,
            max=30,
            step=1,
        )
        if ok:
            self.magnifier.set_zoom_factor(value)

    def _handle_cursor_changed(self, x: int, y: int, valid: bool) -> None:
        if valid:
            self.cursor_status.setText(f"Current (Original): [{x}, {y}]")
            self.magnifier.set_point(QPoint(x, y))
        else:
            self.cursor_status.setText("Current (Original): -")
            self.magnifier.set_point(None)

    def _update_zoom_status(self, zoom: float) -> None:
        if self._image is None:
            self.zoom_status.setText("Zoom: -")
        else:
            self.zoom_status.setText(f"Zoom: {zoom:.2f}x")


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("高精度图像坐标取点工具V1.2")
    icon_path = resource_path("assets/app_icon.ico")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    if icon_path.exists():
        window.setWindowIcon(QIcon(str(icon_path)))
    if len(sys.argv) > 1:
        image_path = Path(sys.argv[1])
        if image_path.exists():
            window.load_image(image_path)
    window.show()
    return app.exec_()


if __name__ == "__main__":
    raise SystemExit(main())
