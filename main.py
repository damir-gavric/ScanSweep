import os
import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QRectF, QSize, QSettings, QThread, Qt, Signal
from PySide6.QtGui import (
    QColor,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QFontDatabase,
    QIcon,
    QPainter,
    QPen,
)
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFontComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from conversion import convert_with_libreoffice, needs_conversion
from audit_log import AuditLog
from processor import (
    DEFAULT_FORMATTING,
    QUOTE_LANGUAGES,
    CleaningCancelled,
    process_docx,
    quote_example,
)


APP_VERSION = "2.2"
OUTPUT_FORMATS = (".docx", ".odt")

# The presets this version replaced, used once to carry an old choice over.
RETIRED_PROFILES = {
    "novel": {"font_name": "Garamond", "font_size": 12, "line_spacing": 1.15,
              "first_line_indent_cm": 1.0, "close_slash_spacing": False,
              "protect_legal_numbering": False},
    "academic": {"font_name": "Arial", "font_size": 11, "line_spacing": 1.15,
                 "first_line_indent_cm": 1.0, "close_slash_spacing": True,
                 "protect_legal_numbering": False},
    "legal": {"font_name": "Times New Roman", "font_size": 12, "line_spacing": 1.0,
              "first_line_indent_cm": 0.0, "close_slash_spacing": False,
              "protect_legal_numbering": True},
}
SETTINGS_FILE_NAME = "ScanSweep.ini"


def application_directory():
    """Directory the app runs from: the executable's when frozen, the source's otherwise."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name):
    """A bundled file: beside the source, or inside the unpacked bundle when frozen."""
    return Path(__file__).resolve().with_name(name)


def theme_stylesheet(sheet):
    """A theme with its image paths filled in; Qt needs forward slashes."""
    return (
        sheet.replace("__CHEVRON_DOWN__", resource_path("chevron_down.svg").as_posix())
        .replace("__CHEVRON_UP__", resource_path("chevron_up.svg").as_posix())
    )


def settings_file_path():
    """Keep settings beside the app so the portable build leaves no trace behind."""
    return str(application_directory() / SETTINGS_FILE_NAME)


DARK_THEME = """
QMainWindow, QWidget#centralPanel {
    background-color: #1f1f1f;
    color: #f4f4f4;
}
QGroupBox {
    font-weight: 600;
    border: 1px solid #3a3a3a;
    border-radius: 10px;
    margin-top: 10px;
    padding-top: 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 4px;
}
QListWidget, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    border: 1px solid #404040;
    border-radius: 8px;
    padding: 6px;
    background-color: #222222;
    color: #f4f4f4;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 26px;
    border: none;
    background: transparent;
}
QComboBox::down-arrow {
    image: url(__CHEVRON_DOWN__);
    width: 12px;
    height: 8px;
}
QSpinBox::up-button, QDoubleSpinBox::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 24px;
    height: 15px;
    margin: 3px 4px 0 0;
    border: none;
    background: transparent;
}
QSpinBox::down-button, QDoubleSpinBox::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 24px;
    height: 15px;
    margin: 0 4px 3px 0;
    border: none;
    background: transparent;
}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {
    image: url(__CHEVRON_UP__);
    width: 11px;
    height: 7px;
}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
    image: url(__CHEVRON_DOWN__);
    width: 11px;
    height: 7px;
}
QComboBox QAbstractItemView {
    border: 1px solid #404040;
    background-color: #222222;
    color: #f4f4f4;
    selection-background-color: #0e6b80;
    selection-color: #ffffff;
    outline: none;
}
QPushButton {
    padding: 8px 14px;
    border: 1px solid #4a4a4a;
    border-radius: 8px;
    background-color: #2b2b2b;
    color: #f4f4f4;
}
QPushButton:disabled {
    color: #8e8e8e;
}
QPushButton[infoButton="true"] {
    min-width: 30px;
    max-width: 30px;
    min-height: 30px;
    max-height: 30px;
    padding: 0;
    border: none;
    background: transparent;
}
QLabel, QCheckBox, QRadioButton, QGroupBox {
    background: transparent;
    color: #f4f4f4;
}
QCheckBox, QRadioButton {
    spacing: 8px;
}
QRadioButton::indicator {
    width: 18px;
    height: 18px;
    border-radius: 10px;
    border: 1px solid #6a6a6a;
    background-color: #222222;
}
QRadioButton::indicator:hover {
    border-color: #8f8f8f;
}
QRadioButton::indicator:checked {
    width: 10px;
    height: 10px;
    border: 5px solid #59c4ff;
    border-radius: 10px;
    background-color: #1f1f1f;
}
QProgressBar {
    border: 1px solid #505050;
    border-radius: 8px;
    text-align: center;
    background-color: #202020;
    color: #f4f4f4;
}
QProgressBar::chunk {
    border-radius: 7px;
    background-color: #59c4ff;
}
QStatusBar {
    color: #c7c7c7;
}
QLabel#appTitle {
    font-size: 28px;
    font-weight: 700;
    color: #f4f4f4;
}
QLabel#appSubtitle, QLabel#hintLabel, QLabel#noteLabel {
    color: #9aa0a6;
    font-size: 13px;
}
ThemeSwitchButton {
    background: transparent;
    border: none;
}
"""


LIGHT_THEME = """
QMainWindow, QWidget#centralPanel {
    background-color: #f4f6fa;
    color: #1d2433;
}
QGroupBox {
    font-weight: 600;
    border: 1px solid #cfd7e6;
    border-radius: 10px;
    margin-top: 10px;
    padding-top: 10px;
    background-color: #fbfcfe;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 4px;
}
QListWidget, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    border: 1px solid #c5d0e0;
    border-radius: 8px;
    padding: 6px;
    background-color: #ffffff;
    color: #1d2433;
}
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 26px;
    border: none;
    background: transparent;
}
QComboBox::down-arrow {
    image: url(__CHEVRON_DOWN__);
    width: 12px;
    height: 8px;
}
QSpinBox::up-button, QDoubleSpinBox::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 24px;
    height: 15px;
    margin: 3px 4px 0 0;
    border: none;
    background: transparent;
}
QSpinBox::down-button, QDoubleSpinBox::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 24px;
    height: 15px;
    margin: 0 4px 3px 0;
    border: none;
    background: transparent;
}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {
    image: url(__CHEVRON_UP__);
    width: 11px;
    height: 7px;
}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
    image: url(__CHEVRON_DOWN__);
    width: 11px;
    height: 7px;
}
QComboBox QAbstractItemView {
    border: 1px solid #c5d0e0;
    background-color: #ffffff;
    color: #1d2433;
    selection-background-color: #cfe9ff;
    selection-color: #102136;
    outline: none;
}
QPushButton {
    padding: 8px 14px;
    border: 1px solid #c5d0e0;
    border-radius: 8px;
    background-color: #ffffff;
    color: #1d2433;
}
QPushButton:disabled {
    color: #8590a3;
}
QPushButton[infoButton="true"] {
    min-width: 30px;
    max-width: 30px;
    min-height: 30px;
    max-height: 30px;
    padding: 0;
    border: none;
    background: transparent;
}
QLabel, QCheckBox, QRadioButton, QGroupBox {
    background: transparent;
    color: #1d2433;
}
QCheckBox, QRadioButton {
    spacing: 8px;
}
QRadioButton::indicator {
    width: 18px;
    height: 18px;
    border-radius: 10px;
    border: 1px solid #a6b2c4;
    background-color: #ffffff;
}
QRadioButton::indicator:hover {
    border-color: #7d8ca3;
}
QRadioButton::indicator:checked {
    width: 10px;
    height: 10px;
    border: 5px solid #1683d8;
    border-radius: 10px;
    background-color: #ffffff;
}
QProgressBar {
    border: 1px solid #c5d0e0;
    border-radius: 8px;
    text-align: center;
    background-color: #ffffff;
    color: #1d2433;
}
QProgressBar::chunk {
    border-radius: 7px;
    background-color: #1683d8;
}
QStatusBar {
    color: #516075;
}
QLabel#appTitle {
    font-size: 28px;
    font-weight: 700;
    color: #1d2433;
}
QLabel#appSubtitle, QLabel#hintLabel, QLabel#noteLabel {
    color: #637086;
    font-size: 13px;
}
ThemeSwitchButton {
    background: transparent;
    border: none;
}
"""


MESSAGE_BOX_DARK_THEME = """
QMessageBox {
    background-color: #1f1f1f;
}
QMessageBox QLabel {
    color: #f4f4f4;
    background: transparent;
    min-width: 280px;
}
QMessageBox QPushButton {
    min-width: 88px;
    padding: 8px 14px;
    border: 1px solid #4a4a4a;
    border-radius: 8px;
    background-color: #2b2b2b;
    color: #f4f4f4;
}
QMessageBox QPushButton:hover {
    background-color: #343434;
}
"""


MESSAGE_BOX_LIGHT_THEME = """
QMessageBox {
    background-color: #f4f6fa;
}
QMessageBox QLabel {
    color: #1d2433;
    background: transparent;
    min-width: 280px;
}
QMessageBox QPushButton {
    min-width: 88px;
    padding: 8px 14px;
    border: 1px solid #c5d0e0;
    border-radius: 8px;
    background-color: #ffffff;
    color: #1d2433;
}
QMessageBox QPushButton:hover {
    background-color: #f8fbff;
}
"""


class FileListWidget(QListWidget):
    files_dropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if not event.mimeData().hasUrls():
            super().dropEvent(event)
            return

        paths = []
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            path = url.toLocalFile()
            if os.path.splitext(path)[1].lower() in {".docx", ".odt"}:
                paths.append(path)

        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class ThemeSwitchButton(QPushButton):
    theme_changed = Signal(str)

    def __init__(self):
        super().__init__()
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedSize(92, 30)
        self.setToolTip("Switch between dark and light theme")
        self.toggled.connect(self._emit_theme)

    def sizeHint(self):
        return QSize(92, 30)

    def theme_name(self):
        return "dark" if self.isChecked() else "light"

    def set_theme(self, theme_name):
        target_checked = theme_name == "dark"
        if self.isChecked() != target_checked:
            self.setChecked(target_checked)
        else:
            self.update()

    def _emit_theme(self, checked):
        self.theme_changed.emit("dark" if checked else "light")
        self.update()

    def paintEvent(self, event):
        del event
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        is_dark = self.isChecked()
        outer_rect = QRectF(0.5, 0.5, self.width() - 1.0, self.height() - 1.0)
        knob_size = self.height() - 6
        knob_y = 3
        knob_x = 3 if is_dark else self.width() - knob_size - 3
        knob_rect = QRectF(knob_x, knob_y, knob_size, knob_size)

        bg_color = QColor("#050505") if is_dark else QColor("#e8eaee")
        border_color = QColor("#050505") if is_dark else QColor("#d3d7df")
        text_color = QColor("#ffffff") if is_dark else QColor("#111827")
        knob_fill = QColor("#ffffff")
        knob_stroke = QColor("#111111") if is_dark else QColor("#d3d7df")
        icon_color = QColor("#111111")

        painter.setPen(QPen(border_color, 1.2))
        painter.setBrush(bg_color)
        painter.drawRoundedRect(outer_rect, self.height() / 2, self.height() / 2)

        painter.setPen(QPen(knob_stroke, 1.0))
        painter.setBrush(knob_fill)
        painter.drawEllipse(knob_rect)

        font = QFont(self.font())
        font.setPointSize(7)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(text_color)

        if is_dark:
            text_rect = QRectF(knob_rect.right() + 8, 0, self.width() - knob_rect.right() - 12, self.height())
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, "DARK")
            self._draw_moon_icon(painter, knob_rect, icon_color)
        else:
            text_rect = QRectF(10, 0, knob_rect.left() - 14, self.height())
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, "LIGHT")
            self._draw_sun_icon(painter, knob_rect, icon_color)

    def _draw_sun_icon(self, painter, rect, color):
        center = rect.center()
        radius = rect.width() * 0.18
        painter.setPen(QPen(color, 1.6))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(center, radius, radius)
        ray_inner = rect.width() * 0.28
        ray_outer = rect.width() * 0.38
        for dx, dy in (
            (0, -1),
            (0.7, -0.7),
            (1, 0),
            (0.7, 0.7),
            (0, 1),
            (-0.7, 0.7),
            (-1, 0),
            (-0.7, -0.7),
        ):
            painter.drawLine(
                center.x() + dx * ray_inner,
                center.y() + dy * ray_inner,
                center.x() + dx * ray_outer,
                center.y() + dy * ray_outer,
            )

    def _draw_moon_icon(self, painter, rect, color):
        center = rect.center()
        radius = rect.width() * 0.22
        painter.setPen(QPen(color, 1.8))
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(center, radius, radius)
        painter.setPen(QPen(Qt.white, 3.2))
        painter.drawEllipse(center.x() + radius * 0.45, center.y() - radius * 0.05, radius * 0.9, radius * 0.9)


class CleanerWorker(QThread):
    log_message = Signal(str)
    file_progress = Signal(int, str)
    overall_progress = Signal(int)
    file_started = Signal(int, int, str)
    finished_ok = Signal()
    cancelled = Signal()
    failed = Signal(str)

    def __init__(
        self,
        sources,
        batch_mode,
        settings,
        quote_language,
        output_format,
        options,
        output_dir=None,
        output_file=None,
    ):
        super().__init__()
        self.sources = sources
        self.batch_mode = batch_mode
        self.settings = settings
        self.quote_language = quote_language
        self.output_format = output_format
        self.options = options
        self.output_dir = output_dir
        self.output_file = output_file
        self.cancel_requested = False

    def request_cancel(self):
        self.cancel_requested = True

    def run(self):
        try:
            total_files = len(self.sources)
            for index, src in enumerate(self.sources, start=1):
                if self.cancel_requested:
                    self.cancelled.emit()
                    return

                self.file_started.emit(index, total_files, src)

                if self.batch_mode:
                    base_name = os.path.splitext(os.path.basename(src))[0]
                    dst = os.path.join(self.output_dir, f"{base_name}_cleaned{self.output_format}")
                else:
                    dst = self.output_file

                self.overall_progress.emit(int(((index - 1) / total_files) * 100))

                def log_callback(message):
                    self.log_message.emit(message)

                def progress_callback(percent, label):
                    self.file_progress.emit(percent, label)
                    overall = ((index - 1) + (percent / 100.0)) / total_files
                    self.overall_progress.emit(int(overall * 100))

                def should_cancel():
                    return self.cancel_requested

                with tempfile.TemporaryDirectory() as temp_dir:
                    audit_log = AuditLog(
                        src=src,
                        dst=dst,
                        formatting=self.settings,
                        quote_language=self.quote_language,
                        output_format=self.output_format,
                        options=self.options,
                    )
                    working_input = src
                    input_needs_conversion = needs_conversion(src, ".docx")
                    output_needs_conversion = self.output_format == ".odt"

                    if input_needs_conversion:
                        if self.cancel_requested:
                            self.cancelled.emit()
                            return
                        progress_callback(2, "Converting input to DOCX")
                        log_callback(f"Converting input to DOCX: {src}")
                        working_input = convert_with_libreoffice(src, temp_dir, ".docx")

                    cleaned_docx = dst if self.output_format == ".docx" else os.path.join(
                        temp_dir, f"{os.path.splitext(os.path.basename(dst))[0]}.docx"
                    )

                    def cleaning_progress(percent, label):
                        if input_needs_conversion and output_needs_conversion:
                            weighted_percent = 5 + int(percent * 0.85)
                        elif input_needs_conversion and not output_needs_conversion:
                            weighted_percent = 5 + int(percent * 0.95)
                        elif not input_needs_conversion and output_needs_conversion:
                            weighted_percent = int(percent * 0.90)
                        else:
                            weighted_percent = percent
                        progress_callback(weighted_percent, label)

                    process_docx(
                        src=working_input,
                        dst=cleaned_docx,
                        do_deframe=self.options["deframe"],
                        do_spacing=self.options["spacing"],
                        do_blanks=self.options["blanks"],
                        do_breaks=self.options["breaks"],
                        do_indents=self.options["indents"],
                        do_unify=self.options["unify"],
                        do_sentfix=self.options["sentfix"],
                        do_quote_uniform=self.options["quote_uniform"],
                        quote_language=self.quote_language,
                        settings=self.settings,
                        log=log_callback,
                        progress_callback=cleaning_progress,
                        should_cancel=should_cancel,
                        audit_log=audit_log,
                    )

                    if self.output_format == ".odt":
                        if self.cancel_requested:
                            self.cancelled.emit()
                            return
                        progress_callback(93, "Converting cleaned file to ODT")
                        log_callback(f"Converting cleaned DOCX to ODT: {dst}")
                        converted_output = convert_with_libreoffice(cleaned_docx, os.path.dirname(dst), ".odt")
                        if os.path.normcase(converted_output) != os.path.normcase(dst):
                            if os.path.exists(dst):
                                os.remove(dst)
                            os.replace(converted_output, dst)
                        progress_callback(100, "Finished")

                    audit_path = audit_log.save(Path(dst).with_suffix(".audit.md"))
                    log_callback(f"Audit log saved to: {audit_path}")

            self.overall_progress.emit(100)
            self.finished_ok.emit()
        except CleaningCancelled:
            self.cancelled.emit()
        except Exception as exc:
            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.active_file_path = None
        self.settings = QSettings(settings_file_path(), QSettings.IniFormat)
        self.setWindowTitle(f"ScanSweep {APP_VERSION}")
        self.setWindowIcon(QIcon(str(Path(__file__).with_name("app_icon.svg"))))
        self.resize(1020, 760)
        self._build_ui()
        self.load_settings()

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("centralPanel")
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(20, 20, 20, 14)
        root_layout.setSpacing(16)

        header_row = QHBoxLayout()
        header_row.setSpacing(12)

        header_layout = QVBoxLayout()
        header_layout.setSpacing(3)
        title = QLabel("ScanSweep")
        title.setObjectName("appTitle")
        subtitle = QLabel("Clean PDF-converted DOCX and ODT documents.")
        subtitle.setObjectName("appSubtitle")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        header_row.addLayout(header_layout, 1)

        self.theme_switch = ThemeSwitchButton()
        self.theme_switch.theme_changed.connect(self.apply_theme)
        header_row.addWidget(self.theme_switch, 0, Qt.AlignTop | Qt.AlignRight)
        root_layout.addLayout(header_row)

        top_layout = QHBoxLayout()
        top_layout.setSpacing(16)
        root_layout.addLayout(top_layout, 1)

        left_column = QVBoxLayout()
        left_column.setSpacing(16)
        top_layout.addLayout(left_column, 3)

        right_column = QVBoxLayout()
        right_column.setSpacing(16)
        top_layout.addLayout(right_column, 2)

        files_group = QGroupBox("Files")
        files_layout = QVBoxLayout(files_group)
        files_layout.setSpacing(12)

        file_toolbar = QHBoxLayout()
        self.add_button = QPushButton("Add Files")
        self.add_button.clicked.connect(self.browse_files)
        file_toolbar.addWidget(self.add_button)

        self.remove_button = QPushButton("Remove Selected")
        self.remove_button.clicked.connect(self.remove_selected_files)
        file_toolbar.addWidget(self.remove_button)

        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_files)
        file_toolbar.addWidget(self.clear_button)

        file_toolbar.addStretch(1)

        self.batch_checkbox = QCheckBox("Batch mode")
        file_toolbar.addWidget(self.batch_checkbox)
        files_layout.addLayout(file_toolbar)

        self.file_list = FileListWidget()
        self.file_list.files_dropped.connect(self.add_files)
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list.setAlternatingRowColors(True)
        self.file_list.setMinimumHeight(200)
        files_layout.addWidget(self.file_list)

        hint = QLabel("Add or drag .docx/.odt files here. In single-file mode only the first item is used.")
        hint.setObjectName("hintLabel")
        hint.setWordWrap(True)
        files_layout.addWidget(hint)
        left_column.addWidget(files_group, 1)

        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setSpacing(12)

        self.current_file_label = QLabel("Current file: -")
        self.current_file_label.setWordWrap(True)
        progress_layout.addWidget(self.current_file_label)

        self.file_stage_label = QLabel("Current stage: -")
        progress_layout.addWidget(self.file_stage_label)

        self.file_progress = QProgressBar()
        self.file_progress.setRange(0, 100)
        self.file_progress.setFormat("%p%")
        self.file_progress.setMinimumHeight(18)
        progress_layout.addWidget(self.file_progress)

        self.overall_label = QLabel("Overall progress: -")
        progress_layout.addWidget(self.overall_label)

        self.overall_progress = QProgressBar()
        self.overall_progress.setRange(0, 100)
        self.overall_progress.setFormat("%p%")
        self.overall_progress.setMinimumHeight(18)
        progress_layout.addWidget(self.overall_progress)
        left_column.addWidget(progress_group)

        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(220)
        log_layout.addWidget(self.log_box)
        left_column.addWidget(log_group, 2)

        settings_group = QGroupBox("Settings")
        settings_layout = QVBoxLayout(settings_group)
        settings_layout.setSpacing(12)

        font_row = QHBoxLayout()
        font_label = QLabel("Font")
        font_label.setMinimumWidth(60)
        font_row.addWidget(font_label)
        self.font_combo = QFontComboBox()
        self.font_combo.setWritingSystem(QFontDatabase.WritingSystem.Latin)
        self.font_combo.setFontFilters(QFontComboBox.FontFilter.ScalableFonts)
        self.font_combo.setMaxVisibleItems(14)
        self.font_combo.setMinimumHeight(34)
        font_row.addWidget(self.font_combo, 1)
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(6, 72)
        self.font_size_spin.setSuffix(" pt")
        self.font_size_spin.setMinimumHeight(34)
        font_row.addWidget(self.font_size_spin)
        settings_layout.addLayout(font_row)

        metrics_row = QHBoxLayout()
        spacing_label = QLabel("Spacing")
        spacing_label.setMinimumWidth(60)
        metrics_row.addWidget(spacing_label)
        self.line_spacing_spin = QDoubleSpinBox()
        self.line_spacing_spin.setRange(0.5, 3.0)
        self.line_spacing_spin.setSingleStep(0.05)
        self.line_spacing_spin.setDecimals(2)
        self.line_spacing_spin.setMinimumHeight(34)
        metrics_row.addWidget(self.line_spacing_spin)
        metrics_row.addSpacing(20)
        metrics_row.addWidget(QLabel("First line"))
        self.indent_spin = QDoubleSpinBox()
        self.indent_spin.setRange(0.0, 5.0)
        self.indent_spin.setSingleStep(0.25)
        self.indent_spin.setDecimals(2)
        self.indent_spin.setSuffix(" cm")
        self.indent_spin.setMinimumHeight(34)
        metrics_row.addWidget(self.indent_spin)
        metrics_row.addStretch(1)
        settings_layout.addLayout(metrics_row)

        output_row = QHBoxLayout()
        output_label = QLabel("Output")
        output_label.setMinimumWidth(60)
        output_row.addWidget(output_label)
        self.output_group = QButtonGroup(self)
        for extension in OUTPUT_FORMATS:
            button = QRadioButton(extension)
            button.setProperty("outputFormat", extension)
            self.output_group.addButton(button)
            output_row.addWidget(button)
            output_row.addSpacing(18)
        output_row.addStretch(1)
        settings_layout.addLayout(output_row)

        quote_label = QLabel("Quotes")
        quote_label.setMinimumWidth(60)

        self.quote_group = QButtonGroup(self)
        # A serif face draws curly quotes as distinct comma shapes; the interface
        # sans renders the opening and closing pair as near-identical strokes.
        marks_font = QFont("Georgia")
        marks_font.setPointSize(22)
        quotes_layout = QHBoxLayout()
        quotes_layout.setSpacing(22)
        quotes_layout.addWidget(quote_label)
        for language in QUOTE_LANGUAGES:
            button = QRadioButton(quote_example(language, " "))  # thin space keeps the pair tight but legible
            button.setFont(marks_font)
            button.setToolTip(language)
            button.setProperty("quoteLanguage", language)
            # The low opening mark drops below the baseline and would otherwise
            # crowd the row beneath it.
            button.setMinimumHeight(54)
            self.quote_group.addButton(button)
            quotes_layout.addWidget(button)
        quotes_layout.addStretch(1)
        settings_layout.addLayout(quotes_layout)

        right_column.addWidget(settings_group)

        options_group = QGroupBox("Cleanup Rules")
        options_layout = QVBoxLayout(options_group)
        options_layout.setSpacing(10)

        self.deframe_checkbox = QCheckBox("Flatten PDF page layout")
        self.deframe_checkbox.setChecked(True)
        self.deframe_checkbox.setToolTip(
            "Release paragraphs that a PDF conversion pinned to fixed page coordinates, "
            "so the text can flow and be edited."
        )
        options_layout.addWidget(self.deframe_checkbox)

        self.spacing_checkbox = QCheckBox("Spacing, punctuation, quotes, ligatures")
        self.spacing_checkbox.setChecked(True)
        options_layout.addWidget(self.spacing_checkbox)

        self.blanks_checkbox = QCheckBox("Delete blank rows")
        self.blanks_checkbox.setChecked(True)
        options_layout.addWidget(self.blanks_checkbox)

        self.breaks_checkbox = QCheckBox("Remove breaks")
        self.breaks_checkbox.setChecked(True)
        options_layout.addWidget(self.breaks_checkbox)

        self.indents_checkbox = QCheckBox("Reset indents")
        self.indents_checkbox.setChecked(True)
        options_layout.addWidget(self.indents_checkbox)

        self.unify_checkbox = QCheckBox("Unify body text")
        self.unify_checkbox.setChecked(True)
        options_layout.addWidget(self.unify_checkbox)

        self.sentfix_checkbox = QCheckBox("Fix broken sentences")
        self.sentfix_checkbox.setChecked(True)
        options_layout.addWidget(self.sentfix_checkbox)

        self.quote_uniform_checkbox = QCheckBox("Uniform quotes at the end")
        self.quote_uniform_checkbox.setChecked(True)
        options_layout.addWidget(self.quote_uniform_checkbox)

        self.slash_checkbox = QCheckBox("Close spaces around slashes (i / ili → i/ili)")
        options_layout.addWidget(self.slash_checkbox)

        self.legal_checkbox = QCheckBox(
            "Keep legal numbering on its own line (Article 1, § 2, (3), 1.1)"
        )
        options_layout.addWidget(self.legal_checkbox)

        options_note = QLabel(
            "Sentence merging stays conservative: headings, lists and title-like lines are protected. "
            "Leave page layout flattening on for documents converted from PDF; without it the text "
            "stays locked to page coordinates and can pile up once breaks are removed."
        )
        options_note.setWordWrap(True)
        options_note.setObjectName("noteLabel")
        options_layout.addWidget(options_note)
        right_column.addWidget(options_group)

        action_group = QGroupBox("Run")
        action_layout = QVBoxLayout(action_group)
        action_layout.setSpacing(10)

        output_note = QLabel(
            "ODT input and output require LibreOffice. DOCX cleanup runs internally before optional conversion."
        )
        output_note.setWordWrap(True)
        output_note.setObjectName("noteLabel")
        action_layout.addWidget(output_note)

        button_row = QHBoxLayout()
        self.run_button = QPushButton("Run Cleaner")
        self.run_button.setMinimumHeight(44)
        self.run_button.clicked.connect(self.run_cleaner)
        button_row.addWidget(self.run_button)

        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setMinimumHeight(44)
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_cleaner)
        button_row.addWidget(self.cancel_button)
        action_layout.addLayout(button_row)
        right_column.addWidget(action_group)
        right_column.addStretch(1)

        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        self.status_message = QLabel("Ready")
        status_bar.addPermanentWidget(self.status_message, 1)
        self.theme_switch.set_theme("dark")
        self.apply_theme("dark")

    def quote_language(self):
        button = self.quote_group.checkedButton()
        if button is None:
            return QUOTE_LANGUAGES[0]
        return button.property("quoteLanguage")

    def set_quote_language(self, language):
        buttons = self.quote_group.buttons()
        for button in buttons:
            if button.property("quoteLanguage") == language:
                button.setChecked(True)
                return
        buttons[0].setChecked(True)

    def formatting(self):
        return {
            "font_name": self.font_combo.currentFont().family(),
            "font_size": self.font_size_spin.value(),
            "line_spacing": self.line_spacing_spin.value(),
            "first_line_indent_cm": self.indent_spin.value(),
            "close_slash_spacing": self.slash_checkbox.isChecked(),
            "protect_legal_numbering": self.legal_checkbox.isChecked(),
        }

    def set_formatting(self, values):
        self.font_combo.setCurrentFont(QFont(values["font_name"]))
        self.font_size_spin.setValue(int(values["font_size"]))
        self.line_spacing_spin.setValue(float(values["line_spacing"]))
        self.indent_spin.setValue(float(values["first_line_indent_cm"]))
        self.slash_checkbox.setChecked(bool(values["close_slash_spacing"]))
        self.legal_checkbox.setChecked(bool(values["protect_legal_numbering"]))

    def stored_formatting(self):
        """Saved settings, or the retired preset they were last used with."""
        if self.settings.value("font_name") is None:
            return RETIRED_PROFILES.get(
                self.settings.value("profile", "academic"), dict(DEFAULT_FORMATTING)
            )
        return {
            "font_name": self.settings.value("font_name", DEFAULT_FORMATTING["font_name"]),
            "font_size": self.settings.value("font_size", DEFAULT_FORMATTING["font_size"], type=int),
            "line_spacing": self.settings.value("line_spacing", DEFAULT_FORMATTING["line_spacing"], type=float),
            "first_line_indent_cm": self.settings.value(
                "first_line_indent_cm", DEFAULT_FORMATTING["first_line_indent_cm"], type=float
            ),
            "close_slash_spacing": self.settings.value("close_slash_spacing", False, type=bool),
            "protect_legal_numbering": self.settings.value("protect_legal_numbering", False, type=bool),
        }

    def output_format(self):
        button = self.output_group.checkedButton()
        if button is None:
            return OUTPUT_FORMATS[0]
        return button.property("outputFormat")

    def set_output_format(self, extension):
        buttons = self.output_group.buttons()
        for button in buttons:
            if button.property("outputFormat") == extension:
                button.setChecked(True)
                return
        buttons[0].setChecked(True)

    def load_settings(self):
        self.batch_checkbox.setChecked(self.settings.value("batch_mode", False, type=bool))
        self.set_formatting(self.stored_formatting())
        self.set_output_format(self.settings.value("output_format", ".docx"))
        self.theme_switch.set_theme(self.settings.value("theme", "dark"))
        stored_quote_language = self.settings.value("quote_language", "serbian")
        quote_aliases = {
            "english": "english-double",
        }
        self.set_quote_language(quote_aliases.get(stored_quote_language, stored_quote_language))
        self.deframe_checkbox.setChecked(self.settings.value("deframe", True, type=bool))
        self.spacing_checkbox.setChecked(self.settings.value("spacing", True, type=bool))
        self.blanks_checkbox.setChecked(self.settings.value("blanks", True, type=bool))
        self.breaks_checkbox.setChecked(self.settings.value("breaks", True, type=bool))
        self.indents_checkbox.setChecked(self.settings.value("indents", True, type=bool))
        self.unify_checkbox.setChecked(self.settings.value("unify", True, type=bool))
        self.sentfix_checkbox.setChecked(self.settings.value("sentfix", True, type=bool))
        self.quote_uniform_checkbox.setChecked(self.settings.value("quote_uniform", True, type=bool))

    def save_settings(self):
        self.settings.setValue("batch_mode", self.batch_checkbox.isChecked())
        for key, value in self.formatting().items():
            self.settings.setValue(key, value)
        self.settings.setValue("output_format", self.output_format())
        self.settings.setValue("theme", self.theme_switch.theme_name())
        self.settings.setValue("quote_language", self.quote_language())
        self.settings.setValue("deframe", self.deframe_checkbox.isChecked())
        self.settings.setValue("spacing", self.spacing_checkbox.isChecked())
        self.settings.setValue("blanks", self.blanks_checkbox.isChecked())
        self.settings.setValue("breaks", self.breaks_checkbox.isChecked())
        self.settings.setValue("indents", self.indents_checkbox.isChecked())
        self.settings.setValue("unify", self.unify_checkbox.isChecked())
        self.settings.setValue("sentfix", self.sentfix_checkbox.isChecked())
        self.settings.setValue("quote_uniform", self.quote_uniform_checkbox.isChecked())
        self.settings.sync()

    def show_message_box(self, icon, title, text):
        box = QMessageBox(self)
        box.setIcon(icon)
        box.setWindowTitle(title)
        box.setText(text)
        box.setStandardButtons(QMessageBox.Ok)
        box.setWindowIcon(QIcon(str(Path(__file__).with_name("app_icon.svg"))))
        if self.theme_switch.theme_name() == "light":
            box.setStyleSheet(MESSAGE_BOX_LIGHT_THEME)
        else:
            box.setStyleSheet(MESSAGE_BOX_DARK_THEME)
        return box.exec()

    def apply_theme(self, theme_name):
        if self.theme_switch.theme_name() != theme_name:
            self.theme_switch.set_theme(theme_name)
        self.setStyleSheet(theme_stylesheet(LIGHT_THEME if theme_name == "light" else DARK_THEME))
        if self.active_file_path:
            self.highlight_active_file(self.active_file_path)
        else:
            self.clear_active_highlight()

    def add_files(self, paths):
        existing = {self.file_list.item(i).data(Qt.UserRole) for i in range(self.file_list.count())}
        for path in paths:
            if path in existing:
                continue
            item = QListWidgetItem(os.path.basename(path))
            item.setToolTip(path)
            item.setData(Qt.UserRole, path)
            self.file_list.addItem(item)

    def browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Choose DOCX or ODT files",
            "",
            "Documents (*.docx *.odt)",
        )
        if paths:
            self.add_files(paths)

    def remove_selected_files(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def clear_files(self):
        self.file_list.clear()

    def append_log(self, message):
        self.log_box.appendPlainText(message)

    def set_running_state(self, running):
        self.run_button.setEnabled(not running)
        self.cancel_button.setEnabled(running)
        self.add_button.setEnabled(not running)
        self.remove_button.setEnabled(not running)
        self.clear_button.setEnabled(not running)
        self.file_list.setEnabled(not running)
        self.batch_checkbox.setEnabled(not running)
        self.font_combo.setEnabled(not running)
        self.font_size_spin.setEnabled(not running)
        self.line_spacing_spin.setEnabled(not running)
        self.indent_spin.setEnabled(not running)
        self.slash_checkbox.setEnabled(not running)
        self.legal_checkbox.setEnabled(not running)
        for button in self.output_group.buttons():
            button.setEnabled(not running)
        self.theme_switch.setEnabled(not running)
        for button in self.quote_group.buttons():
            button.setEnabled(not running)
        self.deframe_checkbox.setEnabled(not running)
        self.spacing_checkbox.setEnabled(not running)
        self.blanks_checkbox.setEnabled(not running)
        self.breaks_checkbox.setEnabled(not running)
        self.indents_checkbox.setEnabled(not running)
        self.unify_checkbox.setEnabled(not running)
        self.sentfix_checkbox.setEnabled(not running)
        self.quote_uniform_checkbox.setEnabled(not running)

    def collect_sources(self):
        paths = [self.file_list.item(i).data(Qt.UserRole) for i in range(self.file_list.count())]
        if not self.batch_checkbox.isChecked() and paths:
            return [paths[0]]
        return paths

    def highlight_active_file(self, path):
        self.active_file_path = path
        is_light = self.theme_switch.theme_name() == "light"
        active_bg = QColor("#cfe9ff") if is_light else QColor("#0e6b80")
        active_fg = QColor("#102136") if is_light else QColor("#ffffff")
        inactive_fg = QColor("#1d2433") if is_light else QColor("#ffffff")
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            item_path = item.data(Qt.UserRole)
            is_active = item_path == path
            item.setSelected(is_active)
            item.setBackground(active_bg if is_active else Qt.transparent)
            item.setForeground(active_fg if is_active else inactive_fg)
            if is_active:
                self.file_list.scrollToItem(item)

    def clear_active_highlight(self):
        self.active_file_path = None
        inactive_fg = QColor("#1d2433") if self.theme_switch.theme_name() == "light" else QColor("#ffffff")
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            item.setBackground(Qt.transparent)
            item.setForeground(inactive_fg)
            item.setSelected(False)

    def run_cleaner(self):
        sources = self.collect_sources()
        if not sources:
            self.show_message_box(QMessageBox.Warning, "No file", "Choose at least one DOCX or ODT file first.")
            return

        selected_format = self.output_format()
        if self.batch_checkbox.isChecked():
            output_dir = QFileDialog.getExistingDirectory(self, "Choose output folder for cleaned files")
            if not output_dir:
                return
            output_file = None
        else:
            first_name = os.path.splitext(os.path.basename(sources[0]))[0] + f"_cleaned{selected_format}"
            output_file, _ = QFileDialog.getSaveFileName(
                self,
                "Save cleaned file as",
                first_name,
                "Word documents (*.docx);;OpenDocument Text (*.odt)",
            )
            if not output_file:
                return
            if not output_file.lower().endswith(selected_format):
                output_file += selected_format
            output_dir = None

        self.save_settings()
        self.log_box.clear()
        self.file_progress.setValue(0)
        self.overall_progress.setValue(0)
        self.file_stage_label.setText("Current stage: Starting...")
        self.overall_label.setText("Overall progress: 0%")
        self.status_message.setText("Running cleaner")
        self.set_running_state(True)

        options = {
            "deframe": self.deframe_checkbox.isChecked(),
            "spacing": self.spacing_checkbox.isChecked(),
            "blanks": self.blanks_checkbox.isChecked(),
            "breaks": self.breaks_checkbox.isChecked(),
            "indents": self.indents_checkbox.isChecked(),
            "unify": self.unify_checkbox.isChecked(),
            "sentfix": self.sentfix_checkbox.isChecked(),
            "quote_uniform": self.quote_uniform_checkbox.isChecked(),
        }

        self.worker = CleanerWorker(
            sources=sources,
            batch_mode=self.batch_checkbox.isChecked(),
            settings=self.formatting(),
            quote_language=self.quote_language(),
            output_format=selected_format,
            options=options,
            output_dir=output_dir,
            output_file=output_file,
        )
        self.worker.log_message.connect(self.append_log)
        self.worker.file_started.connect(self.on_file_started)
        self.worker.file_progress.connect(self.on_file_progress)
        self.worker.overall_progress.connect(self.on_overall_progress)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.cancelled.connect(self.on_cancelled)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def cancel_cleaner(self):
        if self.worker is not None:
            self.worker.request_cancel()
            self.file_stage_label.setText("Current stage: Cancelling...")
            self.status_message.setText("Cancelling...")
            self.append_log("Cancellation requested.")

    def on_file_started(self, index, total, src):
        self.highlight_active_file(src)
        self.current_file_label.setText(f"Current file: {index}/{total} - {os.path.basename(src)}")
        self.current_file_label.setToolTip(src)
        self.file_stage_label.setText("Current stage: Opening file")
        self.file_progress.setValue(0)
        self.status_message.setText(f"Processing {os.path.basename(src)}")

    def on_file_progress(self, percent, label):
        self.file_progress.setValue(percent)
        self.file_stage_label.setText(f"Current stage: {label} ({percent}%)")
        self.status_message.setText(label)

    def on_overall_progress(self, percent):
        self.overall_progress.setValue(percent)
        self.overall_label.setText(f"Overall progress: {percent}%")

    def on_finished(self):
        self.set_running_state(False)
        self.clear_active_highlight()
        self.file_stage_label.setText("Current stage: Finished")
        self.status_message.setText("Finished")
        self.show_message_box(QMessageBox.Information, "Done", "Cleaning finished. See log for details.")
        self.worker = None

    def on_cancelled(self):
        self.set_running_state(False)
        self.clear_active_highlight()
        self.file_stage_label.setText("Current stage: Cancelled")
        self.status_message.setText("Cancelled")
        self.show_message_box(QMessageBox.Information, "Cancelled", "Cleaning was cancelled.")
        self.worker = None

    def on_failed(self, error_message):
        self.set_running_state(False)
        self.clear_active_highlight()
        self.status_message.setText("Error")
        self.show_message_box(QMessageBox.Critical, "Error", error_message)
        self.worker = None

    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("ScanSweep")
    app.setApplicationVersion(APP_VERSION)
    app.setWindowIcon(QIcon(str(Path(__file__).with_name("app_icon.svg"))))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
