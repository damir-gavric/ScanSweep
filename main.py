import os
import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QRectF, QSize, QSettings, QThread, Qt, Signal
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont, QIcon, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
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
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from conversion import convert_with_libreoffice, needs_conversion
from audit_log import AuditLog
from processor import CleaningCancelled, process_docx


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
QListWidget, QPlainTextEdit, QComboBox {
    border: 1px solid #404040;
    border-radius: 8px;
    padding: 6px;
    background-color: #222222;
    color: #f4f4f4;
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
QLabel, QCheckBox, QGroupBox {
    background: transparent;
    color: #f4f4f4;
}
QCheckBox {
    spacing: 8px;
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
QListWidget, QPlainTextEdit, QComboBox {
    border: 1px solid #c5d0e0;
    border-radius: 8px;
    padding: 6px;
    background-color: #ffffff;
    color: #1d2433;
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
QLabel, QCheckBox, QGroupBox {
    background: transparent;
    color: #1d2433;
}
QCheckBox {
    spacing: 8px;
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
        profile_name,
        quote_language,
        output_format,
        options,
        output_dir=None,
        output_file=None,
    ):
        super().__init__()
        self.sources = sources
        self.batch_mode = batch_mode
        self.profile_name = profile_name
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
                        profile_name=self.profile_name,
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
                        do_spacing=self.options["spacing"],
                        do_blanks=self.options["blanks"],
                        do_breaks=self.options["breaks"],
                        do_indents=self.options["indents"],
                        do_unify=self.options["unify"],
                        do_sentfix=self.options["sentfix"],
                        do_quote_uniform=self.options["quote_uniform"],
                        quote_language=self.quote_language,
                        profile_name=self.profile_name,
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
        self.settings = QSettings("CleanDOCX", "ScanSweep")
        self.setWindowTitle("ScanSweep")
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

        progress_divider = QFrame()
        progress_divider.setFrameShape(QFrame.HLine)
        progress_divider.setFrameShadow(QFrame.Sunken)
        progress_layout.addWidget(progress_divider)

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

        profile_row = QHBoxLayout()
        profile_label = QLabel("Profile")
        profile_label.setMinimumWidth(60)
        profile_row.addWidget(profile_label)
        self.profile_combo = QComboBox()
        self.profile_combo.addItems(["novel", "academic", "legal"])
        self.profile_combo.setCurrentText("academic")
        self.profile_combo.setMinimumHeight(34)
        profile_row.addWidget(self.profile_combo, 1)
        self.profile_info_button = self._create_info_button()
        self.profile_info_button.clicked.connect(self.show_profile_info)
        profile_row.addWidget(self.profile_info_button)
        settings_layout.addLayout(profile_row)

        output_row = QHBoxLayout()
        output_label = QLabel("Output")
        output_label.setMinimumWidth(60)
        output_row.addWidget(output_label)
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItems([".docx", ".odt"])
        self.output_format_combo.setCurrentText(".docx")
        self.output_format_combo.setMinimumHeight(34)
        output_row.addWidget(self.output_format_combo, 1)
        settings_layout.addLayout(output_row)

        quote_row = QHBoxLayout()
        quote_label = QLabel("Quotes")
        quote_label.setMinimumWidth(60)
        quote_row.addWidget(quote_label)
        self.quote_language_combo = QComboBox()
        self.quote_language_combo.addItems(["english-double", "english-single", "serbian", "german"])
        self.quote_language_combo.setCurrentText("serbian")
        self.quote_language_combo.setMinimumHeight(34)
        quote_row.addWidget(self.quote_language_combo, 1)
        self.quote_info_button = self._create_info_button()
        self.quote_info_button.clicked.connect(self.show_quote_info)
        quote_row.addWidget(self.quote_info_button)
        settings_layout.addLayout(quote_row)

        right_column.addWidget(settings_group)

        options_group = QGroupBox("Cleanup Rules")
        options_layout = QVBoxLayout(options_group)
        options_layout.setSpacing(10)

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

        self.unify_checkbox = QCheckBox("Unify body text using selected profile")
        self.unify_checkbox.setChecked(True)
        options_layout.addWidget(self.unify_checkbox)

        self.sentfix_checkbox = QCheckBox("Fix broken sentences")
        self.sentfix_checkbox.setChecked(True)
        options_layout.addWidget(self.sentfix_checkbox)

        self.quote_uniform_checkbox = QCheckBox("Uniform quotes at the end")
        self.quote_uniform_checkbox.setChecked(True)
        options_layout.addWidget(self.quote_uniform_checkbox)

        options_note = QLabel(
            "Sentence merging stays conservative: headings, lists and title-like lines are protected."
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

    def _create_info_button(self):
        button = QPushButton()
        button.setProperty("infoButton", True)
        button.setCursor(Qt.PointingHandCursor)
        button.setToolTip("More information")
        button.setIcon(QIcon(str(Path(__file__).with_name("info_icon.svg"))))
        button.setIconSize(button.sizeHint())
        button.setText("")
        return button

    def load_settings(self):
        self.batch_checkbox.setChecked(self.settings.value("batch_mode", False, type=bool))
        stored_profile = self.settings.value("profile", "academic")
        profile_aliases = {
            "roman": "novel",
            "strucni_rad": "academic",
            "pravni_tekst": "legal",
        }
        self.profile_combo.setCurrentText(profile_aliases.get(stored_profile, stored_profile))
        self.output_format_combo.setCurrentText(self.settings.value("output_format", ".docx"))
        self.theme_switch.set_theme(self.settings.value("theme", "dark"))
        stored_quote_language = self.settings.value("quote_language", "serbian")
        quote_aliases = {
            "english": "english-double",
        }
        self.quote_language_combo.setCurrentText(quote_aliases.get(stored_quote_language, stored_quote_language))
        self.spacing_checkbox.setChecked(self.settings.value("spacing", True, type=bool))
        self.blanks_checkbox.setChecked(self.settings.value("blanks", True, type=bool))
        self.breaks_checkbox.setChecked(self.settings.value("breaks", True, type=bool))
        self.indents_checkbox.setChecked(self.settings.value("indents", True, type=bool))
        self.unify_checkbox.setChecked(self.settings.value("unify", True, type=bool))
        self.sentfix_checkbox.setChecked(self.settings.value("sentfix", True, type=bool))
        self.quote_uniform_checkbox.setChecked(self.settings.value("quote_uniform", True, type=bool))

    def save_settings(self):
        self.settings.setValue("batch_mode", self.batch_checkbox.isChecked())
        self.settings.setValue("profile", self.profile_combo.currentText())
        self.settings.setValue("output_format", self.output_format_combo.currentText())
        self.settings.setValue("theme", self.theme_switch.theme_name())
        self.settings.setValue("quote_language", self.quote_language_combo.currentText())
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

    def show_profile_info(self):
        self.show_message_box(
            QMessageBox.Information,
            "Profile Info",
            (
                "Profiles define the cleanup style and final body-text formatting.\n\n"
                "Novel\n"
                "- Garamond 12, first-line indent 1 cm\n"
                "- Better for book-like prose and dialogue-heavy text\n"
                "- Keeps slash spacing less aggressive\n\n"
                "Academic\n"
                "- Arial 11, first-line indent 1 cm\n"
                "- Best general-purpose profile for reports and articles\n"
                "- Applies stricter slash normalization such as i / ili -> i/ili\n\n"
                "Legal\n"
                "- Times New Roman 12, no first-line indent\n"
                "- Better for structured legal text and numbered clauses\n"
                "- Adds stronger protection for Article/Section/Clause patterns"
            ),
        )

    def show_quote_info(self):
        self.show_message_box(
            QMessageBox.Information,
            "Quote Style Info",
            (
                "This option runs at the end of processing and converts quotes to one consistent language-specific style.\n\n"
                "English double\n"
                '- "Proxima Centauri"\n\n'
                "English single\n"
                "- 'Proxima Centauri'\n\n"
                "Serbian\n"
                '- „Proxima Centauri”\n\n'
                "German\n"
                '- „Proxima Centauri“\n\n'
                "Use it when you want the final document to have uniform quotation marks."
            ),
        )

    def apply_theme(self, theme_name):
        if self.theme_switch.theme_name() != theme_name:
            self.theme_switch.set_theme(theme_name)
        self.setStyleSheet(LIGHT_THEME if theme_name == "light" else DARK_THEME)
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
        self.profile_combo.setEnabled(not running)
        self.output_format_combo.setEnabled(not running)
        self.theme_switch.setEnabled(not running)
        self.quote_language_combo.setEnabled(not running)
        self.profile_info_button.setEnabled(not running)
        self.quote_info_button.setEnabled(not running)
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

        selected_format = self.output_format_combo.currentText()
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
            profile_name=self.profile_combo.currentText(),
            quote_language=self.quote_language_combo.currentText(),
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
    app.setWindowIcon(QIcon(str(Path(__file__).with_name("app_icon.svg"))))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
