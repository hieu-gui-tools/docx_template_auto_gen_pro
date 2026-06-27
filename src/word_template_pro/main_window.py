"""Main Window cho Word Template Pro."""
from __future__ import annotations

import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt, QThread, Signal, QByteArray
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QFrame,
    QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPushButton,
    QScrollArea, QSizePolicy, QSplitter,
    QStatusBar, QToolButton, QVBoxLayout, QWidget,
)

from .database import Template, get_session, get_setting, set_setting, init_db
from .field_widgets import make_field_widget, ScriptField
from .template_engine import (
    extract_fields, generate_document,
    load_script_module, get_script_args, call_script_func,
)
from .template_manager_dialog import TemplateManagerDialog
from .batch_dialog import BatchDialog

# ── Global stylesheet ─────────────────────────────────────────────────────────
MAIN_STYLE = """
QMainWindow, QWidget#centralWidget {
    background: #0d1020;
}
/* Sidebar */
QWidget#sidebar {
    background: #111528;
    border-right: 1.5px solid #1e2540;
}
/* Top bar */
QWidget#topbar {
    background: #131828;
    border-bottom: 1.5px solid #1e2540;
}
/* Content area */
QScrollArea, QWidget#formArea {
    background: #0d1020;
    border: none;
}
/* Cards */
QFrame#card {
    background: #141929;
    border: 1.5px solid #1e2845;
    border-radius: 10px;
}
/* Field rows */
QFrame#fieldRow {
    background: #181e32;
    border: 1px solid #222840;
    border-radius: 8px;
}
/* Labels */
QLabel#sectionTitle {
    color: #5568c8;
    font-size: 11px;
    font-weight: bold;
    letter-spacing: 1px;
}
QLabel#fieldLabel {
    color: #c0cce8;
    font-size: 13px;
    font-weight: 600;
    min-width: 160px;
}
QLabel#typeTag {
    color: #6678a8;
    font-size: 11px;
    background: #1a2038;
    border-radius: 4px;
    padding: 2px 6px;
}
/* Buttons – sidebar */
QPushButton#sideBtn {
    background: transparent;
    color: #8898cc;
    border: none;
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 13px;
    text-align: left;
}
QPushButton#sideBtn:hover { background: #1a2240; color: #aabbff; }
QPushButton#sideBtn:checked { background: #1e2a50; color: #6c8aff; font-weight: bold; }
/* Action buttons */
QPushButton#btnGenerate {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 #2d4aaa, stop:1 #1a3580);
    color: #ddeeff;
    border: none;
    border-radius: 8px;
    padding: 10px 28px;
    font-size: 14px;
    font-weight: bold;
}
QPushButton#btnGenerate:hover {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
        stop:0 #3d5acc, stop:1 #2a45a0);
}
QPushButton#btnGenerate:disabled { background: #1a1f35; color: #445; }

QPushButton#btnBatch {
    background: #1a3028;
    color: #6bffaa;
    border: 1.5px solid #2a5040;
    border-radius: 8px;
    padding: 10px 20px;
    font-size: 14px;
    font-weight: bold;
}
QPushButton#btnBatch:hover { background: #2a5040; }

QPushButton#btnLoad {
    background: #1e2a50;
    color: #aabbff;
    border: 1.5px solid #2d3a70;
    border-radius: 8px;
    padding: 8px 18px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton#btnLoad:hover { background: #2d3a70; }

QPushButton#btnRestart {
    background: #2a1a35;
    color: #cc88ff;
    border: 1.5px solid #4a2a60;
    border-radius: 8px;
    padding: 8px 16px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton#btnRestart:hover { background: #3a2a50; }

QPushButton#btnSettings {
    background: transparent;
    color: #6878a8;
    border: 1.5px solid #2a3050;
    border-radius: 8px;
    padding: 8px 16px;
    font-size: 13px;
}
QPushButton#btnSettings:hover { background: #1a2038; color: #aabbff; }

/* Status bar */
QStatusBar {
    background: #0d1020;
    color: #5568a0;
    font-size: 12px;
    border-top: 1px solid #1a2040;
}
/* Inputs */
QLineEdit, QSpinBox, QDoubleSpinBox, QDateEdit, QDateTimeEdit, QComboBox {
    background: #1e2230;
    color: #e8eaf0;
    border: 1.5px solid #3a3f55;
    border-radius: 6px;
    padding: 5px 8px;
    font-size: 13px;
    min-height: 28px;
}
QTextEdit {
    background: #1e2230;
    color: #e8eaf0;
    border: 1.5px solid #3a3f55;
    border-radius: 6px;
    padding: 8px 8px;
    font-size: 13px;
    min-height: 120px;
}
QLineEdit:focus, QTextEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QDateEdit:focus, QDateTimeEdit:focus, QComboBox:focus {
    border-color: #6c8aff;
}
QLineEdit#pathEdit {
    background: #1a1f35;
    color: #8898cc;
    border: 1.5px solid #2a2f50;
    border-radius: 6px;
    padding: 4px 10px;
    font-size: 12px;
}
"""


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word Template Pro")
        self.resize(1100, 720)
        self.setMinimumSize(860, 560)

        geom_str = get_setting("window_geometry")
        if geom_str:
            self.restoreGeometry(QByteArray.fromBase64(geom_str.encode("utf-8")))

        self._current_template: Template | None = None
        self._fields: list[dict] = []
        self._field_widgets: dict[str, QWidget] = {}   # raw -> widget
        self._script_mod = None

        self.setStyleSheet(MAIN_STYLE)
        self._build_ui()
        self._update_ui_state()

    def check_missing_templates(self):
        from .database import get_session, Template
        missing = []
        with get_session() as s:
            rows = s.query(Template).all()
            for row in rows:
                if not Path(row.file_path).exists():
                    missing.append(row)
                    
        if missing:
            names = "\n".join(f"- {m.name} ({m.file_path})" for m in missing[:5])
            if len(missing) > 5:
                names += f"\n... và {len(missing) - 5} file khác."
                
            ret = QMessageBox.question(
                self, "Phát hiện Template bị mất",
                f"Phát hiện {len(missing)} file template trong danh sách quản lý không còn tồn tại trên ổ cứng (có thể đã bị xóa hoặc di chuyển):\n\n"
                f"{names}\n\n"
                f"Bạn có muốn tự động xóa các template này khỏi danh sách quản lý không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if ret == QMessageBox.StandardButton.Yes:
                with get_session() as s:
                    for m in missing:
                        obj = s.get(Template, m.id)
                        if obj:
                            s.delete(obj)
                    s.commit()
                self.status.showMessage(f"Đã dọn dẹp {len(missing)} template không tồn tại khỏi cơ sở dữ liệu.")

    # ── UI build ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)

        root_h = QHBoxLayout(central)
        root_h.setContentsMargins(0, 0, 0, 0)
        root_h.setSpacing(0)

        # ── Sidebar ──
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(200)
        sb_v = QVBoxLayout(sidebar)
        sb_v.setContentsMargins(12, 20, 12, 20)
        sb_v.setSpacing(4)

        logo = QLabel("📄 Word Template\nPro")
        logo.setStyleSheet("color: #6c8aff; font-size: 15px; font-weight: bold; padding: 6px 4px 16px 4px;")
        sb_v.addWidget(logo)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #1e2540;")
        sb_v.addWidget(sep)
        sb_v.addSpacing(8)

        self.btn_load = QPushButton("📂  Mở Template")
        self.btn_load.setObjectName("sideBtn")
        self.btn_load.clicked.connect(self.action_load_template)
        sb_v.addWidget(self.btn_load)

        self.btn_manager = QPushButton("📋  Danh sách Template")
        self.btn_manager.setObjectName("sideBtn")
        self.btn_manager.clicked.connect(self.action_open_manager)
        sb_v.addWidget(self.btn_manager)

        self.btn_settings_side = QPushButton("⚙️  Cài đặt")
        self.btn_settings_side.setObjectName("sideBtn")
        self.btn_settings_side.clicked.connect(self.action_open_settings)
        sb_v.addWidget(self.btn_settings_side)

        self.btn_guide = QPushButton("📖  Hướng dẫn")
        self.btn_guide.setObjectName("sideBtn")
        self.btn_guide.clicked.connect(self.action_show_guide)
        sb_v.addWidget(self.btn_guide)

        sb_v.addStretch()

        self.btn_restart = QPushButton("🔄  Restart App")
        self.btn_restart.setObjectName("btnRestart")
        self.btn_restart.clicked.connect(self.action_restart)
        sb_v.addWidget(self.btn_restart)

        root_h.addWidget(sidebar)

        # ── Main content ──
        content_v = QVBoxLayout()
        content_v.setContentsMargins(0, 0, 0, 0)
        content_v.setSpacing(0)

        # Top bar
        topbar = QWidget()
        topbar.setObjectName("topbar")
        topbar.setFixedHeight(56)
        top_h = QHBoxLayout(topbar)
        top_h.setContentsMargins(20, 0, 20, 0)

        self.lbl_tpl_name = QLabel("Chưa chọn template")
        self.lbl_tpl_name.setStyleSheet("color: #8898cc; font-size: 14px; font-weight: bold;")
        top_h.addWidget(self.lbl_tpl_name)
        top_h.addStretch()

        self.lbl_field_count = QLabel("")
        self.lbl_field_count.setStyleSheet("color: #5568a0; font-size: 12px;")
        top_h.addWidget(self.lbl_field_count)

        content_v.addWidget(topbar)

        # Scroll area for fields
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setObjectName("scrollArea")
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.form_container = QWidget()
        self.form_container.setObjectName("formArea")
        self.form_layout = QVBoxLayout(self.form_container)
        self.form_layout.setContentsMargins(24, 20, 24, 20)
        self.form_layout.setSpacing(10)
        self.form_layout.addStretch()

        self.scroll.setWidget(self.form_container)
        content_v.addWidget(self.scroll, 1)

        # Bottom action bar
        action_bar = QWidget()
        action_bar.setStyleSheet("background: #111528; border-top: 1.5px solid #1e2540;")
        action_bar.setFixedHeight(68)
        act_h = QHBoxLayout(action_bar)
        act_h.setContentsMargins(24, 12, 24, 12)

        # Output path
        out_lbl = QLabel("📁 Output:")
        out_lbl.setStyleSheet("color: #6878a8; font-size: 12px;")
        self.out_path_edit = QLineEdit()
        self.out_path_edit.setObjectName("pathEdit")
        self.out_path_edit.setPlaceholderText("Thư mục xuất file…")
        self.out_path_edit.setText(get_setting("default_output_folder"))
        self.out_path_edit.setMinimumWidth(220)
        btn_browse_out = QPushButton("⚙️")
        btn_browse_out.setObjectName("btnSettings")
        btn_browse_out.setFixedWidth(64)
        btn_browse_out.clicked.connect(self._browse_output)

        self.btn_open_out = QPushButton("📂")
        self.btn_open_out.setObjectName("btnSettings")
        self.btn_open_out.setFixedWidth(64)
        self.btn_open_out.clicked.connect(self._open_output_dir)

        act_h.addWidget(out_lbl)
        act_h.addWidget(self.out_path_edit)
        act_h.addWidget(btn_browse_out)
        act_h.addWidget(self.btn_open_out)
        act_h.addSpacing(16)

        self.btn_batch = QPushButton("📊  Batch")
        self.btn_batch.setObjectName("btnBatch")
        self.btn_batch.clicked.connect(self.action_batch)
        act_h.addWidget(self.btn_batch)

        self.btn_generate = QPushButton("✨ Tạo file")
        self.btn_generate.setObjectName("btnGenerate")
        self.btn_generate.clicked.connect(self.action_generate)
        act_h.addWidget(self.btn_generate)

        content_v.addWidget(action_bar)

        root_h.addLayout(content_v)

        # Status bar
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Sẵn sàng  •  Mở file template .docx để bắt đầu")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _update_ui_state(self):
        has_tpl = self._current_template is not None
        self.btn_generate.setEnabled(has_tpl)
        self.btn_batch.setEnabled(has_tpl)

    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục xuất file",
            self.out_path_edit.text() or str(Path.home()))
        if folder:
            self.out_path_edit.setText(folder)
            set_setting("default_output_folder", folder)

    def _open_output_dir(self):
        folder = self.out_path_edit.text().strip() or get_setting("default_output_folder")
        if not folder and self._current_template:
            folder = str(Path(self._current_template.file_path).parent / "output")
            
        if folder and Path(folder).exists():
            import os
            try:
                os.startfile(folder)
            except Exception as e:
                QMessageBox.warning(self, "Lỗi", f"Không thể mở thư mục:\n{e}")
        else:
            QMessageBox.information(self, "Thư mục không tồn tại", "Thư mục đầu ra chưa được tạo hoặc không tồn tại.")

    def _clear_form(self):
        while self.form_layout.count() > 1:
            item = self.form_layout.takeAt(0)
            if item and item.widget():
                item.widget().deleteLater()
        self._field_widgets.clear()

    def _build_form(self):
        self._clear_form()
        if not self._fields:
            lbl = QLabel("⚠️  Không tìm thấy trường nào trong template.\n"
                          "Đảm bảo template có trường dạng {{ten_truong|type}}")
            lbl.setStyleSheet("color: #6878a0; font-size: 13px; padding: 40px;")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.form_layout.insertWidget(0, lbl)
            return

        sec_lbl = QLabel("NHẬP LIỆU")
        sec_lbl.setObjectName("sectionTitle")
        self.form_layout.insertWidget(0, sec_lbl)

        all_field_names = {f["name"] for f in self._fields}

        for i, field in enumerate(self._fields):
            name  = field["name"]
            ftype = field["type"].lower()
            raw   = field["raw"]

            row_frame = QFrame()
            row_frame.setObjectName("fieldRow")
            row_h = QHBoxLayout(row_frame)
            row_h.setContentsMargins(14, 10, 14, 10)
            row_h.setSpacing(12)

            # Label + type tag
            lbl_col = QVBoxLayout()
            lbl_col.setSpacing(2)
            f_lbl = QLabel(name)
            f_lbl.setObjectName("fieldLabel")
            type_lbl = QLabel(ftype)
            type_lbl.setObjectName("typeTag")
            lbl_col.addWidget(f_lbl)
            lbl_col.addWidget(type_lbl)
            row_h.addLayout(lbl_col)

            # Widget
            script_args = None
            covered_args = None
            if ftype == "script" and self._script_mod:
                script_args = get_script_args(self._script_mod, name)
                covered_args = [arg for arg in script_args if arg in all_field_names]
            w = make_field_widget(ftype, script_args, covered_args)
            w.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            row_h.addWidget(w)
            self._field_widgets[raw] = w

            self.form_layout.insertWidget(i + 1, row_frame)

        self.lbl_field_count.setText(f"{len(self._fields)} trường")

    # ── Actions ───────────────────────────────────────────────────────────────

    def action_load_template(self, tpl_id: int | None = None):
        if tpl_id is False:
            tpl_id = None

        path: str | None = None

        if tpl_id is None:
            # Mở file dialog bình thường
            chosen, _ = QFileDialog.getOpenFileName(
                self, "Chọn file template", "",
                "Word Documents (*.docx)")
            if not chosen:
                return
            # Chuẩn hoá path (resolve symlink, backslash Windows)
            path = str(Path(chosen).resolve())
        else:
            # Mở từ Manager – lấy path từ DB
            with get_session() as s:
                tpl_row = s.get(Template, tpl_id)
                if tpl_row:
                    path = str(Path(tpl_row.file_path).resolve())

            if not path:
                QMessageBox.warning(self, "Không tìm thấy",
                                    f"Không tìm thấy template ID={tpl_id} trong database.")
                return

            if not Path(path).exists():
                # Cho phép user chỉ đường dẫn mới nếu file đã bị di chuyển
                ret = QMessageBox.question(
                    self, "File không tìm thấy",
                    f"File template không còn ở vị trí cũ:\n{path}\n\n"
                    "Bạn có muốn chọn lại file không?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if ret == QMessageBox.StandardButton.Yes:
                    chosen, _ = QFileDialog.getOpenFileName(
                        self, "Chọn lại file template", str(Path(path).parent),
                        "Word Documents (*.docx)")
                    if not chosen:
                        return
                    path = str(Path(chosen).resolve())
                    # Cập nhật path mới vào DB
                    with get_session() as s:
                        tpl_row = s.get(Template, tpl_id)
                        if tpl_row:
                            tpl_row.file_path = path
                            s.commit()
                else:
                    return

        # ── Đọc fields từ template ──────────────────────────────────────────
        try:
            fields = extract_fields(path)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi đọc template", str(e))
            return

        if not fields:
            QMessageBox.warning(
                self, "Không tìm thấy trường dữ liệu",
                "Template này không có bất kỳ trường dữ liệu nào (ví dụ: {{ten_truong}}).\n\n"
                "Vui lòng mở file Word và thêm các trường trước khi load vào ứng dụng."
            )
            return

        # Kiểm tra tính hợp lệ của tên trường script
        invalid_scripts = [f["name"] for f in fields if f["type"].lower() == "script" and not re.match(r'^[A-Za-z0-9_]+$', f["name"])]
        if invalid_scripts:
            QMessageBox.warning(
                self, "Tên trường Script không hợp lệ",
                f"Phát hiện trường type script có chứa khoảng trắng, dấu tiếng Việt hoặc ký tự đặc biệt:\n"
                f"{', '.join(invalid_scripts)}\n\n"
                f"Tên trường script sẽ được dùng làm tên hàm tính toán, do đó chỉ được chứa chữ cái không dấu, số và dấu gạch dưới (_).\n\n"
                f"Ví dụ: Cần thay đổi 'Họ và tên' thành 'ho_va_ten' hoặc 'Ho_va_ten' trong file Word."
            )

        # Load script module nếu có file .py cùng tên
        script_mod = load_script_module(path)

        # ── Lưu/cập nhật DB ────────────────────────────────────────────────
        # Chuẩn hoá path về dạng tuyệt đối trước khi lưu/query
        norm_path = str(Path(path).resolve())
        with get_session() as s:
            # Query không phân biệt slash style bằng cách so sánh cả hai dạng
            tpl_row = (
                s.query(Template).filter_by(file_path=norm_path).first()
                or s.query(Template).filter_by(file_path=path).first()
            )
            if tpl_row is None:
                default_name = Path(path).stem
                name, ok = QInputDialog.getText(
                    self, "Đặt tên template",
                    "Tên hiển thị cho template này:",
                    text=default_name)
                if not ok or not name.strip():
                    name = default_name
                tpl_row = Template(name=name.strip(), file_path=norm_path)
                s.add(tpl_row)
                s.commit()
                s.refresh(tpl_row)
            else:
                # Cập nhật path về dạng chuẩn nếu cần
                if tpl_row.file_path != norm_path:
                    tpl_row.file_path = norm_path
                    s.commit()
            tpl_id      = tpl_row.id
            tpl_name    = tpl_row.name
            tpl_pattern = tpl_row.filename_pattern or "{template_name}_{date}"

        self._current_template = type("T", (), {
            "id": tpl_id, "name": tpl_name, "file_path": path,
            "filename_pattern": tpl_pattern,
            "output_folder": get_setting("default_output_folder"),
        })()
        self._fields = fields
        self._script_mod = script_mod

        self.lbl_tpl_name.setText(f"📄  {tpl_name}")
        self._build_form()
        self._update_ui_state()
        self.status.showMessage(f"Đã tải: {path}  •  {len(fields)} trường")

    def action_open_manager(self):
        dlg = TemplateManagerDialog(self)
        dlg.template_selected.connect(lambda tid: self.action_load_template(tpl_id=tid))
        dlg.exec()

    def action_open_settings(self):
        """Simple settings dialog."""
        dlg = _SettingsDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.out_path_edit.setText(get_setting("default_output_folder"))

    def action_show_guide(self):
        dlg = _GuideDialog(self)
        dlg.exec()

    def action_generate(self):
        if not self._current_template:
            return

        # Collect values
        replacements: dict[str, str] = {}
        field_values: dict[str, Any] = {}

        # Lấy giá trị các trường không phải script trước
        for field in self._fields:
            name  = field["name"]
            ftype = field["type"].lower()
            raw   = field["raw"]
            w     = self._field_widgets.get(raw)
            if w is None or ftype == "script":
                continue
            value_str = w.get_value() if hasattr(w, "get_value") else ""
            value_typed = w.get_typed_value() if hasattr(w, "get_typed_value") else value_str
            replacements[raw] = value_str
            field_values[name] = value_typed

        # Lấy giá trị các trường script
        for field in self._fields:
            name  = field["name"]
            ftype = field["type"].lower()
            raw   = field["raw"]
            w     = self._field_widgets.get(raw)
            if w is None or ftype != "script":
                continue
            # Call script function
            try:
                kwargs = w.get_kwargs() if hasattr(w, "get_kwargs") else {}
                if self._script_mod:
                    expected_args = get_script_args(self._script_mod, name)
                    for arg in expected_args:
                        if arg in field_values:
                            kwargs[arg] = field_values[arg]
                value = call_script_func(self._script_mod, name, kwargs)
            except Exception as e:
                QMessageBox.critical(self, "Lỗi Script",
                                      f"Lỗi chạy hàm '{name}':\n{e}")
                return
            replacements[raw] = str(value)
            field_values[name] = value

        # Build output filename
        output_dir = self.out_path_edit.text().strip() or get_setting("default_output_folder")
        if not output_dir:
            output_dir = str(Path(self._current_template.file_path).parent / "output")

        pattern = self._current_template.filename_pattern or "{template_name}_{date}"
        now = datetime.now()
        fname = pattern
        # Substitute field values by name
        for field in self._fields:
            fname = fname.replace(f"{{{field['name']}}}", replacements.get(field["raw"], ""))
        fname = fname.replace("{template_name}", self._current_template.name)
        fname = fname.replace("{date}", now.strftime("%Y%m%d"))
        fname = fname.replace("{datetime}", now.strftime("%Y%m%d_%H%M%S"))
        # Sanitize
        fname = re.sub(r'[\\/:*?"<>|]', "_", fname)
        base_fname = fname[:-5] if fname.endswith(".docx") else fname
        out_path = Path(output_dir) / f"{base_fname}.docx"
        counter = 1
        while out_path.exists():
            out_path = Path(output_dir) / f"{base_fname}_{counter}.docx"
            counter += 1

        try:
            generate_document(self._current_template.file_path, replacements, out_path)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi tạo file", str(e))
            return

        self.status.showMessage(f"✅ Đã tạo: {out_path}")

        # Open folder + highlight file
        try:
            subprocess.Popen(f'explorer /select,"{str(out_path)}"')
        except Exception:
            try:
                os.startfile(str(out_path.parent))
            except Exception:
                pass

        QMessageBox.information(
            self, "Tạo file thành công",
            f"✅ File đã được tạo:\n{out_path}\n\nThư mục output đã được mở.",
        )

    def action_batch(self):
        if not self._current_template:
            return
        output_dir = self.out_path_edit.text().strip() or get_setting("default_output_folder")
        dlg = BatchDialog(
            docx_path=self._current_template.file_path,
            output_dir=output_dir,
            pattern=self._current_template.filename_pattern,
            fields=self._fields,
            script_mod=self._script_mod,
            template_name=self._current_template.name,
            parent=self,
        )
        dlg.exec()

    def action_restart(self):
        ret = QMessageBox.question(
            self, "Restart App",
            "Khởi động lại ứng dụng?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if ret == QMessageBox.StandardButton.Yes:
            # Lưu geometry trước khi restart
            geom_str = self.saveGeometry().toBase64().data().decode("utf-8")
            set_setting("window_geometry", geom_str)
            
            python = sys.executable
            QApplication.quit()
            subprocess.Popen([python] + sys.argv)

    def closeEvent(self, event):
        geom_str = self.saveGeometry().toBase64().data().decode("utf-8")
        set_setting("window_geometry", geom_str)
        super().closeEvent(event)

# ── Settings dialog ───────────────────────────────────────────────────────────

class _SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cài đặt")
        self.setMinimumWidth(480)
        self.setStyleSheet("""
            QDialog { background: #13172a; }
            QLabel { color: #c0cce8; font-size: 13px; }
            QLineEdit {
                background: #1e2230; color: #e8eaf0;
                border: 1.5px solid #3a3f55; border-radius: 6px;
                padding: 5px 8px; font-size: 13px;
            }
            QPushButton {
                border-radius: 6px; padding: 7px 18px;
                font-size: 13px; font-weight: bold;
            }
            QPushButton#ok { background: #2d3a8a; color: #aabbff; border: none; }
            QPushButton#ok:hover { background: #3d50c0; }
            QPushButton#cancel { background: #1e2535; color: #8898cc; border: 1.5px solid #2a3050; }
            QPushButton#cancel:hover { background: #2a3050; }
        """)
        v = QVBoxLayout(self)
        v.setSpacing(14)
        v.setContentsMargins(20, 20, 20, 20)

        title = QLabel("⚙️  Cài đặt chung")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #aabbff;")
        v.addWidget(title)

        # Output folder
        v.addWidget(QLabel("Thư mục xuất file mặc định:"))
        h1 = QHBoxLayout()
        self.out_edit = QLineEdit(get_setting("default_output_folder"))
        btn_browse = QPushButton("…")
        btn_browse.setObjectName("cancel")
        btn_browse.setFixedWidth(64)
        btn_browse.clicked.connect(self._browse)
        h1.addWidget(self.out_edit)
        h1.addWidget(btn_browse)
        v.addLayout(h1)

        v.addSpacing(8)

        # Buttons
        btn_row = QHBoxLayout()
        btn_ok = QPushButton("✔  Lưu")
        btn_ok.setObjectName("ok")
        btn_ok.clicked.connect(self._save)
        btn_cancel = QPushButton("✖  Hủy")
        btn_cancel.setObjectName("cancel")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        v.addLayout(btn_row)

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Chọn thư mục", self.out_edit.text())
        if folder:
            self.out_edit.setText(folder)

    def _save(self):
        set_setting("default_output_folder", self.out_edit.text().strip())
        self.accept()

# ── Guide dialog ──────────────────────────────────────────────────────────────

class _GuideDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Hướng dẫn tạo Template")
        self.resize(750, 600)
        self.setStyleSheet("""
            QDialog { background: #13172a; }
            QTextBrowser {
                background: #1a1f35; color: #e8eaf0;
                border: 1px solid #2a3050; border-radius: 6px;
                padding: 16px; font-size: 14px; line-height: 1.6;
            }
        """)
        v = QVBoxLayout(self)
        v.setContentsMargins(20, 20, 20, 20)
        
        from PySide6.QtWidgets import QTextBrowser
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        html = """
        <h2 style="color: #6c8aff; margin-bottom: 5px;">Cách tạo trường dữ liệu trong file Word</h2>
        <p>Để đánh dấu một vị trí cần tự động điền dữ liệu trong file Word, bạn sử dụng cú pháp: <br>
        <b style="color: #6bffaa; font-size: 16px;">{{ten_truong|kieu_du_lieu}}</b> hoặc chỉ cần <b style="color: #6bffaa; font-size: 16px;">{{ten_truong}}</b></p>
        <p>Ví dụ: <code>{{ho_va_ten|text}}</code>, <code>{{ngay_sinh|date}}</code>, hoặc đơn giản là <code>{{dia_chi}}</code> (mặc định sẽ là kiểu text).</p>
        
        <h3 style="color: #aabbff; margin-top: 20px;">Các kiểu dữ liệu (Type) hỗ trợ:</h3>
        <ul style="margin-top: 5px;">
            <li style="margin-bottom: 8px;"><b>text</b>: Văn bản 1 dòng bình thường.</li>
            <li style="margin-bottom: 8px;"><b>m_text</b>: Văn bản nhiều dòng (có thể ấn Enter xuống dòng thoải mái).</li>
            <li style="margin-bottom: 8px;"><b>text_lower</b>: Tự động chuyển thành <i>chữ thường</i>.</li>
            <li style="margin-bottom: 8px;"><b>text_upper</b>: Tự động <i>IN HOA TOÀN BỘ</i>.</li>
            <li style="margin-bottom: 8px;"><b>text_capitalize</b>: <i>Viết hoa chữ cái đầu tiên</i> của câu.</li>
            <li style="margin-bottom: 8px;"><b>text_title</b>: <i>Viết Hoa Chữ Cái Đầu</i> Của Từng Từ.</li>
            <li style="margin-bottom: 8px;"><b>number / float</b>: Số thập phân (có thể dùng để tính toán trong script).</li>
            <li style="margin-bottom: 8px;"><b>integer</b>: Số nguyên.</li>
            <li style="margin-bottom: 8px;"><b>date</b>: Ngày tháng năm (sẽ hiển thị bộ chọn lịch).</li>
            <li style="margin-bottom: 8px;"><b>datetime</b>: Ngày tháng năm và giờ phút.</li>
            <li style="margin-bottom: 8px; color: #ff8b8b;"><b>script</b>: Gọi hàm Python để tự động tính toán từ các trường khác.<br>
            <i>Lưu ý quan trọng: Tên trường script không được chứa dấu cách, tiếng Việt có dấu hay ký tự đặc biệt vì nó sẽ được dùng làm tên hàm. (Nên dùng: ho_va_ten).</i></li>
        </ul>
        
        <h3 style="color: #aabbff; margin-top: 20px;">Mẹo sử dụng:</h3>
        <p>- Nếu bạn khai báo cùng một tên trường nhiều lần trong Word (ví dụ nhiều chỗ cùng gọi <code>{{ngay_sinh|date}}</code>), bạn chỉ phải nhập liệu 1 lần trên App, hệ thống sẽ tự động điền vào mọi chỗ tương ứng.</p>
        <p>- <b>Tạo tự động hàng loạt bằng Excel (Batch):</b> File Excel của bạn cần có hàng đầu tiên (Header) là các cột có tên trùng khớp với <b>ten_truong</b> trong Word. Khi đó App sẽ tự động đọc từng dòng Excel và xuất ra các file Word tương ứng.</p>
        
        <h3 style="color: #aabbff; margin-top: 20px;">Quy tắc đặt tên file đầu ra (Output Pattern):</h3>
        <p>Bạn có thể cấu hình tên file được tạo ra trong phần quản lý template. Hệ thống hỗ trợ các từ khóa động (đặt trong ngoặc nhọn):</p>
        <ul style="margin-top: 5px;">
            <li style="margin-bottom: 8px;"><b>{template_name}</b>: Tên của template bạn đã lưu trong hệ thống.</li>
            <li style="margin-bottom: 8px;"><b>{date}</b>: Ngày hiện tại (ví dụ: <i>20260627</i>).</li>
            <li style="margin-bottom: 8px;"><b>{datetime}</b>: Ngày và giờ hiện tại (ví dụ: <i>20260627_153000</i>).</li>
            <li style="margin-bottom: 8px;"><b>{ten_truong}</b>: Bất kỳ tên trường nào có trong file Word của bạn. Ví dụ: <code>hop_dong_{ho_va_ten}_{date}</code> sẽ tạo ra file tên <i>hop_dong_Nguyen_Van_A_20260627.docx</i>.</li>
        </ul>
        <p><i>Lưu ý: Nếu file được tạo ra bị trùng tên, hệ thống sẽ tự động thêm hậu tố _1, _2... để không ghi đè file cũ.</i></p>
        """
        browser.setHtml(html)
        v.addWidget(browser)

