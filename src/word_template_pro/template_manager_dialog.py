"""Dialog quản lý danh sách template."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from .database import Template, get_session

STYLE = """
QDialog {
    background: #13172a;
}
QLabel {
    color: #c8d0e8;
    font-size: 13px;
}
QTableWidget {
    background: #1a1f35;
    color: #dde3f5;
    gridline-color: #2a2f48;
    border: 1.5px solid #2a2f48;
    border-radius: 8px;
    font-size: 13px;
    selection-background-color: #2d3a6e;
}
QHeaderView::section {
    background: #1e2545;
    color: #8898cc;
    padding: 6px;
    border: none;
    font-size: 12px;
    font-weight: bold;
}
QPushButton {
    border-radius: 6px;
    padding: 6px 16px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton#danger {
    background: #4a1a2a;
    color: #ff6b8a;
    border: 1.5px solid #6b2035;
}
QPushButton#danger:hover { background: #6b2035; }
QPushButton#primary {
    background: #2d3a8a;
    color: #aabbff;
    border: 1.5px solid #3d50c0;
}
QPushButton#primary:hover { background: #3d50c0; }
QPushButton#secondary {
    background: #1e2535;
    color: #8898cc;
    border: 1.5px solid #2a3050;
}
QPushButton#secondary:hover { background: #2a3a5a; }
"""


class TemplateManagerDialog(QDialog):
    template_selected = Signal(int)   # template id

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Quản lý Template")
        self.setMinimumSize(780, 440)
        self.setStyleSheet(STYLE)
        self._build_ui()
        self._load_data()

    def _build_ui(self):
        v = QVBoxLayout(self)
        v.setSpacing(12)
        v.setContentsMargins(18, 18, 18, 18)

        title = QLabel("📋  Danh sách Template")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #aabbff;")
        v.addWidget(title)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["ID", "Tên template", "Tên file", "Mẫu tên output file"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(2, 180)
        self.table.setColumnWidth(3, 220)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked)
        self.table.setAlternatingRowColors(True)
        self.table.itemChanged.connect(self._on_item_changed)
        v.addWidget(self.table)

        # Buttons
        btn_row = QHBoxLayout()
        self.btn_open = QPushButton("✏️  Mở template này")
        self.btn_open.setObjectName("primary")
        self.btn_open.clicked.connect(self._open_selected)

        self.btn_open_folder = QPushButton("📂 Mở thư mục")
        self.btn_open_folder.setObjectName("secondary")
        self.btn_open_folder.clicked.connect(self._open_selected_folder)

        self.btn_del = QPushButton("🗑  Xóa")
        self.btn_del.setObjectName("danger")
        self.btn_del.clicked.connect(self._delete_selected)

        btn_close = QPushButton("✖  Đóng")
        btn_close.setObjectName("secondary")
        btn_close.clicked.connect(self.reject)

        btn_row.addWidget(self.btn_open)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_open_folder)
        btn_row.addWidget(self.btn_del)
        btn_row.addWidget(btn_close)
        v.addLayout(btn_row)

    def _load_data(self):
        self.table.itemChanged.disconnect(self._on_item_changed)
        self.table.setRowCount(0)
        with get_session() as s:
            rows = s.query(Template).order_by(Template.id).all()
            for tpl in rows:
                r = self.table.rowCount()
                self.table.insertRow(r)
                id_item = QTableWidgetItem(str(tpl.id))
                id_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                self.table.setItem(r, 0, id_item)
                self.table.setItem(r, 1, QTableWidgetItem(tpl.name))
                file_name_item = QTableWidgetItem(Path(tpl.file_path).name)
                file_name_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                self.table.setItem(r, 2, file_name_item)
                self.table.setItem(r, 3, QTableWidgetItem(tpl.filename_pattern))
        self.table.itemChanged.connect(self._on_item_changed)

    def _open_selected_folder(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Chưa chọn", "Vui lòng chọn một template để mở thư mục.")
            return
        id_item = self.table.item(row, 0)
        if not id_item:
            return
        tpl_id = int(id_item.text())
        with get_session() as s:
            tpl = s.get(Template, tpl_id)
            if tpl:
                import subprocess
                from pathlib import Path
                folder = Path(tpl.file_path).parent
                if folder.exists():
                    try:
                        subprocess.Popen(f'explorer /select,"{tpl.file_path}"')
                    except Exception as e:
                        QMessageBox.warning(self, "Lỗi", f"Không thể mở thư mục:\n{e}")
                else:
                    QMessageBox.warning(self, "Lỗi", "Thư mục không tồn tại!")

    def _on_item_changed(self, item: QTableWidgetItem):
        row = item.row()
        id_item = self.table.item(row, 0)
        if id_item is None:
            return
        tpl_id = int(id_item.text())
        col = item.column()
        with get_session() as s:
            tpl = s.get(Template, tpl_id)
            if tpl is None:
                return
            if col == 1:
                tpl.name = item.text()
            elif col == 3:
                tpl.filename_pattern = item.text()
            s.commit()

    def _open_selected(self):
        rows = self.table.selectedItems()
        if not rows:
            return
        row = self.table.currentRow()
        id_item = self.table.item(row, 0)
        if id_item:
            self.template_selected.emit(int(id_item.text()))
            self.accept()

    def _delete_selected(self):
        row = self.table.currentRow()
        if row < 0:
            return
        id_item = self.table.item(row, 0)
        if not id_item:
            return
        tpl_id = int(id_item.text())
        name_item = self.table.item(row, 1)
        name = name_item.text() if name_item else ""
        ret = QMessageBox.question(
            self, "Xác nhận xóa",
            f"Xóa template '{name}' khỏi danh sách?\n(File gốc không bị xóa)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if ret == QMessageBox.StandardButton.Yes:
            with get_session() as s:
                tpl = s.get(Template, tpl_id)
                if tpl:
                    s.delete(tpl)
                    s.commit()
            self._load_data()
