"""Các widget nhập liệu tương ứng với từng kiểu trường."""
from __future__ import annotations

import re
from datetime import date, datetime

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDateEdit, QDateTimeEdit,
    QHBoxLayout, QLabel, QLineEdit, QTextEdit, QSizePolicy,
    QSpinBox, QWidget, QDoubleSpinBox, QFrame,
)
from PySide6.QtGui import QDoubleValidator, QIntValidator


# ── Palette helper ────────────────────────────────────────────────────────────
STYLE_INPUT = """
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
    QComboBox::drop-down { border: none; }
    QComboBox::down-arrow { image: none; width: 0; }
"""


class TextField(QWidget):
    """text / m_text / text_lower / text_upper / text_capitalize / text_title"""

    FORMAT_MAP = {
        "text":           lambda s: s,
        "m_text":         lambda s: s,
        "text_lower":     lambda s: s.lower(),
        "text_upper":     lambda s: s.upper(),
        "text_capitalize": lambda s: s.capitalize(),
        "text_title":     lambda s: s.title(),
    }

    def __init__(self, ftype: str = "text", parent=None):
        super().__init__(parent)
        self.ftype = ftype
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        if ftype == "m_text":
            self.edit = QTextEdit()
            self.edit.setAcceptRichText(False)
            self.edit.setPlaceholderText("Nhập văn bản (nhiều dòng)…")
            self.edit.setFixedHeight(120) # Đủ rộng cho 4-5 dòng rõ rệt
        else:
            self.edit = QLineEdit()
            self.edit.setPlaceholderText("Nhập văn bản…")
            
        self.edit.setStyleSheet(STYLE_INPUT)
        layout.addWidget(self.edit)

        if ftype not in ("text", "m_text"):
            self.combo = QComboBox()
            self.combo.addItems(["Giữ nguyên", "Chữ thường", "CHỮ HOA",
                                  "Viết hoa đầu câu", "Viết Hoa Đầu Từ"])
            type_map = {
                "text_lower": 1, "text_upper": 2,
                "text_capitalize": 3, "text_title": 4,
            }
            self.combo.setCurrentIndex(type_map.get(ftype, 0))
            self.combo.setStyleSheet(STYLE_INPUT)
            self.combo.setFixedWidth(160)
            layout.addWidget(self.combo)
        else:
            self.combo = None

    def get_value(self) -> str:
        if isinstance(self.edit, QTextEdit):
            text = self.edit.toPlainText()
        else:
            text = self.edit.text()
            
        if self.combo:
            idx = self.combo.currentIndex()
            fns = [
                lambda s: s,
                lambda s: s.lower(),
                lambda s: s.upper(),
                lambda s: s.capitalize(),
                lambda s: s.title(),
            ]
            return fns[idx](text)
        return self.FORMAT_MAP.get(self.ftype, lambda s: s)(text)

    def get_typed_value(self) -> str:
        return self.get_value()

    def set_value(self, v: str):
        if isinstance(self.edit, QTextEdit):
            self.edit.setPlainText(v)
        else:
            self.edit.setText(v)


class NumberField(QWidget):
    """number / integer / float"""

    def __init__(self, ftype: str = "number", parent=None):
        super().__init__(parent)
        self.ftype = ftype
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        if ftype == "integer":
            self.spin = QSpinBox()
            self.spin.setRange(-2_147_483_648, 2_147_483_647)
            self.spin.setGroupSeparatorShown(True)
        else:
            self.spin = QDoubleSpinBox()
            self.spin.setRange(-1e15, 1e15)
            self.spin.setDecimals(4)
            self.spin.setGroupSeparatorShown(True)
        self.spin.setStyleSheet(STYLE_INPUT)
        layout.addWidget(self.spin)

    def get_value(self) -> str:
        v = self.spin.value()
        if self.ftype == "integer":
            return str(int(v))
        # Remove trailing zeros
        s = f"{v:.4f}".rstrip("0").rstrip(".")
        return s

    def get_typed_value(self) -> int | float:
        v = self.spin.value()
        if self.ftype == "integer":
            return int(v)
        return float(v)

    def set_value(self, v: str):
        try:
            self.spin.setValue(float(v))
        except Exception:
            pass


class DateField(QWidget):
    """date"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.edit = QDateEdit()
        self.edit.setCalendarPopup(True)
        self.edit.setDate(QDate.currentDate())
        self.edit.setDisplayFormat("dd/MM/yyyy")
        self.edit.setStyleSheet(STYLE_INPUT)
        layout.addWidget(self.edit)

    def get_value(self) -> str:
        return self.edit.date().toString("dd/MM/yyyy")

    def get_typed_value(self) -> date:
        d = self.edit.date()
        return date(d.year(), d.month(), d.day())

    def set_value(self, v: str):
        try:
            d = datetime.strptime(v, "%d/%m/%Y")
            self.edit.setDate(QDate(d.year, d.month, d.day))
        except Exception:
            pass


class DateTimeField(QWidget):
    """datetime"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.edit = QDateTimeEdit()
        self.edit.setCalendarPopup(True)
        self.edit.setDateTime(datetime.now())
        self.edit.setDisplayFormat("dd/MM/yyyy HH:mm")
        self.edit.setStyleSheet(STYLE_INPUT)
        layout.addWidget(self.edit)

    def get_value(self) -> str:
        return self.edit.dateTime().toString("dd/MM/yyyy HH:mm")

    def get_typed_value(self) -> datetime:
        dt = self.edit.dateTime()
        d = dt.date()
        t = dt.time()
        return datetime(d.year(), d.month(), d.day(), t.hour(), t.minute())

    def set_value(self, v: str):
        try:
            dt = datetime.strptime(v, "%d/%m/%Y %H:%M")
            from PySide6.QtCore import QDateTime
            self.edit.setDateTime(QDateTime(dt.year, dt.month, dt.day,
                                            dt.hour, dt.minute))
        except Exception:
            pass


class ScriptField(QWidget):
    """script – không cần input (hoặc hiện input nếu có args)."""

    def __init__(self, arg_names: list[str] | None = None, covered_args: list[str] | None = None, parent=None):
        super().__init__(parent)
        self.arg_names = arg_names or []
        self.covered_args = covered_args or []
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.arg_edits: dict[str, QLineEdit] = {}

        uncovered = [name for name in self.arg_names if name not in self.covered_args]

        if not uncovered:
            lbl = QLabel("(tự động từ script)")
            lbl.setStyleSheet("color: #8090b0; font-style: italic; font-size: 12px;")
            layout.addWidget(lbl)
        else:
            for name in uncovered:
                lbl = QLabel(f"{name}:")
                lbl.setStyleSheet("color: #a0aec0; font-size: 12px;")
                edit = QLineEdit()
                edit.setPlaceholderText(name)
                edit.setStyleSheet(STYLE_INPUT)
                self.arg_edits[name] = edit
                layout.addWidget(lbl)
                layout.addWidget(edit)

    def get_kwargs(self) -> dict[str, str]:
        return {k: v.text() for k, v in self.arg_edits.items()}

    def get_value(self) -> str:
        return ""  # Computed at generate time

    def get_typed_value(self) -> str:
        return self.get_value()


def make_field_widget(ftype: str, script_args: list[str] | None = None, covered_args: list[str] | None = None) -> QWidget:
    """Factory: tạo widget phù hợp với ftype."""
    ftype = ftype.strip().lower()
    if ftype in ("text", "m_text", "text_lower", "text_upper", "text_capitalize", "text_title"):
        return TextField(ftype)
    if ftype in ("number", "integer", "float"):
        return NumberField(ftype)
    if ftype == "date":
        return DateField()
    if ftype in ("datetime", "date_time"):
        return DateTimeField()
    if ftype == "script":
        return ScriptField(script_args, covered_args)
    # fallback
    return TextField("text")
