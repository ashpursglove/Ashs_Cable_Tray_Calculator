"""
gui.py

PyQt5 GUI for the Cable Tray Calculator.

Key UI elements:
- Cable library & custom input (diameter, weight, quantity).
- Tray library & custom input (width, height, capacity, self-weight).
- Results panel showing structural loading and area fill statistics.

All heavy calculations are delegated to models.py.
"""

from __future__ import annotations

import json
import os
import math
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import csv

from PyQt5 import QtCore, QtGui, QtWidgets

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas




from models import (
    CableType,
    TrayType,
    compute_cable_tray_stats,
    get_default_cables,
    get_default_trays,
)


class CableTrayCalculator(QtWidgets.QMainWindow):
    """
    Main window for the Cable Tray Calculator.

    Responsibilities:
    - Set up all widgets, layouts, and dark theme.
    - Let the user build a list of cables in the tray.
    - Let the user choose or define a tray.
    - Show computed structural and fill metrics in real time.
    """


    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Ash's Cable Tray Calculator")
        self.resize(1000, 650)

        # Currently loaded config file path (None = unsaved / new)
        self.current_config_path: Optional[str] = None

        # Core data: default libraries
        self.default_cables: List[CableType] = get_default_cables()


        self.default_trays: List[TrayType] = get_default_trays()

        # Build UI
        self._create_menu_bar()
        self._create_widgets()
        self._create_layouts()
        self._connect_signals()
        self._populate_libraries()

        # Apply dark style
        self._apply_dark_blue_style()

        # Initial compute
        self.recalculate()

    # ------------------------------------------------------------------
    # UI creation
    # ------------------------------------------------------------------

    def _create_widgets(self) -> None:
        """Create all GUI widgets."""
        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        self.central_layout = QtWidgets.QVBoxLayout(central)
        self.central_layout.setContentsMargins(10, 10, 10, 10)
        self.central_layout.setSpacing(10)

        # Top-level split: left = cables, right = tray + results
        self.main_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self.central_layout.addWidget(self.main_splitter)

        # Left side: cable configuration
        self.cable_widget = QtWidgets.QWidget()
        self.cable_layout = QtWidgets.QVBoxLayout(self.cable_widget)
        self.cable_layout.setContentsMargins(0, 0, 0, 0)
        self.cable_layout.setSpacing(8)

        self.cable_header_label = QtWidgets.QLabel("Cables in Tray")
        font = self.cable_header_label.font()
        font.setPointSize(11)
        font.setBold(True)
        self.cable_header_label.setFont(font)

        self.cable_layout.addWidget(self.cable_header_label)


        # Cable input form
        self.cable_form = QtWidgets.QFormLayout()
        self.cable_form.setLabelAlignment(QtCore.Qt.AlignRight)
        self.cable_form.setFormAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)

        # Cable name (for custom or library display)
        self.cable_name_edit = QtWidgets.QLineEdit()
        self.cable_name_edit.setPlaceholderText("Custom cable name")

        self.cable_combo = QtWidgets.QComboBox()
        self.cable_combo.setEditable(False)

        self.custom_cable_diameter_spin = QtWidgets.QDoubleSpinBox()
        self.custom_cable_diameter_spin.setRange(1.0, 200.0)
        self.custom_cable_diameter_spin.setDecimals(1)
        self.custom_cable_diameter_spin.setSuffix(" mm")

        self.custom_cable_weight_spin = QtWidgets.QDoubleSpinBox()
        self.custom_cable_weight_spin.setRange(0.001, 100.0)
        self.custom_cable_weight_spin.setDecimals(3)
        self.custom_cable_weight_spin.setSuffix(" kg/m")

        self.cable_quantity_spin = QtWidgets.QSpinBox()
        self.cable_quantity_spin.setRange(1, 9999)
        self.cable_quantity_spin.setValue(1)

        self.cable_form.addRow("Name:", self.cable_name_edit)
        self.cable_form.addRow("Cable type:", self.cable_combo)
        self.cable_form.addRow("Diameter:", self.custom_cable_diameter_spin)
        self.cable_form.addRow("Weight:", self.custom_cable_weight_spin)
        self.cable_form.addRow("Quantity:", self.cable_quantity_spin)


        self.cable_layout.addLayout(self.cable_form)

        # Buttons for managing cable list
        self.cable_button_row = QtWidgets.QHBoxLayout()
        self.add_cable_btn = QtWidgets.QPushButton("Add cable to tray")
        self.remove_cable_btn = QtWidgets.QPushButton("Remove selected")
        self.clear_cable_btn = QtWidgets.QPushButton("Clear all")
        self.cable_button_row.addWidget(self.add_cable_btn)
        self.cable_button_row.addWidget(self.remove_cable_btn)
        self.cable_button_row.addWidget(self.clear_cable_btn)

        self.cable_layout.addLayout(self.cable_button_row)

        # Table: list of cables in tray
        self.cable_table = QtWidgets.QTableWidget()
        self.cable_table.setColumnCount(4)
        self.cable_table.setHorizontalHeaderLabels(
            ["Cable", "Diameter (mm)", "Weight (kg/m)", "Qty"]
        )
        self.cable_table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Stretch
        )
        self.cable_table.verticalHeader().setVisible(False)
        self.cable_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.cable_table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.cable_table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.SelectedClicked
        )
        self.cable_table.setAlternatingRowColors(True)

        self.cable_layout.addWidget(self.cable_table)

        # Right side: tray + results
        self.right_widget = QtWidgets.QWidget()
        self.right_layout = QtWidgets.QVBoxLayout(self.right_widget)
        self.right_layout.setContentsMargins(0, 0, 0, 0)
        self.right_layout.setSpacing(8)





        self.tray_group = QtWidgets.QGroupBox("Tray Configuration")
        self.tray_form = QtWidgets.QFormLayout(self.tray_group)
        self.tray_form.setLabelAlignment(QtCore.Qt.AlignRight)

        # Tray name (for custom name or showing preset name)
        self.tray_name_edit = QtWidgets.QLineEdit()
        self.tray_name_edit.setPlaceholderText("Custom tray name")

        self.tray_combo = QtWidgets.QComboBox()
        self.tray_combo.setEditable(False)

        self.tray_width_spin = QtWidgets.QDoubleSpinBox()
        self.tray_width_spin.setRange(50.0, 1000.0)
        self.tray_width_spin.setDecimals(1)
        self.tray_width_spin.setSuffix(" mm")

        self.tray_height_spin = QtWidgets.QDoubleSpinBox()
        self.tray_height_spin.setRange(25.0, 300.0)
        self.tray_height_spin.setDecimals(1)
        self.tray_height_spin.setSuffix(" mm")

        self.tray_max_load_spin = QtWidgets.QDoubleSpinBox()
        self.tray_max_load_spin.setRange(1.0, 500.0)
        self.tray_max_load_spin.setDecimals(1)
        self.tray_max_load_spin.setSuffix(" kg/m")

        self.tray_self_weight_spin = QtWidgets.QDoubleSpinBox()
        self.tray_self_weight_spin.setRange(0.1, 100.0)
        self.tray_self_weight_spin.setDecimals(2)
        self.tray_self_weight_spin.setSuffix(" kg/m")

        self.tray_fill_ratio_spin = QtWidgets.QDoubleSpinBox()
        self.tray_fill_ratio_spin.setRange(0.1, 1.0)
        self.tray_fill_ratio_spin.setDecimals(2)
        self.tray_fill_ratio_spin.setSingleStep(0.05)
        self.tray_fill_ratio_spin.setValue(0.6)

        self.tray_form.addRow("Tray name:", self.tray_name_edit)
        self.tray_form.addRow("Tray type:", self.tray_combo)
        self.tray_form.addRow("Width:", self.tray_width_spin)
        self.tray_form.addRow("Side height:", self.tray_height_spin)
        self.tray_form.addRow("Max load:", self.tray_max_load_spin)
        self.tray_form.addRow("Tray self-weight:", self.tray_self_weight_spin)
        self.tray_form.addRow("Max fill ratio:", self.tray_fill_ratio_spin)

        self.right_layout.addWidget(self.tray_group)

        # Results group
        self.results_group = QtWidgets.QGroupBox("Results")
        self.results_form = QtWidgets.QFormLayout(self.results_group)
        self.results_form.setLabelAlignment(QtCore.Qt.AlignRight)

        def _make_result_label() -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel("-")
            lbl.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            mono_font = QtGui.QFont("Consolas")
            mono_font.setPointSize(9)
            lbl.setFont(mono_font)
            return lbl

        self.lbl_total_cable_weight = _make_result_label()
        self.lbl_tray_self_weight = _make_result_label()
        self.lbl_total_weight = _make_result_label()
        self.lbl_allowable_load = _make_result_label()
        self.lbl_structural_util = _make_result_label()
        self.lbl_cable_area = _make_result_label()
        self.lbl_tray_area = _make_result_label()
        self.lbl_area_fill = _make_result_label()
        self.lbl_area_fill_limit = _make_result_label()
        self.lbl_status = _make_result_label()

        self.results_form.addRow("Cable weight:", self.lbl_total_cable_weight)
        self.results_form.addRow("Tray self-weight:", self.lbl_tray_self_weight)
        self.results_form.addRow("Total weight:", self.lbl_total_weight)
        self.results_form.addRow("Tray allowable load:", self.lbl_allowable_load)
        self.results_form.addRow("Structural utilisation:", self.lbl_structural_util)
        self.results_form.addRow("Total cable area:", self.lbl_cable_area)
        self.results_form.addRow("Tray usable area:", self.lbl_tray_area)
        self.results_form.addRow("Area fill:", self.lbl_area_fill)
        self.results_form.addRow("Recommended max fill:", self.lbl_area_fill_limit)
        self.results_form.addRow("Status:", self.lbl_status)

        self.right_layout.addWidget(self.results_group)


        # Bottom-right controls: Export PDF + Recalculate






        # Bottom-right controls: Export PDF / CSV + Recalculate
        self.button_row = QtWidgets.QHBoxLayout()
        self.export_pdf_button = QtWidgets.QPushButton("Export PDF report")
        self.export_csv_button = QtWidgets.QPushButton("Export CSV")
        self.recalc_button = QtWidgets.QPushButton("Recalculate")

        self.button_row.addWidget(self.export_pdf_button)
        self.button_row.addWidget(self.export_csv_button)
        self.button_row.addStretch(1)
        self.button_row.addWidget(self.recalc_button)

        self.right_layout.addLayout(self.button_row)

        # self.button_row = QtWidgets.QHBoxLayout()
        # self.export_pdf_button = QtWidgets.QPushButton("Export PDF report")
        # self.recalc_button = QtWidgets.QPushButton("Recalculate")

        # self.button_row.addWidget(self.export_pdf_button)
        # self.button_row.addStretch(1)
        # self.button_row.addWidget(self.recalc_button)

        # self.right_layout.addLayout(self.button_row)


        # Put left/right into splitter
        self.main_splitter.addWidget(self.cable_widget)
        self.main_splitter.addWidget(self.right_widget)
        self.main_splitter.setStretchFactor(0, 3)
        self.main_splitter.setStretchFactor(1, 2)

    def _create_layouts(self) -> None:
        """Layouts are created inside _create_widgets; nothing extra here."""
        pass
    
    # ------------------------------------------------------------------
    # Menu bar
    # ------------------------------------------------------------------

    def _create_menu_bar(self) -> None:
        """
        Create a standard File menu with New / Open / Save / Save As / Exit,
        plus a sassy About dialog.
        """
        menubar = self.menuBar()

        # ----------------- File menu -----------------
        file_menu = menubar.addMenu("&File")

        self.action_new = QtWidgets.QAction("&New", self)
        self.action_new.setShortcut("Ctrl+N")
        self.action_new.setStatusTip("Clear current cables and tray configuration")

        self.action_open = QtWidgets.QAction("&Open…", self)
        self.action_open.setShortcut("Ctrl+O")
        self.action_open.setStatusTip("Open a saved cable tray configuration")

        self.action_save = QtWidgets.QAction("&Save", self)
        self.action_save.setShortcut("Ctrl+S")
        self.action_save.setStatusTip("Save current configuration")

        self.action_save_as = QtWidgets.QAction("Save &As…", self)
        self.action_save_as.setStatusTip("Save current configuration to a new file")

        self.action_exit = QtWidgets.QAction("E&xit", self)
        self.action_exit.setShortcut("Ctrl+Q")
        self.action_exit.setStatusTip("Quit the application")

        file_menu.addAction(self.action_new)
        file_menu.addSeparator()
        file_menu.addAction(self.action_open)
        file_menu.addSeparator()
        file_menu.addAction(self.action_save)
        file_menu.addAction(self.action_save_as)
        file_menu.addSeparator()
        file_menu.addAction(self.action_exit)

        # ----------------- About menu -----------------
        about_menu = menubar.addMenu("&About")

        self.action_about = QtWidgets.QAction("&About Cable Tray Calculator", self)
        self.action_about.setStatusTip("Who built this thing and why?")

        about_menu.addAction(self.action_about)

        # Wire up actions
        self.action_new.triggered.connect(self._file_new)
        self.action_open.triggered.connect(self._file_open)
        self.action_save.triggered.connect(self._file_save)
        self.action_save_as.triggered.connect(self._file_save_as)
        self.action_exit.triggered.connect(self.close)
        self.action_about.triggered.connect(self._show_about_dialog)



    def _connect_signals(self) -> None:
        """Connect all signals to slots."""
        self.cable_combo.currentIndexChanged.connect(self._on_cable_combo_changed)
        self.add_cable_btn.clicked.connect(self._on_add_cable_clicked)
        self.remove_cable_btn.clicked.connect(self._on_remove_cable_clicked)
        self.clear_cable_btn.clicked.connect(self._on_clear_cables_clicked)

        self.tray_combo.currentIndexChanged.connect(self._on_tray_combo_changed)

        # Recalculate when key values change
        self.tray_width_spin.valueChanged.connect(self.recalculate)
        self.tray_height_spin.valueChanged.connect(self.recalculate)
        self.tray_max_load_spin.valueChanged.connect(self.recalculate)
        self.tray_self_weight_spin.valueChanged.connect(self.recalculate)


        # self.tray_fill_ratio_spin.valueChanged.connect(self.recalculate)
        # self.cable_table.itemChanged.connect(self._on_cable_table_item_changed)

        # # self.recalc_button.clicked.connect(self.recalculate)
        # self.recalc_button.clicked.connect(self.recalculate)
        # self.export_pdf_button.clicked.connect(self._on_export_pdf_clicked)

        self.tray_fill_ratio_spin.valueChanged.connect(self.recalculate)
        self.cable_table.itemChanged.connect(self._on_cable_table_item_changed)

        self.recalc_button.clicked.connect(self.recalculate)
        self.export_pdf_button.clicked.connect(self._on_export_pdf_clicked)
        self.export_csv_button.clicked.connect(self._on_export_csv_clicked)







    def _populate_libraries(self) -> None:
        """Populate cable and tray combo boxes from the default libraries."""
        self.cable_combo.blockSignals(True)
        self.cable_combo.clear()
        # First item is a "Custom cable" entry
        self.cable_combo.addItem("Custom cable...")
        for cable in self.default_cables:
            self.cable_combo.addItem(cable.name, cable)
        self.cable_combo.setCurrentIndex(0)
        self.cable_combo.blockSignals(False)
        self._on_cable_combo_changed(0)

        self.tray_combo.blockSignals(True)
        self.tray_combo.clear()
        # First item is a "Custom tray" entry
        self.tray_combo.addItem("Custom tray...")
        for tray in self.default_trays:
            self.tray_combo.addItem(tray.name, tray)
        self.tray_combo.setCurrentIndex(1 if self.tray_combo.count() > 1 else 0)
        self.tray_combo.blockSignals(False)
        self._on_tray_combo_changed(self.tray_combo.currentIndex())

    # ------------------------------------------------------------------
    # Styling
    # ------------------------------------------------------------------

    def _apply_dark_blue_style(self) -> None:
        """
        Apply a dark blue QSS theme, reusable across your other tools.
        """
        dark_bg = "#0b1f3b"    # deep navy
        dark_bg_alt = "#12284a"
        panel_bg = "#16243a"
        text_color = "#f0f4ff"
        accent = "#ff8c00"
        border_color = "#2b3f5f"
        highlight = "#2854a0"

        qss = f"""
        QWidget {{
            background-color: {dark_bg};
            color: {text_color};
            selection-background-color: {highlight};
            selection-color: {text_color};
        }}
        QMainWindow::separator {{
            background-color: {border_color};
        }}
        QGroupBox {{
            background-color: {panel_bg};
            border: 1px solid {border_color};
            border-radius: 4px;
            margin-top: 6px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 2px 6px;
            color: {accent};
        }}
        QTableWidget {{
            gridline-color: {border_color};
            background-color: {dark_bg_alt};
            alternate-background-color: {dark_bg};
        }}
        QHeaderView::section {{
            background-color: {panel_bg};
            color: {text_color};
            padding: 4px;
            border: 1px solid {border_color};
        }}
        QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
            background-color: {dark_bg_alt};
            border: 1px solid {border_color};
            border-radius: 3px;
            padding: 2px 4px;
            selection-background-color: {highlight};
            selection-color: {text_color};
        }}
        QComboBox QAbstractItemView {{
            background-color: {dark_bg_alt};
            selection-background-color: {highlight};
        }}
        QPushButton {{
            background-color: {accent};
            color: #000000;
            border-radius: 4px;
            padding: 4px 10px;
            font-weight: bold;
        }}
        QPushButton:hover {{
            background-color: #ff9e26;
        }}
        QPushButton:pressed {{
            background-color: #e67a00;
        }}
        QSplitter::handle {{
            background-color: {border_color};
        }}
        QScrollBar:vertical, QScrollBar:horizontal {{
            background: {dark_bg_alt};
        }}
        """
        self.setStyleSheet(qss)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------


    def _on_cable_combo_changed(self, index: int) -> None:
        """
        Update name/diameter/weight fields when user selects a cable type.

        Index 0 is "Custom cable..." so fields are editable and name is free text.
        Other indices pull from the default cable library and lock fields + name.
        """
        if index <= 0:
            # Custom cable
            self.custom_cable_diameter_spin.setReadOnly(False)
            self.custom_cable_weight_spin.setReadOnly(False)
            self.cable_name_edit.setReadOnly(False)
            # Don't forcibly clear name, user might be tweaking
            if not self.cable_name_edit.text().strip():
                self.cable_name_edit.setText("Custom cable")
            return

        cable_data = self.cable_combo.itemData(index)
        if isinstance(cable_data, CableType):
            self.custom_cable_diameter_spin.setValue(cable_data.diameter_mm)
            self.custom_cable_weight_spin.setValue(cable_data.weight_kg_per_m)
            self.custom_cable_diameter_spin.setReadOnly(True)
            self.custom_cable_weight_spin.setReadOnly(True)

            # Show the library name and lock it
            self.cable_name_edit.setText(cable_data.name)
            self.cable_name_edit.setReadOnly(True)




    def _on_tray_combo_changed(self, index: int) -> None:
        """
        Update tray fields when user selects a tray type.

        Index 0 is "Custom tray..." so fields are editable and name is free text.
        Other indices pull from the default tray library and lock fields + name.
        """
        if index <= 0:
            # Custom tray
            self.tray_width_spin.setReadOnly(False)
            self.tray_height_spin.setReadOnly(False)
            self.tray_max_load_spin.setReadOnly(False)
            self.tray_self_weight_spin.setReadOnly(False)
            self.tray_name_edit.setReadOnly(False)
            if not self.tray_name_edit.text().strip():
                self.tray_name_edit.setText("Custom tray")
            return

        tray_data = self.tray_combo.itemData(index)
        if isinstance(tray_data, TrayType):
            self.tray_width_spin.setValue(tray_data.width_mm)
            self.tray_height_spin.setValue(tray_data.height_mm)
            self.tray_max_load_spin.setValue(tray_data.max_load_kg_per_m)
            self.tray_self_weight_spin.setValue(tray_data.self_weight_kg_per_m)
            self.tray_fill_ratio_spin.setValue(tray_data.max_fill_ratio)

            self.tray_width_spin.setReadOnly(True)
            self.tray_height_spin.setReadOnly(True)
            self.tray_max_load_spin.setReadOnly(True)
            self.tray_self_weight_spin.setReadOnly(True)

            # Show the preset name and lock it
            self.tray_name_edit.setText(tray_data.name)
            self.tray_name_edit.setReadOnly(True)

        # Recompute when tray changes
        self.recalculate()






    def _on_add_cable_clicked(self) -> None:
        """Add a new cable entry to the table based on current input fields."""

        # Determine cable name
        index = self.cable_combo.currentIndex()
        if index <= 0:
            cable_name = self.cable_name_edit.text().strip() or "Custom cable"
        else:
            # For library entries we keep the preset name (also reflected in the name edit)
            cable_name = self.cable_name_edit.text().strip() or self.cable_combo.currentText()


        diameter = self.custom_cable_diameter_spin.value()
        weight = self.custom_cable_weight_spin.value()
        qty = self.cable_quantity_spin.value()

        if diameter <= 0 or weight <= 0 or qty <= 0:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid cable data",
                "Please ensure diameter, weight, and quantity are all > 0.",
            )
            return

        row = self.cable_table.rowCount()
        self.cable_table.insertRow(row)

        def _mk_item(text: str) -> QtWidgets.QTableWidgetItem:
            item = QtWidgets.QTableWidgetItem(text)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            return item

        self.cable_table.setItem(row, 0, _mk_item(cable_name))
        self.cable_table.setItem(row, 1, _mk_item(f"{diameter:.1f}"))
        self.cable_table.setItem(row, 2, _mk_item(f"{weight:.3f}"))
        self.cable_table.setItem(row, 3, _mk_item(str(qty)))

        self.recalculate()

    def _on_remove_cable_clicked(self) -> None:
        """Remove the currently selected cable row."""
        row = self.cable_table.currentRow()
        if row < 0:
            return
        self.cable_table.removeRow(row)
        self.recalculate()

    def _on_clear_cables_clicked(self) -> None:
        """Clear all cable rows."""
        self.cable_table.setRowCount(0)
        self.recalculate()

    def _on_cable_table_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        """
        When the user edits the cable table (e.g. quantity), recalculate.
        """
        # Basic validation: ensure numeric columns stay numeric
        col = item.column()
        if col in (1, 2, 3):
            text = item.text().strip()
            if col == 3:  # quantity (int)
                try:
                    value = int(text)
                    if value <= 0:
                        raise ValueError
                except ValueError:
                    item.setText("1")
            else:  # diameter / weight (float)
                try:
                    value = float(text)
                    if value <= 0:
                        raise ValueError
                except ValueError:
                    # Reset to 0.0, but the recalc will ignore it
                    item.setText("0.0")
        self.recalculate()

    # ------------------------------------------------------------------
    # Calculation logic
    # ------------------------------------------------------------------

    def _collect_cables_from_table(self) -> List[Tuple[CableType, int]]:
        """
        Read the cable rows and build a list of CableType + quantity.

        For simplicity:
        - Each row becomes a CableType instance (even if from library),
          so that editing the diameter/weight is fully supported.
        """
        result: List[Tuple[CableType, int]] = []

        for row in range(self.cable_table.rowCount()):
            name_item = self.cable_table.item(row, 0)
            dia_item = self.cable_table.item(row, 1)
            wt_item = self.cable_table.item(row, 2)
            qty_item = self.cable_table.item(row, 3)
            if not (name_item and dia_item and wt_item and qty_item):
                continue

            try:
                name = name_item.text().strip() or "Cable"
                diameter = float(dia_item.text())
                weight = float(wt_item.text())
                qty = int(qty_item.text())
            except ValueError:
                continue

            if diameter <= 0 or weight <= 0 or qty <= 0:
                continue

            cable = CableType(name=name, diameter_mm=diameter, weight_kg_per_m=weight)
            result.append((cable, qty))

        return result

    def _build_tray_from_fields(self) -> TrayType:
        """
        Construct a TrayType instance from the current spinbox values.

        This supports both library and custom tray configurations.
        """


        # Prefer explicit name field; fall back to combo text or a generic label
        tray_name = self.tray_name_edit.text().strip()
        if not tray_name:
            tray_name = self.tray_combo.currentText() or "Tray"

        width = self.tray_width_spin.value()
        height = self.tray_height_spin.value()
        max_load = self.tray_max_load_spin.value()
        self_weight = self.tray_self_weight_spin.value()
        max_fill_ratio = self.tray_fill_ratio_spin.value()

        return TrayType(
            name=tray_name,
            width_mm=width,
            height_mm=height,
            max_load_kg_per_m=max_load,
            self_weight_kg_per_m=self_weight,
            max_fill_ratio=max_fill_ratio,
        )



    def recalculate(self) -> None:
        """
        Collect current data, run calculations, and update result labels.
        """
        cables_with_qty = self._collect_cables_from_table()
        tray = self._build_tray_from_fields()

        stats = compute_cable_tray_stats(cables_with_qty, tray)

        # Extract
        total_cable_weight = stats["total_cable_weight_kg_per_m"]
        tray_self_weight = stats["tray_self_weight_kg_per_m"]
        total_weight = stats["total_weight_kg_per_m"]
        allowable_load = stats["allowable_load_kg_per_m"]
        struct_util_pct = stats["structural_utilisation_percent"]
        cable_area = stats["total_cable_area_mm2"]
        tray_area = stats["tray_usable_area_mm2"]
        area_fill_pct = stats["area_fill_percent"]
        area_fill_limit = stats["recommended_max_area_fill_percent"]

        self.lbl_total_cable_weight.setText(f"{total_cable_weight:.3f} kg/m")
        self.lbl_tray_self_weight.setText(f"{tray_self_weight:.3f} kg/m")
        self.lbl_total_weight.setText(f"{total_weight:.3f} kg/m")
        self.lbl_allowable_load.setText(f"{allowable_load:.1f} kg/m")
        self.lbl_structural_util.setText(f"{struct_util_pct:.1f} %")
        self.lbl_cable_area.setText(f"{cable_area:,.0f} mm²")
        self.lbl_tray_area.setText(f"{tray_area:,.0f} mm²")
        self.lbl_area_fill.setText(f"{area_fill_pct:.1f} %")
        self.lbl_area_fill_limit.setText(f"{area_fill_limit:.1f} %")


        # -----------------------------------------
# Status logic + detailed colour warnings
# -----------------------------------------

        # Determine overload states
        overloaded_struct = allowable_load > 0 and total_cable_weight > allowable_load
        overloaded_area = area_fill_pct > area_fill_limit


        # Helper to colour labels (override QSS with per-widget style)
        def set_label_color(lbl: QtWidgets.QLabel, color: str) -> None:
            lbl.setStyleSheet(f"color: {color};")


        # Default: all labels normal text colour
        default_color = "#f0f4ff"
        warning_color = "#ff4d4d"   # red
        ok_color = "#7cd67c"        # green

        # STRUCTURAL % colour
        if overloaded_struct:
            set_label_color(self.lbl_structural_util, warning_color)
        else:
            set_label_color(self.lbl_structural_util, ok_color if struct_util_pct < 80 else default_color)

        # AREA FILL % colour
        if overloaded_area:
            set_label_color(self.lbl_area_fill, warning_color)
        else:
            set_label_color(self.lbl_area_fill, ok_color if area_fill_pct < (0.8 * area_fill_limit) else default_color)

        # TOTAL WEIGHT colour (warn if exceeding tray allowable)
        if overloaded_struct:
            set_label_color(self.lbl_total_weight, warning_color)
        else:
            set_label_color(self.lbl_total_weight, default_color)

        # Status banner logic
        if not cables_with_qty:
            status_text = "No cables defined."
            status_color = "#cccccc"
        elif overloaded_struct and overloaded_area:
            status_text = "OVERLOADED: structural + fill limits exceeded"
            status_color = warning_color
        elif overloaded_struct:
            status_text = "OVERLOADED: structural limit exceeded"
            status_color = warning_color
        elif overloaded_area:
            status_text = "WARNING: area fill above recommended limit"
            status_color = warning_color
        else:
            status_text = "OK: within structural and fill limits"
            status_color = ok_color

        # Apply status styling
        self.lbl_status.setText(status_text)
        set_label_color(self.lbl_status, status_color)
    # ------------------------------------------------------------------
    # Save / load configuration to JSON
    # ------------------------------------------------------------------

    def _export_config(self) -> dict:
        """
        Build a JSON-serialisable dict representing the current setup.

        Structure:
            {
                "version": 1,
                "cables": [
                    {
                        "name": str,
                        "diameter_mm": float,
                        "weight_kg_per_m": float,
                        "qty": int
                    },
                    ...
                ],
                "tray": {
                    "name": str,
                    "width_mm": float,
                    "height_mm": float,
                    "max_load_kg_per_m": float,
                    "self_weight_kg_per_m": float,
                    "max_fill_ratio": float
                }
            }
        """
        cables: List[dict] = []

        for row in range(self.cable_table.rowCount()):
            name_item = self.cable_table.item(row, 0)
            dia_item = self.cable_table.item(row, 1)
            wt_item = self.cable_table.item(row, 2)
            qty_item = self.cable_table.item(row, 3)
            if not (name_item and dia_item and wt_item and qty_item):
                continue

            try:
                name = name_item.text().strip() or "Cable"
                diameter = float(dia_item.text())
                weight = float(wt_item.text())
                qty = int(qty_item.text())
            except ValueError:
                continue

            if diameter <= 0 or weight <= 0 or qty <= 0:
                continue

            cables.append(
                {
                    "name": name,
                    "diameter_mm": diameter,
                    "weight_kg_per_m": weight,
                    "qty": qty,
                }
            )

        tray = self._build_tray_from_fields()
        tray_dict = {
            "name": tray.name,
            "width_mm": float(tray.width_mm),
            "height_mm": float(tray.height_mm),
            "max_load_kg_per_m": float(tray.max_load_kg_per_m),
            "self_weight_kg_per_m": float(tray.self_weight_kg_per_m),
            "max_fill_ratio": float(tray.max_fill_ratio),
        }

        return {
            "version": 1,
            "cables": cables,
            "tray": tray_dict,
        }

    def _import_config(self, cfg: dict) -> None:
        """
        Apply a previously saved configuration dict to the UI.
        """
        # Restore cables
        self.cable_table.blockSignals(True)
        self.cable_table.setRowCount(0)

        for cable_cfg in cfg.get("cables", []):
            try:
                name = str(cable_cfg.get("name", "Cable"))
                diameter = float(cable_cfg.get("diameter_mm", 0.0))
                weight = float(cable_cfg.get("weight_kg_per_m", 0.0))
                qty = int(cable_cfg.get("qty", 0))
            except (TypeError, ValueError):
                continue

            if diameter <= 0 or weight <= 0 or qty <= 0:
                continue

            row = self.cable_table.rowCount()
            self.cable_table.insertRow(row)

            def _mk_item(text: str) -> QtWidgets.QTableWidgetItem:
                item = QtWidgets.QTableWidgetItem(text)
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                return item

            self.cable_table.setItem(row, 0, _mk_item(name))
            self.cable_table.setItem(row, 1, _mk_item(f"{diameter:.1f}"))
            self.cable_table.setItem(row, 2, _mk_item(f"{weight:.3f}"))
            self.cable_table.setItem(row, 3, _mk_item(str(qty)))

        self.cable_table.blockSignals(False)

        # Restore tray as "Custom tray" so fields are editable but populated
        tray_cfg = cfg.get("tray", {})

        self.tray_combo.blockSignals(True)
        if self.tray_combo.count() > 0:
            self.tray_combo.setCurrentIndex(0)  # Custom tray...
        self.tray_combo.blockSignals(False)

        try:
            width = float(tray_cfg.get("width_mm", self.tray_width_spin.value()))
            height = float(tray_cfg.get("height_mm", self.tray_height_spin.value()))
            max_load = float(
                tray_cfg.get("max_load_kg_per_m", self.tray_max_load_spin.value())
            )
            self_weight = float(
                tray_cfg.get(
                    "self_weight_kg_per_m", self.tray_self_weight_spin.value()
                )
            )
            max_fill = float(
                tray_cfg.get("max_fill_ratio", self.tray_fill_ratio_spin.value())
            )
        except (TypeError, ValueError):
            width = self.tray_width_spin.value()
            height = self.tray_height_spin.value()
            max_load = self.tray_max_load_spin.value()
            self_weight = self.tray_self_weight_spin.value()
            max_fill = self.tray_fill_ratio_spin.value()

        self.tray_width_spin.setValue(width)
        self.tray_height_spin.setValue(height)
        self.tray_max_load_spin.setValue(max_load)
        self.tray_self_weight_spin.setValue(self_weight)
        self.tray_fill_ratio_spin.setValue(max_fill)


        # Restore tray name into the editable field
        self.tray_name_edit.setReadOnly(False)
        self.tray_name_edit.setText(tray_cfg.get("name", "Custom tray"))

        # Optionally, show the loaded tray name in the window title
        tray_name = tray_cfg.get("name", "Custom tray")
        self.setWindowTitle(f"Cable Tray Calculator – {tray_name}")

        # Recalculate results with the imported data
        self.recalculate()

    # ------------------------------------------------------------------
    # File menu handlers
    # ------------------------------------------------------------------

    def _file_new(self) -> None:
        """
        Clear current cables and reset tray to current 'custom' values.
        """
        self.cable_table.setRowCount(0)
        self.current_config_path = None
        self.setWindowTitle("Cable Tray Calculator")
        self.recalculate()

    def _file_open(self) -> None:
        """
        Open a JSON configuration file and load it into the UI.
        """
        start_dir = self.current_config_path or os.getcwd()
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open cable tray configuration",
            start_dir,
            "Cable tray config (*.json);;All files (*.*)",
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception as exc:  # broad but surfaced to user
            QtWidgets.QMessageBox.critical(
                self,
                "Failed to open",
                f"Could not read configuration file:\n{exc}",
            )
            return

        self.current_config_path = path
        self._import_config(cfg)

    def _file_save(self) -> None:
        """
        Save current configuration. If no path yet, fall back to Save As.
        """
        if not self.current_config_path:
            self._file_save_as()
            return

        cfg = self._export_config()

        try:
            with open(self.current_config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(
                self,
                "Failed to save",
                f"Could not write configuration file:\n{exc}",
            )
            return

    def _file_save_as(self) -> None:
        """
        Prompt user for a file name and save configuration to that file.
        """
        start_dir = self.current_config_path or os.getcwd()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save cable tray configuration",
            start_dir,
            "Cable tray config (*.json);;All files (*.*)",
        )
        if not path:
            return

        # Ensure .json extension
        if not os.path.splitext(path)[1]:
            path = path + ".json"

        cfg = self._export_config()

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(
                self,
                "Failed to save",
                f"Could not write configuration file:\n{exc}",
            )
            return

        self.current_config_path = path
        # Update window title to show the file name
        base = os.path.basename(path)
        self.setWindowTitle(f"Ash's Cable Tray Calculator – {base}")


    # ------------------------------------------------------------------
    # About dialog
    # ------------------------------------------------------------------

    def _show_about_dialog(self) -> None:
        """
        Show a sassy little About box explaining why this exists.
        """


        text = (
            "<b>Ash's Cable Tray Calculator</b><br><br>"
            "I made this steaming miracle of rage because if I had to fix one more "
            "cable tray Excel sheet held together with spiderwebs, dreams and duct tape "
            "I was going to commit a hate crime.<br><br>"
            "Let’s be honest: if you're here, you've clearly fucked up somewhere in life.<br><br>"
            "Because nobody – <i>nobody</i> – wakes up dreaming about calculating cable tray "
            "loading like it’s the sexy part of engineering.<br><br>"
            "But here you are, scrolling the different weights of cable trays, "
            "pretending gravity isn’t about to rearrange your spine or Civil Defense isn't "
            "about to breed you harder than a Norfolk sister.<br><br>"
            "<ul>"
            "<li>Excel betrayed you like the dirty little spreadsheet whore it is.</li>"
            "<li>You were given formulas written by someone with a tenuous grasp on physics.</li>"
            "<li>Half the cells reference other sheets that don’t even fucking exist anymore.</li>"
            "<li>One workbook tried so hard to divide by zero it's actually classed as a war crime.</li>"
            "</ul>"
            "<br>"
            "So I snapped harder than Epstein in a cell and wrote this monstrosity instead:<br><br>"
            "<ul>"
            "<li>Dark mode because your retinas are already cooked from years of fluorescent despair.</li>"
            "<li>Presets because nobody wants to measure cable diameters like a caveman.</li>"
            "<li>Warnings so gravity and/or Civil Defense don’t rearrange your internal organs "
            "(see that’s a sex joke but also a suspended loads joke... very deep...)</li>"
            "<li>Outputs a PDF report so you can pretend you're a professional, even if just for a second.</li>"
            "</ul>"
            "<br>"
            "And now you, yes YOU, are using it – which means one thing:<br>"
            "<b>Your life choices have led you directly into this situation and you only have yourself to blame.</b><br>"
            "Congratulations, champ.<br><br>"
            "If this app saves you from even one meltdown, one catastrophic miscalculation, "
            "or one supervisor breathing down your neck asking ‘is the tray okay?’ like a lost puppy,<br><br>"
            "<b>Then I have done God's work and you owe me a beer!!</b>"
        )



        msg = QtWidgets.QMessageBox(self)
        msg.setWindowTitle("About Ash's Cable Tray Calculator")
        msg.setText(text)
        msg.setIcon(QtWidgets.QMessageBox.Information)

        # Force a wider box (adjust as you like)
        msg.setStyleSheet("QLabel{min-width: 800px;}")  

        msg.exec_()


    # ------------------------------------------------------------------
    # PDF export
    # ------------------------------------------------------------------

    def _on_export_pdf_clicked(self) -> None:
        """
        Let the user choose a filename, then export a nicely formatted PDF
        with all tray + cable calculations, and open it automatically.
        """
        cables_with_qty = self._collect_cables_from_table()
        if not cables_with_qty:
            QtWidgets.QMessageBox.warning(
                self,
                "Nothing to export",
                "There are no cables defined in the tray. Add at least one cable "
                "before exporting a report.",
            )
            return

        # Suggest a default filename
        default_name = "tray_report.pdf"
        start_dir = self.current_config_path or os.getcwd()
        default_path = os.path.join(start_dir, default_name)

        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save tray calculation report as PDF",
            default_path,
            "PDF files (*.pdf);;All files (*.*)",
        )
        if not path:
            return

        # Ensure .pdf extension
        root, ext = os.path.splitext(path)
        if not ext:
            path = root + ".pdf"

        try:
            self._generate_pdf_report(path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(
                self,
                "PDF export failed",
                f"An error occurred while generating the PDF:\n{exc}",
            )
            return

        # Try to open the PDF automatically
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                QtGui.QDesktopServices.openUrl(
                    QtCore.QUrl.fromLocalFile(path)
                )
        except Exception:
            # Non-fatal: just don't auto-open
            pass

        QtWidgets.QMessageBox.information(
            self,
            "PDF created",
            f"Tray calculation report has been saved to:\n{path}",
        )

    # ------------------------------------------------------------------
    # CSV export
    # ------------------------------------------------------------------

    def _on_export_csv_clicked(self) -> None:
        """
        Let the user choose a filename, then export tray config, summary
        stats, and cable list as a CSV file.
        """
        cables_with_qty = self._collect_cables_from_table()
        if not cables_with_qty:
            QtWidgets.QMessageBox.warning(
                self,
                "Nothing to export",
                "There are no cables defined in the tray. Add at least one cable "
                "before exporting a CSV report.",
            )
            return

        # Suggest a default filename
        default_name = "tray_report.csv"
        start_dir = self.current_config_path or os.getcwd()
        default_path = os.path.join(start_dir, default_name)

        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save tray calculation report as CSV",
            default_path,
            "CSV files (*.csv);;All files (*.*)",
        )
        if not path:
            return

        root, ext = os.path.splitext(path)
        if not ext:
            path = root + ".csv"

        try:
            self._export_csv_report(path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(
                self,
                "CSV export failed",
                f"An error occurred while generating the CSV:\n{exc}",
            )
            return

        QtWidgets.QMessageBox.information(
            self,
            "CSV created",
            f"Tray calculation CSV report has been saved to:\n{path}",
        )

    def _export_csv_report(self, path: str) -> None:
        """
        Export the current tray configuration, summary statistics, and
        cable list to a CSV file.

        Structure is deliberately simple and Excel-friendly:
        - A header block with tray info
        - A summary block with key calculations
        - A cable list table
        """
        cables_with_qty = self._collect_cables_from_table()
        tray = self._build_tray_from_fields()
        stats = compute_cable_tray_stats(cables_with_qty, tray)

        total_cable_weight = stats["total_cable_weight_kg_per_m"]
        tray_self_weight = stats["tray_self_weight_kg_per_m"]
        total_weight = stats["total_weight_kg_per_m"]
        allowable_load = stats["allowable_load_kg_per_m"]
        struct_util_pct = stats["structural_utilisation_percent"]
        cable_area = stats["total_cable_area_mm2"]
        tray_area = stats["tray_usable_area_mm2"]
        area_fill_pct = stats["area_fill_percent"]
        area_fill_limit = stats["recommended_max_area_fill_percent"]

        overloaded_struct = allowable_load > 0 and total_cable_weight > allowable_load
        overloaded_area = area_fill_pct > area_fill_limit

        # Status strings similar to what the GUI shows
        if not cables_with_qty:
            status_text = "No cables defined."
        elif overloaded_struct and overloaded_area:
            status_text = "OVERLOADED: structural + fill limits exceeded"
        elif overloaded_struct:
            status_text = "OVERLOADED: structural limit exceeded"
        elif overloaded_area:
            status_text = "WARNING: area fill above recommended limit"
        else:
            status_text = "OK: within structural and fill limits"

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Write CSV
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

            # Header / metadata block
            writer.writerow(["Ash's Cable Tray Calculator CSV Report"])
            writer.writerow([f"Generated at", timestamp])
            writer.writerow([])

            # Tray configuration
            writer.writerow(["Tray configuration"])
            writer.writerow(["Tray name", tray.name])
            writer.writerow(["Width (mm)", f"{tray.width_mm:.1f}"])
            writer.writerow(["Side height (mm)", f"{tray.height_mm:.1f}"])
            writer.writerow(["Tray self weight (kg/m)", f"{tray.self_weight_kg_per_m:.3f}"])
            writer.writerow(["Maximum allowable load (kg/m)", f"{tray.max_load_kg_per_m:.1f}"])
            writer.writerow([
                "Maximum fill ratio",
                f"{tray.max_fill_ratio:.2f} (recommended {area_fill_limit:.1f} % area fill)",
            ])
            writer.writerow([])

            # Summary stats
            writer.writerow(["Summary"])
            writer.writerow(["Cable weight (kg/m)", f"{total_cable_weight:.3f}"])
            writer.writerow(["Tray self weight (kg/m)", f"{tray_self_weight:.3f}"])
            writer.writerow(["Total weight (kg/m)", f"{total_weight:.3f}"])
            writer.writerow(["Allowable load (kg/m)", f"{allowable_load:.1f}"])
            writer.writerow(["Structural utilisation (%)", f"{struct_util_pct:.1f}"])
            writer.writerow(["Total cable area (mm2)", f"{cable_area:,.0f}"])
            writer.writerow(["Tray usable area (mm2)", f"{tray_area:,.0f}"])
            writer.writerow(["Area fill (%)", f"{area_fill_pct:.1f}"])
            writer.writerow(["Recommended max area fill (%)", f"{area_fill_limit:.1f}"])
            writer.writerow(["Overall status", status_text])
            writer.writerow([])

            # Cables table
            writer.writerow(["Cables in tray"])
            writer.writerow([
                "Cable name",
                "Diameter (mm)",
                "Weight (kg/m)",
                "Quantity",
                "Total weight (kg/m)",
                "Total area (mm2)",
            ])

            for cable, qty in cables_with_qty:
                total_w = cable.weight_kg_per_m * qty
                radius_mm = cable.diameter_mm / 2.0
                area_mm2 = math.pi * (radius_mm ** 2) * qty

                writer.writerow([
                    cable.name,
                    f"{cable.diameter_mm:.1f}",
                    f"{cable.weight_kg_per_m:.3f}",
                    qty,
                    f"{total_w:.3f}",
                    f"{area_mm2:,.0f}",
                ])



    def _generate_pdf_report(self, path: str) -> None:
        """
        Generate a nicely formatted PDF report with all current inputs and
        calculated stats.

        The PDF document title is set to:
            "Ash's Tray Calculation Report"
        so browsers show that in the tab.
        """
        cables_with_qty = self._collect_cables_from_table()
        tray = self._build_tray_from_fields()
        stats = compute_cable_tray_stats(cables_with_qty, tray)

        # Unpack stats
        total_cable_weight = stats["total_cable_weight_kg_per_m"]
        tray_self_weight = stats["tray_self_weight_kg_per_m"]
        total_weight = stats["total_weight_kg_per_m"]
        allowable_load = stats["allowable_load_kg_per_m"]
        struct_util_pct = stats["structural_utilisation_percent"]
        cable_area = stats["total_cable_area_mm2"]
        tray_area = stats["tray_usable_area_mm2"]
        area_fill_pct = stats["area_fill_percent"]
        area_fill_limit = stats["recommended_max_area_fill_percent"]

        overloaded_struct = allowable_load > 0 and total_cable_weight > allowable_load
        overloaded_area = area_fill_pct > area_fill_limit

        # Prepare PDF canvas
        c = canvas.Canvas(path, pagesize=A4)
        c.setTitle("Ash's Tray Calculation Report")  # browser tab title

        width, height = A4
        margin = 20 * mm
        y = height - margin

        # Colours
        dark_blue = (11 / 255.0, 31 / 255.0, 59 / 255.0)
        accent_orange = (1.0, 0.55, 0.0)
        grey = (0.6, 0.6, 0.6)
        light_grey = (0.85, 0.85, 0.85)
        red = (0.9, 0.2, 0.2)
        green = (0.2, 0.7, 0.3)
        black = (0, 0, 0)

        # Helper: draw section title
        def section_heading(text: str, ypos: float) -> float:
            c.setFillColor(dark_blue)
            c.setFont("Helvetica-Bold", 14)
            c.drawString(margin, ypos, text)
            ypos -= 4 * mm
            c.setStrokeColor(accent_orange)
            c.setLineWidth(1.0)
            c.line(margin, ypos, width - margin, ypos)
            return ypos - 6 * mm

        # Helper: maybe start new page
        def ensure_space(current_y: float, needed: float) -> float:
            if current_y - needed < margin:
                c.showPage()
                c.setTitle("Ash's Tray Calculation Report")
                return height - margin
            return current_y

        # Title block
        c.setFillColor(dark_blue)
        c.setFont("Helvetica-Bold", 20)
        c.drawString(margin, y, "Ash's Cable Tray Calculation Report")
        y -= 10 * mm

        c.setFont("Helvetica", 9)
        c.setFillColor(black)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.drawString(margin, y, f"Generated: {timestamp}")
        y -= 6 * mm
        c.drawString(margin, y, f"Tray: {tray.name}")
        y -= 10 * mm

        # Tray section
        y = section_heading("Tray Configuration", y)

        c.setFont("Helvetica", 10)
        c.setFillColor(black)
        line_h = 5 * mm

        tray_lines = [
            f"Tray name: {tray.name}",
            f"Width: {tray.width_mm:.1f} mm",
            f"Side height: {tray.height_mm:.1f} mm",
            f"Tray self weight: {tray.self_weight_kg_per_m:.3f} kg/m",
            f"Maximum allowable load: {tray.max_load_kg_per_m:.1f} kg/m",
            f"Maximum fill ratio: {tray.max_fill_ratio:.2f} (i.e. {area_fill_limit:.1f} % recommended)",
        ]
        for line in tray_lines:
            y = ensure_space(y, line_h)
            c.drawString(margin + 5 * mm, y, line)
            y -= line_h

        y -= 4 * mm

        # Structural summary
        y = section_heading("Structural Load Summary", y)

        # Use colour for utilisation
        util_color = red if overloaded_struct else (green if struct_util_pct <= 80.0 else dark_blue)
        c.setFont("Helvetica", 10)

        y = ensure_space(y, 5 * line_h)

        c.setFillColor(black)
        c.drawString(
            margin + 5 * mm,
            y,
            f"Cable weight: {total_cable_weight:.3f} kg/m",
        )
        y -= line_h
        c.drawString(
            margin + 5 * mm,
            y,
            f"Tray self weight: {tray_self_weight:.3f} kg/m",
        )
        y -= line_h
        c.drawString(
            margin + 5 * mm,
            y,
            f"Total weight: {total_weight:.3f} kg/m",
        )
        y -= line_h
        c.drawString(
            margin + 5 * mm,
            y,
            f"Allowable load: {allowable_load:.1f} kg/m",
        )
        y -= line_h

        c.setFillColor(util_color)
        c.drawString(
            margin + 5 * mm,
            y,
            f"Structural utilisation: {struct_util_pct:.1f} %",
        )
        y -= line_h

        # Structural status line
        c.setFillColor(red if overloaded_struct else green)
        struct_status = (
            "OVERLOADED – check tray sizing / loading assumptions."
            if overloaded_struct
            else "OK – within structural loading limits (based on current assumptions)."
        )
        y = ensure_space(y, line_h)
        c.drawString(margin + 5 * mm, y, struct_status)
        y -= 8 * mm

        # Fill summary
        y = section_heading("Fill and Area Utilisation", y)

        y = ensure_space(y, 4 * line_h)

        c.setFillColor(black)
        c.drawString(
            margin + 5 * mm,
            y,
            f"Total cable area: {cable_area:,.0f} mm^2",
        )
        y -= line_h
        c.drawString(
            margin + 5 * mm,
            y,
            f"Tray usable area: {tray_area:,.0f} mm^2",
        )
        y -= line_h

        fill_color = red if overloaded_area else (green if area_fill_pct <= 0.8 * area_fill_limit else dark_blue)
        c.setFillColor(fill_color)
        c.drawString(
            margin + 5 * mm,
            y,
            f"Area fill: {area_fill_pct:.1f} %",
        )
        y -= line_h

        c.setFillColor(black)
        y = ensure_space(y, line_h)
        c.drawString(
            margin + 5 * mm,
            y,
            f"Recommended maximum fill: {area_fill_limit:.1f} %",
        )
        y -= line_h

        c.setFillColor(red if overloaded_area else green)
        fill_status = (
            "WARNING – area fill exceeds recommended limit."
            if overloaded_area
            else "OK – area fill within recommended limit."
        )
        y = ensure_space(y, line_h)
        c.drawString(margin + 5 * mm, y, fill_status)
        y -= 10 * mm

        # New page for cable table if needed
        y = ensure_space(y, 60 * mm)
        if y < height / 2:
            c.showPage()
            c.setTitle("Ash's Tray Calculation Report")
            y = height - margin

        # Cable table section
        y = section_heading("Cables in Tray", y)

        # Table header + dynamic column widths to keep within page
        headers = ["Cable", "Diameter (mm)", "Weight (kg/m)", "Qty", "Total (kg/m)", "Area (mm^2)"]

        # Base widths in mm (we'll scale them if needed)
        base_col_widths_mm = [60.0, 25.0, 25.0, 15.0, 30.0, 30.0]
        total_width_mm = sum(base_col_widths_mm)
        available_width_mm = (width - 2 * margin) / mm  # convert to mm

        if total_width_mm > available_width_mm:
            scale = available_width_mm / total_width_mm
            base_col_widths_mm = [w * scale for w in base_col_widths_mm]

        col_widths = [w * mm for w in base_col_widths_mm]
        table_width = sum(col_widths)

        table_x = margin  # could center if you want: margin + (width - 2*margin - table_width)/2
        header_h = 7 * mm
        row_h = 6 * mm

        # Header background
        c.setFillColor(light_grey)
        c.rect(table_x, y - header_h, table_width, header_h, stroke=0, fill=1)
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 9)

        x = table_x + 2 * mm
        for i, htxt in enumerate(headers):
            c.drawString(x, y - header_h + 2 * mm, htxt)
            x += col_widths[i]

        y -= header_h

        # Table rows
        c.setFont("Helvetica", 8)
        c.setStrokeColor(grey)
        row_index = 0

        for cable, qty in cables_with_qty:
            y = ensure_space(y, row_h + 2 * mm)

            # Alternating row background
            bg = light_grey if row_index % 2 == 0 else (1, 1, 1)
            c.setFillColor(bg)
            c.rect(table_x, y - row_h, table_width, row_h, stroke=0, fill=1)

            # Compute per-cable info
            total_w = cable.weight_kg_per_m * qty
            radius_mm = cable.diameter_mm / 2.0
            area_mm2 = math.pi * (radius_mm ** 2) * qty

            # Row text
            c.setFillColor(black)
            x = table_x + 2 * mm
            c.drawString(x, y - row_h + 2 * mm, cable.name[:40])
            x += col_widths[0]
            c.drawRightString(x + col_widths[1] - 3 * mm, y - row_h + 2 * mm, f"{cable.diameter_mm:.1f}")
            x += col_widths[1]
            c.drawRightString(x + col_widths[2] - 3 * mm, y - row_h + 2 * mm, f"{cable.weight_kg_per_m:.3f}")
            x += col_widths[2]
            c.drawRightString(x + col_widths[3] - 3 * mm, y - row_h + 2 * mm, str(qty))
            x += col_widths[3]
            c.drawRightString(x + col_widths[4] - 3 * mm, y - row_h + 2 * mm, f"{total_w:.3f}")
            x += col_widths[4]
            c.drawRightString(x + col_widths[5] - 3 * mm, y - row_h + 2 * mm, f"{area_mm2:,.0f}")

            # Row divider
            c.setStrokeColor(grey)
            c.setLineWidth(0.3)
            c.line(table_x, y - row_h, table_x + table_width, y - row_h)

            y -= row_h
            row_index += 1


        # Table rows
        c.setFont("Helvetica", 8)
        c.setStrokeColor(grey)
        for cable, qty in cables_with_qty:
            y = ensure_space(y, row_h + 2 * mm)
            # row background stripes
            c.setFillColor(light_grey if int((height - y) / row_h) % 2 == 0 else (1, 1, 1))
            c.rect(table_x, y - row_h, sum(col_widths), row_h, stroke=0, fill=1)

            # Compute per-cable info
            total_w = cable.weight_kg_per_m * qty
            radius_mm = cable.diameter_mm / 2.0
            area_mm2 = math.pi * (radius_mm ** 2) * qty

            c.setFillColor(black)
            x = table_x + 2 * mm
            c.drawString(x, y - row_h + 2 * mm, cable.name[:40])
            x += col_widths[0]
            c.drawRightString(x + col_widths[1] - 3 * mm, y - row_h + 2 * mm, f"{cable.diameter_mm:.1f}")
            x += col_widths[1]
            c.drawRightString(x + col_widths[2] - 3 * mm, y - row_h + 2 * mm, f"{cable.weight_kg_per_m:.3f}")
            x += col_widths[2]
            c.drawRightString(x + col_widths[3] - 3 * mm, y - row_h + 2 * mm, str(qty))
            x += col_widths[3]
            c.drawRightString(x + col_widths[4] - 3 * mm, y - row_h + 2 * mm, f"{total_w:.3f}")
            x += col_widths[4]
            c.drawRightString(x + col_widths[5] - 3 * mm, y - row_h + 2 * mm, f"{area_mm2:,.0f}")

            # Row divider
            c.setStrokeColor(grey)
            c.setLineWidth(0.3)
            c.line(table_x, y - row_h, table_x + sum(col_widths), y - row_h)

            y -= row_h

        # Footer / note
        y = ensure_space(y, 15 * mm)
        c.setFont("Helvetica-Oblique", 8)
        c.setFillColor(grey)
        c.drawString(
            margin,
            margin + 5 * mm,
            "Note: Values are based on current tray and cable data in the calculator. "
            "Always verify against manufacturer data and applicable standards.",
        )

        c.showPage()
        c.save()


