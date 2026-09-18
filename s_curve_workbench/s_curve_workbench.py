"""
Parametric S-Curve Workbench - Phase 1 Implementation
Desktop application for generating time-phased expenditure profiles using parametric S-curves.

Technology Stack: Python 3.12 + PySide6 + matplotlib + openpyxl
"""

import sys
import json
import math
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any

import numpy as np
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QGroupBox, QLabel, QLineEdit, QPushButton, QSlider, 
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QFileDialog, QFrame, QScrollArea, QComboBox, QDateEdit,
    QDoubleSpinBox, QSpinBox, QTextEdit, QSplitter, QToolBar,
    QStatusBar, QMenu, QMenuBar, QTabWidget, QFormLayout, QInputDialog,
    QAbstractItemView
)
from PySide6.QtGui import QAction, QFont, QColor, QPalette, QActionGroup
from PySide6.QtCore import Qt, Signal, Slot, QDate, QTimer
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import openpyxl
from openpyxl.utils import get_column_letter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('s_curve_workbench.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================================
# CORE MATHEMATICS (Normative - Do Not Modify)
# ============================================================================

def compute_F(T: float, A: float, B: float) -> float:
    """
    Compute the cumulative percentage F(T) using the normative polynomial formula.
    
    F(T) = 10·T²·(1−T)²·(A + B·T) + T⁴·(5 − 4·T), for 0<T<1
    F(T) = 0 for T≤0
    F(T) = 1 for T≥1
    
    This implements a parametric S-curve where:
    - A controls front-loading (higher A = more spend early)  
    - B controls back-loading (higher B = more spend late)
    
    The formula ensures:
    - F(0) = 0, F(1) = 1 (boundary conditions)
    - Smooth S-curve behavior for admissible parameters
    
    Parameters:
        T: Normalized time in [0, 1]
        A: Shape parameter A (range typically [0, 1])
        B: Shape parameter B (range typically [0, 1])
    
    Returns:
        Cumulative percentage F(T) in [0, 1]
    """
    # Handle boundary conditions
    if T <= 0:
        return 0.0
    if T >= 1:
        return 1.0
    
    # Normative formula: F(T) = 10·T²·(1−T)²·(A + B·T) + T⁴·(5 − 4·T)
    T_sq = T * T
    T_fourth = T_sq * T_sq
    one_minus_T = 1.0 - T
    one_minus_T_sq = one_minus_T * one_minus_T
    
    term1 = 10.0 * T_sq * one_minus_T_sq * (A + B * T)
    term2 = T_fourth * (5.0 - 4.0 * T)
    
    result = term1 + term2
    
    return result


def compute_monthly_rate(T_start: float, T_end: float, A: float, B: float) -> float:
    """
    Compute the monthly expenditure rate as the difference in cumulative values.
    
    Monthly rate = F(T_end) - F(T_start)
    
    Parameters:
        T_start: Start of period (normalized)
        T_end: End of period (normalized)
        A, B: Shape parameters
    
    Returns:
        Monthly expenditure rate (fraction of total budget)
    """
    return compute_F(T_end, A, B) - compute_F(T_start, A, B)


def generate_s_curve_data(budget: float, n_months: int, A: float, B: float) -> Dict[str, List]:
    """
    Generate complete S-curve data for all periods.
    
    This is the SINGLE SOURCE OF TRUTH - all charts, tables, and exports
    must use this computed array to avoid desynchronization.
    
    Parameters:
        budget: Total budget C
        n_months: Number of months N
        A, B: Shape parameters
    
    Returns:
        Dictionary containing:
            - periods: List of period numbers (1 to N)
            - monthly_amounts: Monthly expenditure amounts
            - monthly_percentages: Monthly percentages
            - cumulative_amounts: Cumulative expenditure amounts
            - cumulative_percentages: Cumulative percentages
            - T_values: Normalized time points at period boundaries
    """
    periods = list(range(1, n_months + 1))
    monthly_amounts = []
    cumulative_amounts = []
    monthly_percentages = []
    cumulative_percentages = []
    T_values = [0.0]  # T at start (period 0)
    
    cumulative = 0.0
    for i in range(n_months):
        T_start = i / n_months
        T_end = (i + 1) / n_months
        T_values.append(T_end)
        
        # Compute monthly rate using differencing
        monthly_pct = compute_monthly_rate(T_start, T_end, A, B)
        monthly_amt = monthly_pct * budget
        
        monthly_percentages.append(monthly_pct * 100)  # As percentage
        monthly_amounts.append(monthly_amt)
        
        cumulative += monthly_amt
        cumulative_amounts.append(cumulative)
        cumulative_percentages.append((cumulative / budget) * 100 if budget != 0 else 0)
    
    # Ensure final cumulative equals budget exactly (reconciliation)
    if cumulative_amounts:
        cumulative_amounts[-1] = budget
        cumulative_percentages[-1] = 100.0
    
    return {
        'periods': periods,
        'monthly_amounts': monthly_amounts,
        'monthly_percentages': monthly_percentages,
        'cumulative_amounts': cumulative_amounts,
        'cumulative_percentages': cumulative_percentages,
        'T_values': T_values,
        'F_values': [compute_F(t, A, B) for t in T_values]
    }


def get_milestone_summary(A: float, B: float) -> Dict[float, float]:
    """
    Compute milestone summary at standard checkpoints.
    
    Returns F(T) at T = 0.10, 0.25, 0.50, 0.65, 0.75, 0.90 as percentages.
    """
    milestones = [0.10, 0.25, 0.50, 0.65, 0.75, 0.90]
    return {t: compute_F(t, A, B) * 100 for t in milestones}


# ============================================================================
# QA / GOVERNANCE ENGINE
# ============================================================================

class QAResult:
    """Container for QA test results."""
    
    def __init__(self):
        self.admissibility_pass = True
        self.boundary_pass = True
        self.range_pass = True
        self.monotonicity_pass = True
        self.period_test_pass = True
        self.reconciliation_pass = True
        self.errors: List[str] = []
        self.warnings: List[str] = []
        self.min_monthly: float = 0.0
        self.reconciliation_diff: float = 0.0
    
    @property
    def overall_pass(self) -> bool:
        return all([
            self.admissibility_pass,
            self.boundary_pass,
            self.range_pass,
            self.monotonicity_pass,
            self.period_test_pass,
            self.reconciliation_pass
        ])


def run_qa_checks(budget: float, n_months: int, A: float, B: float, 
                  data: Dict[str, List]) -> QAResult:
    """
    Run comprehensive QA checks on the S-curve parameters and generated data.
    
    Tests:
        1. Admissibility: 0 ≤ A and A+B ≤ 1
        2. Boundary: F(0) ≈ 0, F(1) ≈ 1
        3. Range: 0 ≤ F(T) ≤ 1 for all T in grid
        4. Monotonicity: No negative increments
        5. Period test: Minimum monthly ≥ 0
        6. Reconciliation: Σ Monthly vs C
    """
    result = QAResult()
    
    # Test 1: Admissibility
    if A < 0:
        result.admissibility_pass = False
        result.errors.append(f"Admissibility FAIL: A={A:.4f} < 0")
    
    if A + B > 1:
        result.admissibility_pass = False
        result.errors.append(f"Admissibility FAIL: A+B={A+B:.4f} > 1")
    
    # Test 2: Boundary conditions
    F_0 = compute_F(0.0, A, B)
    F_1 = compute_F(1.0, A, B)
    
    if abs(F_0) > 1e-9:
        result.boundary_pass = False
        result.errors.append(f"Boundary FAIL: F(0)={F_0:.2e} ≠ 0")
    
    if abs(F_1 - 1.0) > 1e-9:
        result.boundary_pass = False
        result.errors.append(f"Boundary FAIL: F(1)={F_1:.10f} ≠ 1")
    
    # Test 3: Range check on 1000-point grid
    for i in range(1001):
        T = i / 1000.0
        F_T = compute_F(T, A, B)
        if F_T < -1e-9 or F_T > 1.0 + 1e-9:
            result.range_pass = False
            result.errors.append(f"Range FAIL: F({T:.3f})={F_T:.6f} outside [0,1]")
            break
    
    # Test 4: Monotonicity (successive differences)
    F_prev = compute_F(0.0, A, B)
    for i in range(1, 1001):
        T = i / 1000.0
        F_T = compute_F(T, A, B)
        if F_T < F_prev - 1e-9:
            result.monotonicity_pass = False
            result.errors.append(f"Monotonicity FAIL: Decrease at T={T:.3f}")
            break
        F_prev = F_T
    
    # Also check period boundaries
    for i in range(len(data['cumulative_percentages']) - 1):
        if data['cumulative_percentages'][i+1] < data['cumulative_percentages'][i] - 1e-9:
            result.monotonicity_pass = False
            result.errors.append(f"Monotonicity FAIL: Period {i+1} to {i+2}")
            break
    
    # Test 5: Period test (minimum monthly >= 0)
    if data['monthly_amounts']:
        result.min_monthly = min(data['monthly_amounts'])
        if result.min_monthly < -0.01:  # Allow small floating point errors
            result.period_test_pass = False
            result.errors.append(f"Period test FAIL: Min monthly = {result.min_monthly:.2f}")
    
    # Test 6: Reconciliation
    total_monthly = sum(data['monthly_amounts'])
    result.reconciliation_diff = abs(total_monthly - budget)
    if result.reconciliation_diff > 0.01:
        result.reconciliation_pass = False
        result.errors.append(f"Reconciliation FAIL: Diff = {result.reconciliation_diff:.2f}")
    
    return result


# ============================================================================
# AUDIT TRAIL
# ============================================================================

class AuditTrail:
    """Manages timestamped audit log of parameter changes."""
    
    def __init__(self, log_file: str = "audit_trail.json"):
        self.log_file = Path(log_file)
        self.entries: List[Dict] = []
        self._load()
    
    def _load(self):
        """Load existing audit trail from file."""
        if self.log_file.exists():
            try:
                with open(self.log_file, 'r') as f:
                    self.entries = json.load(f)
            except (json.JSONDecodeError, IOError):
                self.entries = []
    
    def _save(self):
        """Save audit trail to file."""
        with open(self.log_file, 'w') as f:
            json.dump(self.entries, f, indent=2, default=str)
    
    def log_change(self, event_type: str, old_value: Any, new_value: Any, 
                   rationale: str = "", approver: str = "System"):
        """Log a parameter change."""
        entry = {
            'timestamp': datetime.now().isoformat(),
            'event_type': event_type,
            'old_value': old_value,
            'new_value': new_value,
            'rationale': rationale,
            'approver': approver,
            'version': "1.0.0"
        }
        self.entries.append(entry)
        self._save()
        logger.info(f"Audit: {event_type} - {old_value} → {new_value}")
    
    def log_coefficient_selection(self, package_name: str, basis: str, 
                                   A: float, B: float, admissible: bool,
                                   rationale: str = "", approver: str = "System"):
        """Log coefficient selection record."""
        entry = {
            'timestamp': datetime.now().isoformat(),
            'event_type': 'coefficient_selection',
            'package': package_name,
            'basis': basis,
            'A': A,
            'B': B,
            'admissible': admissible,
            'rationale': rationale,
            'approver': approver,
            'version': "1.0.0"
        }
        self.entries.append(entry)
        self._save()
    
    def get_export_record(self) -> List[Dict]:
        """Get audit trail for export."""
        return self.entries


# ============================================================================
# PRESETS
# ============================================================================

PRESETS = {
    "JSC": {"scale_A": 66, "scale_B": 84, "A": 0.32, "B": 0.68},
    "Rear Loaded": {"scale_A": 50, "scale_B": 50, "A": 0.00, "B": 0.00},
    "Front Loaded": {"scale_A": 100, "scale_B": 50, "A": 1.00, "B": 0.00},
    "Mid Loaded": {"scale_A": 50, "scale_B": 100, "A": 0.00, "B": 1.00},
    "Balanced": {"scale_A": 75, "scale_B": 50, "A": 0.50, "B": 0.00},
}


def scale_to_param(scale: float) -> float:
    """Convert slider scale (0-100) to parameter (-1 to +1)."""
    return (scale - 50) / 50.0


def param_to_scale(param: float) -> float:
    """Convert parameter (-1 to +1) to slider scale (0-100)."""
    return 50 * (param + 1)


# ============================================================================
# EXCEL EXPORT
# ============================================================================

def export_to_excel(data: Dict, inputs: Dict, qa_result: QAResult, 
                    audit_trail: AuditTrail, filepath: str):
    """
    Export all data to Excel workbook with multiple sheets.
    
    Sheets: Inputs, Monthly, Cumulative, Milestones, QA_Log, AnnexB_Record
    """
    wb = openpyxl.Workbook()
    
    # Sheet 1: Inputs
    ws_inputs = wb.active
    ws_inputs.title = "Inputs"
    ws_inputs.append(["Parameter", "Value"])
    ws_inputs.append(["Budget (C)", inputs['budget']])
    ws_inputs.append(["Duration (N months)", inputs['n_months']])
    ws_inputs.append(["Start Date", inputs['start_date']])
    ws_inputs.append(["Shape Parameter A", inputs['A']])
    ws_inputs.append(["Shape Parameter B", inputs['B']])
    ws_inputs.append(["Scale A", inputs['scale_A']])
    ws_inputs.append(["Scale B", inputs['scale_B']])
    ws_inputs.append(["A + B", inputs['A'] + inputs['B']])
    
    # Sheet 2: PackageRegister (placeholder for Phase 1)
    ws_pkg = wb.create_sheet("PackageRegister")
    ws_pkg.append(["ID", "Description", "Cost Basis", "Budget", "Start", "Finish", "Profile", "Owner"])
    ws_pkg.append(["PKG-001", "Single Package (Phase 1)", "Budget", inputs['budget'], 
                   inputs['start_date'], "TBD", "S-Curve", "Default"])
    
    # Sheet 3: Overlays (placeholder for Phase 2)
    ws_overlay = wb.create_sheet("Overlays")
    ws_overlay.append(["Package ID", "Date", "Amount", "Evidence Ref", "Type"])
    
    # Sheet 4: Monthly
    ws_monthly = wb.create_sheet("Monthly")
    ws_monthly.append(["Period", "Monthly Amount", "Monthly %"])
    for i, (period, amt, pct) in enumerate(zip(
            data['periods'], data['monthly_amounts'], data['monthly_percentages'])):
        ws_monthly.append([period, amt, pct])
    
    # Sheet 5: Cumulative
    ws_cumulative = wb.create_sheet("Cumulative")
    ws_cumulative.append(["Period", "Cumulative Amount", "Cumulative %"])
    for period, amt, pct in zip(
            data['periods'], data['cumulative_amounts'], data['cumulative_percentages']):
        ws_cumulative.append([period, amt, pct])
    
    # Sheet 6: Milestones
    ws_milestones = wb.create_sheet("Milestones")
    ws_milestones.append(["T", "Cumulative %", "Description"])
    milestone_data = get_milestone_summary(inputs['A'], inputs['B'])
    for t, pct in sorted(milestone_data.items()):
        ws_milestones.append([t, pct, f"T={t:.2f}"])
    
    # Sheet 7: QA_Log
    ws_qa = wb.create_sheet("QA_Log")
    ws_qa.append(["Test", "Status", "Details"])
    ws_qa.append(["Admissibility", "PASS" if qa_result.admissibility_pass else "FAIL", 
                  "; ".join([e for e in qa_result.errors if "Admissibility" in e])])
    ws_qa.append(["Boundary", "PASS" if qa_result.boundary_pass else "FAIL",
                  "; ".join([e for e in qa_result.errors if "Boundary" in e])])
    ws_qa.append(["Range", "PASS" if qa_result.range_pass else "FAIL",
                  "; ".join([e for e in qa_result.errors if "Range" in e])])
    ws_qa.append(["Monotonicity", "PASS" if qa_result.monotonicity_pass else "FAIL",
                  "; ".join([e for e in qa_result.errors if "Monotonicity" in e])])
    ws_qa.append(["Period Test", "PASS" if qa_result.period_test_pass else "FAIL",
                  "; ".join([e for e in qa_result.errors if "Period" in e])])
    ws_qa.append(["Reconciliation", "PASS" if qa_result.reconciliation_pass else "FAIL",
                  f"Diff: {qa_result.reconciliation_diff:.2f}"])
    ws_qa.append(["Overall", "PASS" if qa_result.overall_pass else "FAIL", ""])
    
    # Sheet 8: AnnexB_Record
    ws_annex = wb.create_sheet("AnnexB_Record")
    ws_annex.append(["Timestamp", "Event Type", "Old Value", "New Value", "Rationale", "Approver", "Version"])
    for entry in audit_trail.get_export_record():
        ws_annex.append([
            entry.get('timestamp', ''),
            entry.get('event_type', ''),
            str(entry.get('old_value', '')),
            str(entry.get('new_value', '')),
            entry.get('rationale', ''),
            entry.get('approver', ''),
            entry.get('version', '')
        ])
    
    # Save workbook
    wb.save(filepath)
    logger.info(f"Exported Excel workbook to {filepath}")


# ============================================================================
# MAIN APPLICATION WINDOW
# ============================================================================

class SCurveWorkbench(QMainWindow):
    """Main application window for the S-Curve Workbench."""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Parametric S-Curve Workbench v1.0")
        self.setMinimumSize(1400, 900)
        
        # Initialize state
        self.budget = 120000000.0
        self.n_months = 24
        self.start_date = QDate.currentDate()
        self.scale_A = 66.0  # JSC preset default
        self.scale_B = 84.0
        self.A = 0.32
        self.B = 0.68
        
        self.audit_trail = AuditTrail()
        self.current_data: Optional[Dict] = None
        self.current_qa: Optional[QAResult] = None
        
        # Log initial state
        self.audit_trail.log_change(
            "initialization", None, 
            f"Budget={self.budget}, N={self.n_months}, A={self.A}, B={self.B}",
            "Application startup"
        )
        
        # Setup UI
        self._setup_ui()
        self._connect_signals()
        self._recalculate()
        
        logger.info("S-Curve Workbench initialized")
    
    def _setup_ui(self):
        """Setup the user interface."""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create toolbar
        self._setup_toolbar()
        
        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # Left panel: Inputs
        left_panel = self._create_input_panel()
        splitter.addWidget(left_panel)
        
        # Right panel: Charts and tables
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # Tab widget for chart and data
        tab_widget = QTabWidget()
        
        # Chart tab
        chart_tab = self._create_chart_tab()
        tab_widget.addTab(chart_tab, "Chart & Analysis")
        
        # Data grid tab
        data_tab = self._create_data_grid_tab()
        tab_widget.addTab(data_tab, "Data Grid")
        
        right_layout.addWidget(tab_widget)
        splitter.addWidget(right_panel)
        
        # Set splitter proportions
        splitter.setSizes([350, 1050])
        
        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Ready")
    
    def _setup_toolbar(self):
        """Setup application toolbar."""
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)
        
        # Export actions
        export_menu = QMenu("Export", self)
        
        excel_action = QAction("Export Excel (.xlsx)", self)
        excel_action.triggered.connect(self._export_excel)
        export_menu.addAction(excel_action)
        
        pdf_action = QAction("Export PDF Report", self)
        pdf_action.triggered.connect(self._export_pdf)
        export_menu.addAction(pdf_action)
        
        json_save_action = QAction("Save Project (.json)", self)
        json_save_action.triggered.connect(self._save_project)
        export_menu.addAction(json_save_action)
        
        json_load_action = QAction("Load Project (.json)", self)
        json_load_action.triggered.connect(self._load_project)
        export_menu.addAction(json_load_action)
        
        toolbar.addWidget(QLabel("Export: "))
        export_button = QPushButton("▼")
        export_button.setMenu(export_menu)
        toolbar.addWidget(export_button)
        
        toolbar.addSeparator()
        
        # Preset buttons in toolbar
        toolbar.addWidget(QLabel("Presets: "))
        for preset_name in PRESETS.keys():
            btn = QPushButton(preset_name)
            btn.clicked.connect(lambda checked, name=preset_name: self._apply_preset(name))
            toolbar.addWidget(btn)
    
    def _create_input_panel(self) -> QWidget:
        """Create the left input panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(15)
        
        # Budget input
        budget_group = QGroupBox("Budget & Duration")
        budget_layout = QFormLayout()
        
        self.budget_input = QDoubleSpinBox()
        self.budget_input.setRange(0, 1e12)
        self.budget_input.setValue(self.budget)
        self.budget_input.setPrefix("$ ")
        self.budget_input.setDecimals(2)
        budget_layout.addRow("Total Budget (C):", self.budget_input)
        
        self.months_input = QSpinBox()
        self.months_input.setRange(1, 120)
        self.months_input.setValue(self.n_months)
        budget_layout.addRow("Duration (N months):", self.months_input)
        
        self.start_date_input = QDateEdit()
        self.start_date_input.setDate(self.start_date)
        self.start_date_input.setCalendarPopup(True)
        budget_layout.addRow("Start Date:", self.start_date_input)
        
        budget_group.setLayout(budget_layout)
        layout.addWidget(budget_group)
        
        # Shape parameters
        shape_group = QGroupBox("Shape Parameters")
        shape_layout = QVBoxLayout()
        
        # Parameter A
        a_layout = QHBoxLayout()
        a_label = QLabel("A:")
        a_label.setFixedWidth(20)
        self.slider_A = QSlider(Qt.Horizontal)
        self.slider_A.setRange(0, 100)
        self.slider_A.setValue(int(self.scale_A))
        self.slider_A.setTickPosition(QSlider.TicksBelow)
        self.slider_A.setTickInterval(10)
        self.input_A = QDoubleSpinBox()
        self.input_A.setRange(-1, 1)
        self.input_A.setValue(self.A)
        self.input_A.setDecimals(2)
        self.input_A.setSingleStep(0.01)
        self.input_A.setFixedWidth(80)
        a_layout.addWidget(a_label)
        a_layout.addWidget(self.slider_A)
        a_layout.addWidget(self.input_A)
        shape_layout.addLayout(a_layout)
        
        # Parameter B
        b_layout = QHBoxLayout()
        b_label = QLabel("B:")
        b_label.setFixedWidth(20)
        self.slider_B = QSlider(Qt.Horizontal)
        self.slider_B.setRange(0, 100)
        self.slider_B.setValue(int(self.scale_B))
        self.slider_B.setTickPosition(QSlider.TicksBelow)
        self.slider_B.setTickInterval(10)
        self.input_B = QDoubleSpinBox()
        self.input_B.setRange(-1, 1)
        self.input_B.setValue(self.B)
        self.input_B.setDecimals(2)
        self.input_B.setSingleStep(0.01)
        self.input_B.setFixedWidth(80)
        b_layout.addWidget(b_label)
        b_layout.addWidget(self.slider_B)
        b_layout.addWidget(self.input_B)
        shape_layout.addLayout(b_layout)
        
        # A+B display
        self.sum_display = QLabel("A + B = 1.00")
        self.sum_display.setFont(QFont("Arial", 10, QFont.Bold))
        shape_layout.addWidget(self.sum_display)
        
        shape_group.setLayout(shape_layout)
        layout.addWidget(shape_group)
        
        # Preset buttons
        preset_group = QGroupBox("Quick Presets")
        preset_layout = QVBoxLayout()
        for preset_name, preset_data in PRESETS.items():
            btn = QPushButton(f"{preset_name} (A={preset_data['A']:.2f}, B={preset_data['B']:.2f})")
            btn.clicked.connect(lambda checked, name=preset_name: self._apply_preset(name))
            preset_layout.addWidget(btn)
        preset_group.setLayout(preset_layout)
        layout.addWidget(preset_group)
        
        layout.addStretch()
        
        return panel
    
    def _create_chart_tab(self) -> QWidget:
        """Create the chart and analysis tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Matplotlib figure
        self.figure = Figure(figsize=(10, 6), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        # QA Panel
        qa_group = QGroupBox("QA / Governance Status")
        qa_layout = QVBoxLayout()
        
        self.qa_status_label = QLabel("QA Status: PASS")
        self.qa_status_label.setFont(QFont("Arial", 14, QFont.Bold))
        self.qa_status_label.setStyleSheet("color: green;")
        qa_layout.addWidget(self.qa_status_label)
        
        self.qa_details = QTextEdit()
        self.qa_details.setReadOnly(True)
        self.qa_details.setMaximumHeight(150)
        qa_layout.addWidget(self.qa_details)
        
        qa_group.setLayout(qa_layout)
        layout.addWidget(qa_group)
        
        # Milestone summary table
        milestone_group = QGroupBox("Milestone Summary")
        milestone_layout = QVBoxLayout()
        
        self.milestone_table = QTableWidget()
        self.milestone_table.setColumnCount(2)
        self.milestone_table.setHorizontalHeaderLabels(["T", "Cumulative %"])
        self.milestone_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.milestone_table.setMaximumHeight(150)
        milestone_layout.addWidget(self.milestone_table)
        
        milestone_group.setLayout(milestone_layout)
        layout.addWidget(milestone_group)
        
        return widget
    
    def _create_data_grid_tab(self) -> QWidget:
        """Create the data grid tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        self.data_grid = QTableWidget()
        self.data_grid.setColumnCount(4)
        self.data_grid.setHorizontalHeaderLabels(["Period", "Monthly Amount", "Cumulative Amount", "Cumulative %"])
        self.data_grid.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.data_grid)
        
        return widget
    
    def _connect_signals(self):
        """Connect UI signals to slots."""
        # Budget and duration
        self.budget_input.valueChanged.connect(self._on_budget_changed)
        self.months_input.valueChanged.connect(self._on_months_changed)
        self.start_date_input.dateChanged.connect(self._on_date_changed)
        
        # Sliders
        self.slider_A.valueChanged.connect(self._on_slider_A_changed)
        self.slider_B.valueChanged.connect(self._on_slider_B_changed)
        
        # Direct inputs
        self.input_A.valueChanged.connect(self._on_input_A_changed)
        self.input_B.valueChanged.connect(self._on_input_B_changed)
    
    @Slot(float)
    def _on_budget_changed(self, value: float):
        """Handle budget change."""
        old_value = self.budget
        self.budget = value
        self.audit_trail.log_change("budget", old_value, value, "User input")
        self._recalculate()
    
    @Slot(int)
    def _on_months_changed(self, value: int):
        """Handle months change."""
        old_value = self.n_months
        self.n_months = value
        self.audit_trail.log_change("duration", old_value, value, "User input")
        self._recalculate()
    
    @Slot(QDate)
    def _on_date_changed(self, date: QDate):
        """Handle start date change."""
        self.start_date = date
        self.audit_trail.log_change("start_date", None, date.toString(), "User input")
    
    @Slot(int)
    def _on_slider_A_changed(self, value: int):
        """Handle slider A change."""
        self.scale_A = float(value)
        self.A = scale_to_param(self.scale_A)
        self.input_A.blockSignals(True)
        self.input_A.setValue(self.A)
        self.input_A.blockSignals(False)
        self._update_sum_display()
        self.audit_trail.log_change("parameter_A", None, self.A, f"Slider A={value}")
        self._recalculate()
    
    @Slot(int)
    def _on_slider_B_changed(self, value: int):
        """Handle slider B change."""
        self.scale_B = float(value)
        self.B = scale_to_param(self.scale_B)
        self.input_B.blockSignals(True)
        self.input_B.setValue(self.B)
        self.input_B.blockSignals(False)
        self._update_sum_display()
        self.audit_trail.log_change("parameter_B", None, self.B, f"Slider B={value}")
        self._recalculate()
    
    @Slot(float)
    def _on_input_A_changed(self, value: float):
        """Handle direct input A change."""
        old_value = self.A
        self.A = value
        self.scale_A = param_to_scale(self.A)
        self.slider_A.blockSignals(True)
        self.slider_A.setValue(int(round(self.scale_A)))
        self.slider_A.blockSignals(False)
        self._update_sum_display()
        self.audit_trail.log_change("parameter_A", old_value, value, "Direct input")
        self._recalculate()
    
    @Slot(float)
    def _on_input_B_changed(self, value: float):
        """Handle direct input B change."""
        old_value = self.B
        self.B = value
        self.scale_B = param_to_scale(self.B)
        self.slider_B.blockSignals(True)
        self.slider_B.setValue(int(round(self.scale_B)))
        self.slider_B.blockSignals(False)
        self._update_sum_display()
        self.audit_trail.log_change("parameter_B", old_value, value, "Direct input")
        self._recalculate()
    
    def _update_sum_display(self):
        """Update the A+B sum display."""
        sum_val = self.A + self.B
        self.sum_display.setText(f"A + B = {sum_val:.2f}")
        if sum_val > 1.0:
            self.sum_display.setStyleSheet("color: red; font-weight: bold;")
        else:
            self.sum_display.setStyleSheet("color: green; font-weight: bold;")
    
    def _apply_preset(self, preset_name: str):
        """Apply a preset configuration."""
        if preset_name not in PRESETS:
            return
        
        preset = PRESETS[preset_name]
        old_A, old_B = self.A, self.B
        
        self.scale_A = preset["scale_A"]
        self.scale_B = preset["scale_B"]
        self.A = preset["A"]
        self.B = preset["B"]
        
        # Update UI without triggering signals
        self.slider_A.blockSignals(True)
        self.slider_B.blockSignals(True)
        self.input_A.blockSignals(True)
        self.input_B.blockSignals(True)
        
        self.slider_A.setValue(int(self.scale_A))
        self.slider_B.setValue(int(self.scale_B))
        self.input_A.setValue(self.A)
        self.input_B.setValue(self.B)
        
        self.slider_A.blockSignals(False)
        self.slider_B.blockSignals(False)
        self.input_A.blockSignals(False)
        self.input_B.blockSignals(False)
        
        self._update_sum_display()
        
        self.audit_trail.log_coefficient_selection(
            "Single Package", "Budget", self.A, self.B, 
            self.A >= 0 and self.A + self.B <= 1,
            f"Preset: {preset_name}"
        )
        
        logger.info(f"Applied preset: {preset_name}")
        self._recalculate()
    
    def _recalculate(self):
        """Recalculate all data and update UI."""
        # Generate S-curve data (single source of truth)
        self.current_data = generate_s_curve_data(self.budget, self.n_months, self.A, self.B)
        
        # Run QA checks
        self.current_qa = run_qa_checks(self.budget, self.n_months, self.A, self.B, self.current_data)
        
        # Update UI components
        self._update_chart()
        self._update_qa_panel()
        self._update_milestone_table()
        self._update_data_grid()
        self.statusBar.showMessage(f"Recalculated: Budget=${self.budget:,.2f}, N={self.n_months} months")
    
    def _update_chart(self):
        """Update the matplotlib chart."""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        periods = self.current_data['periods']
        monthly_pct = self.current_data['monthly_percentages']
        cumulative_pct = self.current_data['cumulative_percentages']
        
        # Create twin axes
        ax_bar = ax
        ax_line = ax.twinx()
        
        # Bar chart for monthly percentages
        bars = ax_bar.bar(periods, monthly_pct, width=0.8, alpha=0.7, 
                          color='steelblue', label='Monthly %')
        ax_bar.set_xlabel('Period (Month)')
        ax_bar.set_ylabel('Monthly Expenditure (%)', color='steelblue')
        ax_bar.tick_params(axis='y', labelcolor='steelblue')
        ax_bar.set_ylim(0, max(monthly_pct) * 1.2 if monthly_pct else 10)
        
        # Line chart for cumulative percentage
        line = ax_line.plot(periods, cumulative_pct, 'r-', linewidth=2, 
                            marker='o', markersize=4, label='Cumulative %')
        ax_line.set_ylabel('Cumulative Expenditure (%)', color='red')
        ax_line.tick_params(axis='y', labelcolor='red')
        ax_line.set_ylim(0, 105)
        ax_line.grid(True, alpha=0.3)
        
        # Title
        ax.set_title(f'S-Curve Profile (A={self.A:.2f}, B={self.B:.2f}, A+B={self.A+self.B:.2f})')
        
        # Combined legend
        lines1, labels1 = ax_bar.get_legend_handles_labels()
        lines2, labels2 = ax_line.get_legend_handles_labels()
        ax.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def _update_qa_panel(self):
        """Update the QA status panel."""
        if self.current_qa.overall_pass:
            self.qa_status_label.setText("QA Status: PASS")
            self.qa_status_label.setStyleSheet("color: green; font-weight: bold;")
            details = "All QA checks passed.\n\n"
            details += f"• Admissibility: PASS (A≥0, A+B≤1)\n"
            details += f"• Boundary: PASS (F(0)=0, F(1)=1)\n"
            details += f"• Range: PASS (0≤F≤1 on grid)\n"
            details += f"• Monotonicity: PASS (no decreases)\n"
            details += f"• Period Test: PASS (min monthly ≥ 0)\n"
            details += f"• Reconciliation: PASS (diff={self.current_qa.reconciliation_diff:.2e})"
        else:
            self.qa_status_label.setText("QA Status: FAIL")
            self.qa_status_label.setStyleSheet("color: red; font-weight: bold;")
            details = "QA CHECKS FAILED:\n\n"
            for error in self.current_qa.errors:
                details += f"• {error}\n"
            details += "\n⚠ Export requires logged override reason."
        
        self.qa_details.setText(details)
    
    def _update_milestone_table(self):
        """Update the milestone summary table."""
        milestones = get_milestone_summary(self.A, self.B)
        self.milestone_table.setRowCount(len(milestones))
        
        for row, (t, pct) in enumerate(sorted(milestones.items())):
            self.milestone_table.setItem(row, 0, QTableWidgetItem(f"{t:.2f}"))
            self.milestone_table.setItem(row, 1, QTableWidgetItem(f"{pct:.2f}%"))
    
    def _update_data_grid(self):
        """Update the data grid table."""
        self.data_grid.setRowCount(len(self.current_data['periods']))
        
        for row, (period, monthly, cumulative, cum_pct) in enumerate(zip(
                self.current_data['periods'],
                self.current_data['monthly_amounts'],
                self.current_data['cumulative_amounts'],
                self.current_data['cumulative_percentages'])):
            self.data_grid.setItem(row, 0, QTableWidgetItem(str(period)))
            self.data_grid.setItem(row, 1, QTableWidgetItem(f"${monthly:,.2f}"))
            self.data_grid.setItem(row, 2, QTableWidgetItem(f"${cumulative:,.2f}"))
            self.data_grid.setItem(row, 3, QTableWidgetItem(f"{cum_pct:.4f}%"))
    
    def _export_excel(self):
        """Export data to Excel."""
        if not self.current_qa.overall_pass:
            # Require override reason
            reason, ok = QInputDialog.getText(
                self, "Override Required", 
                "QA checks failed. Enter override reason to proceed with export:",
                QLineEdit.Normal, ""
            )
            if not ok or not reason:
                return
            
            self.audit_trail.log_change(
                "export_override", None, reason,
                "User override for failed QA"
            )
        
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Export to Excel", "", "Excel Files (*.xlsx)"
        )
        
        if filepath:
            inputs = {
                'budget': self.budget,
                'n_months': self.n_months,
                'start_date': self.start_date.toString("yyyy-MM-dd"),
                'A': self.A,
                'B': self.B,
                'scale_A': self.scale_A,
                'scale_B': self.scale_B
            }
            
            export_to_excel(self.current_data, inputs, self.current_qa, 
                           self.audit_trail, filepath)
            self.statusBar.showMessage(f"Exported to {filepath}")
    
    def _export_pdf(self):
        """Export to PDF report."""
        # For Phase 1, we'll save the figure as PDF
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Export PDF Report", "", "PDF Files (*.pdf)"
        )
        
        if filepath:
            self.figure.savefig(filepath, bbox_inches='tight')
            self.statusBar.showMessage(f"PDF report saved to {filepath}")
            logger.info(f"Exported PDF to {filepath}")
    
    def _save_project(self):
        """Save project to JSON file."""
        filepath, _ = QFileDialog.getSaveFileName(
            self, "Save Project", "", "JSON Files (*.json)"
        )
        
        if filepath:
            project_data = {
                'version': '1.0.0',
                'timestamp': datetime.now().isoformat(),
                'inputs': {
                    'budget': self.budget,
                    'n_months': self.n_months,
                    'start_date': self.start_date.toString("yyyy-MM-dd"),
                    'scale_A': self.scale_A,
                    'scale_B': self.scale_B,
                    'A': self.A,
                    'B': self.B
                },
                'data': self.current_data,
                'audit_trail': self.audit_trail.get_export_record()
            }
            
            with open(filepath, 'w') as f:
                json.dump(project_data, f, indent=2)
            
            self.statusBar.showMessage(f"Project saved to {filepath}")
            logger.info(f"Saved project to {filepath}")
    
    def _load_project(self):
        """Load project from JSON file."""
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Load Project", "", "JSON Files (*.json)"
        )
        
        if filepath:
            try:
                with open(filepath, 'r') as f:
                    project_data = json.load(f)
                
                inputs = project_data.get('inputs', {})
                self.budget = inputs.get('budget', self.budget)
                self.n_months = inputs.get('n_months', self.n_months)
                self.scale_A = inputs.get('scale_A', self.scale_A)
                self.scale_B = inputs.get('scale_B', self.scale_B)
                self.A = inputs.get('A', self.A)
                self.B = inputs.get('B', self.B)
                
                # Update UI
                self.budget_input.setValue(self.budget)
                self.months_input.setValue(self.n_months)
                self.slider_A.setValue(int(self.scale_A))
                self.slider_B.setValue(int(self.scale_B))
                self.input_A.setValue(self.A)
                self.input_B.setValue(self.B)
                self._update_sum_display()
                
                self._recalculate()
                self.statusBar.showMessage(f"Project loaded from {filepath}")
                logger.info(f"Loaded project from {filepath}")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load project: {str(e)}")


# AUTOMATED TESTS (TC1-TC5 for Phase 1)
# ============================================================================

def run_tests():
    """Run automated acceptance tests TC1-TC5."""
    print("=" * 70)
    print("PARAMETRIC S-CURVE WORKBENCH - AUTOMATED TEST REPORT")
    print("=" * 70)
    
    all_passed = True
    results = []
    
    # TC1: JSC preset verification
    print("\n[TC1] Testing JSC preset (A=0.32, B=0.68, C=120M, N=24)")
    try:
        A, B, C, N = 0.32, 0.68, 120000000.0, 24
        data = generate_s_curve_data(C, N, A, B)
        
        # Check specific periods
        P1_cum = data['cumulative_percentages'][0]
        P6_cum = data['cumulative_percentages'][5]
        P12_cum = data['cumulative_percentages'][11]
        P18_cum = data['cumulative_percentages'][17]
        P24_cum = data['cumulative_percentages'][23]
        
        P12_amt = data['cumulative_amounts'][11]
        P18_amt = data['cumulative_amounts'][17]
        P24_amt = data['cumulative_amounts'][23]
        
        print(f"  Period 1 Cum%: {P1_cum:.2f}% (expected ≈0.56%)")
        print(f"  Period 6 Cum%: {P6_cum:.2f}% (expected ≈18.79%)")
        print(f"  Period 12 Cum%: {P12_cum:.4f}% (expected =60.0000%)")
        print(f"  Period 12 Amount: ${P12_amt:,.2f} (expected =$72,000,000.00)")
        print(f"  Period 18 Cum%: {P18_cum:.4f}% (expected ≈92.4609%)")
        print(f"  Period 18 Amount: ${P18_amt:,.2f} (expected =$110,953,125.00)")
        print(f"  Period 24 Cum%: {P24_cum:.2f}% (expected =100.00%)")
        print(f"  Period 24 Amount: ${P24_amt:,.2f} (expected =$120,000,000.00)")
        
        # Verify tolerances
        tc1_pass = (
            abs(P1_cum - 0.56) < 0.1 and
            abs(P6_cum - 18.79) < 0.1 and
            abs(P12_cum - 60.0) < 0.001 and
            abs(P12_amt - 72000000.0) < 1.0 and
            abs(P18_cum - 92.4609) < 0.01 and
            abs(P18_amt - 110953125.0) < 1.0 and
            abs(P24_cum - 100.0) < 0.001 and
            abs(P24_amt - 120000000.0) < 1.0
        )
        
        if tc1_pass:
            print("  ✓ TC1 PASSED")
            results.append(("TC1", "PASS", ""))
        else:
            print("  ✗ TC1 FAILED")
            results.append(("TC1", "FAIL", "Values outside tolerance"))
            all_passed = False
            
    except Exception as e:
        print(f"  ✗ TC1 FAILED: {str(e)}")
        results.append(("TC1", "FAIL", str(e)))
        all_passed = False
    
    # TC1b: Milestone summary verification
    print("\n[TC1b] Testing milestone summary percentages")
    try:
        milestones = get_milestone_summary(0.32, 0.68)
        expected = {
            0.10: 3.19,
            0.25: 18.79,
            0.50: 60.00,
            0.65: 82.28,
            0.75: 92.46,
            0.90: 99.40
        }
        
        tc1b_pass = True
        for t, exp_val in expected.items():
            actual = milestones[t]
            diff = abs(actual - exp_val)
            status = "✓" if diff < 0.01 else "✗"
            print(f"  T={t:.2f}: {actual:.2f}% (expected {exp_val:.2f}%) {status}")
            if diff >= 0.01:
                tc1b_pass = False
        
        if tc1b_pass:
            print("  ✓ TC1b PASSED")
            results.append(("TC1b", "PASS", ""))
        else:
            print("  ✗ TC1b FAILED")
            results.append(("TC1b", "FAIL", "Milestone values outside ±0.01 tolerance"))
            all_passed = False
            
    except Exception as e:
        print(f"  ✗ TC1b FAILED: {str(e)}")
        results.append(("TC1b", "FAIL", str(e)))
        all_passed = False
    
    # TC2: Negative A value (QA FAIL)
    print("\n[TC2] Testing QA failure detection (A=-0.06, B=0.68)")
    try:
        A, B, C, N = -0.06, 0.68, 120000000.0, 24
        data = generate_s_curve_data(C, N, A, B)
        qa = run_qa_checks(C, N, A, B, data)
        
        print(f"  Admissibility PASS: {qa.admissibility_pass}")
        print(f"  Errors: {qa.errors}")
        
        # Check for negative monthly in period 1
        P1_monthly = data['monthly_amounts'][0]
        print(f"  Period 1 Monthly: ${P1_monthly:,.2f} (expected ≈-$58,840)")
        
        tc2_pass = (
            not qa.admissibility_pass and
            P1_monthly < 0 and
            abs(P1_monthly - (-58840)) < 1000
        )
        
        if tc2_pass:
            print("  ✓ TC2 PASSED (correctly detected QA failure)")
            results.append(("TC2", "PASS", ""))
        else:
            print("  ✗ TC2 FAILED")
            results.append(("TC2", "FAIL", "Did not detect QA failure correctly"))
            all_passed = False
            
    except Exception as e:
        print(f"  ✗ TC2 FAILED: {str(e)}")
        results.append(("TC2", "FAIL", str(e)))
        all_passed = False
    
    # TC3: Boundary case (A=1.00, B=0.00)
    print("\n[TC3] Testing boundary case (A=1.00, B=0.00)")
    try:
        A, B, C, N = 1.00, 0.00, 120000000.0, 24
        data = generate_s_curve_data(C, N, A, B)
        qa = run_qa_checks(C, N, A, B, data)
        
        print(f"  A+B = {A+B} (should be 1.0)")
        print(f"  Overall QA PASS: {qa.overall_pass}")
        print(f"  Monotonicity PASS: {qa.monotonicity_pass}")
        
        # Verify monotonicity
        is_monotonic = all(
            data['cumulative_percentages'][i] <= data['cumulative_percentages'][i+1] + 1e-9
            for i in range(len(data['cumulative_percentages'])-1)
        )
        
        tc3_pass = qa.overall_pass and is_monotonic
        
        if tc3_pass:
            print("  ✓ TC3 PASSED")
            results.append(("TC3", "PASS", ""))
        else:
            print("  ✗ TC3 FAILED")
            results.append(("TC3", "FAIL", "Boundary case failed"))
            all_passed = False
            
    except Exception as e:
        print(f"  ✗ TC3 FAILED: {str(e)}")
        results.append(("TC3", "FAIL", str(e)))
        all_passed = False
    
    # TC4: Preset reconciliation test
    print("\n[TC4] Testing preset reconciliation (Σ Monthly = C)")
    try:
        tc4_pass = True
        for preset_name, preset_data in PRESETS.items():
            A, B = preset_data['A'], preset_data['B']
            data = generate_s_curve_data(120000000.0, 24, A, B)
            total = sum(data['monthly_amounts'])
            diff = abs(total - 120000000.0)
            
            status = "✓" if diff < 0.01 else "✗"
            print(f"  {preset_name}: Σ=${total:,.2f}, diff=${diff:.2f} {status}")
            
            if diff >= 0.01:
                tc4_pass = False
        
        if tc4_pass:
            print("  ✓ TC4 PASSED")
            results.append(("TC4", "PASS", ""))
        else:
            print("  ✗ TC4 FAILED")
            results.append(("TC4", "FAIL", "Reconciliation failed for some presets"))
            all_passed = False
            
    except Exception as e:
        print(f"  ✗ TC4 FAILED: {str(e)}")
        results.append(("TC4", "FAIL", str(e)))
        all_passed = False
    
    # TC5: Export test
    print("\n[TC5] Testing Excel export")
    try:
        # Generate test data
        data = generate_s_curve_data(120000000.0, 24, 0.32, 0.68)
        qa = run_qa_checks(120000000.0, 24, 0.32, 0.68, data)
        audit = AuditTrail("test_audit.json")
        
        inputs = {
            'budget': 120000000.0,
            'n_months': 24,
            'start_date': '2025-01-01',
            'A': 0.32,
            'B': 0.68,
            'scale_A': 66.0,
            'scale_B': 84.0
        }
        
        test_file = "test_export.xlsx"
        export_to_excel(data, inputs, qa, audit, test_file)
        
        # Verify file exists and can be opened
        import os
        if os.path.exists(test_file):
            # Try to read it back
            wb = openpyxl.load_workbook(test_file)
            sheet_names = wb.sheetnames
            print(f"  Created file: {test_file}")
            print(f"  Sheets: {sheet_names}")
            
            # Verify key sheets exist
            required_sheets = ['Inputs', 'Monthly', 'Cumulative', 'QA_Log']
            missing = [s for s in required_sheets if s not in sheet_names]
            
            if not missing:
                print("  ✓ TC5 PASSED")
                results.append(("TC5", "PASS", ""))
            else:
                print(f"  ✗ TC5 FAILED: Missing sheets: {missing}")
                results.append(("TC5", "FAIL", f"Missing sheets: {missing}"))
                all_passed = False
        else:
            print("  ✗ TC5 FAILED: File not created")
            results.append(("TC5", "FAIL", "File not created"))
            all_passed = False
            
    except Exception as e:
        print(f"  ✗ TC5 FAILED: {str(e)}")
        results.append(("TC5", "FAIL", str(e)))
        all_passed = False
    
    # Summary
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    for test_id, status, detail in results:
        symbol = "✓" if status == "PASS" else "✗"
        print(f"{symbol} {test_id}: {status} {detail}")
    
    print("\n" + "=" * 70)
    if all_passed:
        print("ALL TESTS PASSED ✓")
    else:
        print("SOME TESTS FAILED ✗")
    print("=" * 70)
    
    return all_passed


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point for the application."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Parametric S-Curve Workbench")
    parser.add_argument('--test', action='store_true', help='Run automated tests')
    args = parser.parse_args()
    
    if args.test:
        success = run_tests()
        sys.exit(0 if success else 1)
    else:
        app = QApplication(sys.argv)
        app.setStyle('Fusion')
        
        window = SCurveWorkbench()
        window.show()
        
        sys.exit(app.exec())


if __name__ == "__main__":
    main()
