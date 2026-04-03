"""
Enhanced Power Factor Correction Calculator
Copyright (c) 2026 Clan
MIT License

Professional power factor correction calculator with multiple loads, cost analysis, and reporting
"""

import sys
import math
import json
import os
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QGroupBox,
    QFrame, QMessageBox, QDoubleSpinBox, QComboBox, QTextEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QDialog,
    QDialogButtonBox, QFormLayout, QTabWidget, QFileDialog,
    QProgressBar, QSpinBox
)
from PySide6.QtCore import Qt, QTimer, QThread, Signal
from PySide6.QtGui import QFont, QIcon, QColor, QPalette

# Try to import optional dependencies
try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("Matplotlib not available. Charts disabled.")

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("Pandas not available. Excel export disabled.")

# ==================== LOAD ITEM CLASS ====================
class LoadItem:
    def __init__(self, name="", power=0, pf_current=0.85, pf_target=0.95):
        self.name = name
        self.power = power
        self.pf_current = pf_current
        self.pf_target = pf_target

# ==================== ADD LOAD DIALOG ====================
class AddLoadDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Load")
        self.setGeometry(400, 300, 400, 250)
        self.setStyleSheet("""
            QDialog { background-color: #1e1e2e; }
            QLabel { color: #e4e4e7; }
            QLineEdit, QDoubleSpinBox { 
                background-color: #1a1a26; 
                color: #e4e4e7; 
                border: 1px solid #2a2a38;
                border-radius: 6px;
                padding: 6px;
            }
            QPushButton { 
                background-color: #3b82f6; 
                color: white; 
                border: none; 
                padding: 8px 16px;
                border-radius: 6px;
            }
            QPushButton:hover { background-color: #2563eb; }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("e.g., Motor 1, Pump A")
        form_layout.addRow("Load Name:", self.name_edit)
        
        self.power_spin = QDoubleSpinBox()
        self.power_spin.setRange(0, 10000)
        self.power_spin.setValue(50)
        self.power_spin.setSuffix(" kW")
        form_layout.addRow("Power:", self.power_spin)
        
        self.pf_current_spin = QDoubleSpinBox()
        self.pf_current_spin.setRange(0.5, 0.99)
        self.pf_current_spin.setValue(0.85)
        self.pf_current_spin.setSingleStep(0.01)
        self.pf_current_spin.setDecimals(2)
        form_layout.addRow("Current PF:", self.pf_current_spin)
        
        self.pf_target_spin = QDoubleSpinBox()
        self.pf_target_spin.setRange(0.8, 0.99)
        self.pf_target_spin.setValue(0.95)
        self.pf_target_spin.setSingleStep(0.01)
        self.pf_target_spin.setDecimals(2)
        form_layout.addRow("Target PF:", self.pf_target_spin)
        
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
    
    def get_load(self):
        return LoadItem(
            name=self.name_edit.text() or f"Load {datetime.now().strftime('%H%M%S')}",
            power=self.power_spin.value(),
            pf_current=self.pf_current_spin.value(),
            pf_target=self.pf_target_spin.value()
        )

# ==================== EQUIPMENT DATABASE ====================
class EquipmentDatabase:
    def __init__(self):
        self.equipment = {
            "Induction Motors": [
                {"name": "10 HP Motor", "power": 7.46, "typical_pf": 0.82},
                {"name": "25 HP Motor", "power": 18.65, "typical_pf": 0.84},
                {"name": "50 HP Motor", "power": 37.30, "typical_pf": 0.86},
                {"name": "75 HP Motor", "power": 55.95, "typical_pf": 0.87},
                {"name": "100 HP Motor", "power": 74.60, "typical_pf": 0.88},
                {"name": "150 HP Motor", "power": 111.90, "typical_pf": 0.89},
                {"name": "200 HP Motor", "power": 149.20, "typical_pf": 0.90},
                {"name": "300 HP Motor", "power": 223.80, "typical_pf": 0.91},
            ],
            "Pumps": [
                {"name": "Small Centrifugal Pump", "power": 15, "typical_pf": 0.80},
                {"name": "Medium Centrifugal Pump", "power": 37, "typical_pf": 0.83},
                {"name": "Large Centrifugal Pump", "power": 75, "typical_pf": 0.85},
                {"name": "Slurry Pump", "power": 55, "typical_pf": 0.82},
            ],
            "Compressors": [
                {"name": "Small Screw Compressor", "power": 37, "typical_pf": 0.85},
                {"name": "Medium Screw Compressor", "power": 75, "typical_pf": 0.87},
                {"name": "Large Screw Compressor", "power": 150, "typical_pf": 0.89},
                {"name": "Centrifugal Compressor", "power": 300, "typical_pf": 0.90},
            ],
            "Lighting": [
                {"name": "LED Lighting (100 fixtures)", "power": 10, "typical_pf": 0.95},
                {"name": "Fluorescent Lighting", "power": 20, "typical_pf": 0.55},
                {"name": "HID Lighting", "power": 30, "typical_pf": 0.70},
            ],
            "HVAC": [
                {"name": "Small HVAC Unit", "power": 25, "typical_pf": 0.85},
                {"name": "Medium HVAC Unit", "power": 50, "typical_pf": 0.87},
                {"name": "Large HVAC Unit", "power": 100, "typical_pf": 0.88},
            ]
        }
    
    def get_categories(self):
        return list(self.equipment.keys())
    
    def get_equipment(self, category):
        return self.equipment.get(category, [])
    
    def get_typical_pf(self, category, equipment_name):
        for item in self.equipment.get(category, []):
            if item["name"] == equipment_name:
                return item["typical_pf"]
        return 0.85

# ==================== COST ESTIMATOR ====================
class CostEstimator:
    def __init__(self):
        self.capacitor_cost_per_kvar = 18  # $ per kVAR
        self.installation_cost_per_kvar = 12  # $ per kVAR
        self.base_installation = 500  # $ base
        self.controller_cost = 800  # $ for automatic systems
        self.annual_maintenance = 100  # $ per year
        
        # Utility rates
        self.demand_rate = 12.50  # $/kW
        self.energy_rate = 0.12  # $/kWh
        self.pf_penalty_threshold = 0.85
        self.pf_penalty_rate = 1.5  # multiplier for low PF
    
    def calculate_costs(self, Qc, is_automatic):
        capacitor_cost = Qc * self.capacitor_cost_per_kvar
        installation_cost = self.base_installation + (Qc * self.installation_cost_per_kvar)
        
        if is_automatic or Qc > 50:
            controller_cost = self.controller_cost
        else:
            controller_cost = 0
        
        total_initial = capacitor_cost + installation_cost + controller_cost
        
        return {
            "capacitor_cost": capacitor_cost,
            "installation_cost": installation_cost,
            "controller_cost": controller_cost,
            "total_initial": total_initial,
            "annual_maintenance": self.annual_maintenance
        }
    
    def calculate_savings(self, P, pf_current, pf_target, Qc, operating_hours=730):
        """Calculate monthly savings (operating_hours = hours per month)"""
        # Demand charge savings (kW reduction)
        demand_reduction = P * (1 - pf_current / pf_target)
        demand_savings = demand_reduction * self.demand_rate
        
        # Penalty avoidance
        if pf_current < self.pf_penalty_threshold:
            penalty_savings = P * self.pf_penalty_rate
        else:
            penalty_savings = 0
        
        # Energy savings from reduced losses (estimated 2% loss reduction)
        loss_reduction = Qc * 0.02
        energy_savings = loss_reduction * operating_hours * self.energy_rate
        
        total_monthly = demand_savings + penalty_savings + energy_savings
        
        return {
            "demand_savings": demand_savings,
            "penalty_savings": penalty_savings,
            "energy_savings": energy_savings,
            "total_monthly": total_monthly,
            "total_annual": total_monthly * 12
        }
    
    def calculate_roi(self, total_cost, annual_savings):
        if annual_savings > 0:
            payback_years = total_cost / annual_savings
            payback_months = payback_years * 12
            return {
                "years": payback_years,
                "months": payback_months,
                "roi_percentage": (annual_savings / total_cost) * 100
            }
        return {"years": 0, "months": 0, "roi_percentage": 0}

# ==================== HARMONIC ANALYZER ====================
class HarmonicAnalyzer:
    def __init__(self):
        self.harmonic_sources = {
            "VFDs": {"typical_thd": 0.35, "recommendation": "Use 5.67% detuned reactors"},
            "UPS": {"typical_thd": 0.15, "recommendation": "Use 7% detuned reactors"},
            "Welding": {"typical_thd": 0.40, "recommendation": "Use active harmonic filter"},
            "Motors": {"typical_thd": 0.05, "recommendation": "Standard capacitor bank OK"},
            "Lighting": {"typical_thd": 0.20, "recommendation": "Use 7% detuned reactors"},
        }
    
    def analyze(self, Qc, P, system_impedance=0.05):
        """Analyze harmonic resonance risk"""
        # Calculate resonance frequency
        Qc_kvar = Qc
        S_sc = P / system_impedance  # Short circuit power (approx)
        
        resonance_harmonic = math.sqrt(S_sc / Qc_kvar) if Qc_kvar > 0 else 0
        
        results = {
            "resonance_harmonic": resonance_harmonic,
            "risk_level": "Low",
            "recommendations": []
        }
        
        if 4 < resonance_harmonic < 7:
            results["risk_level"] = "High"
            results["recommendations"].append("⚠️ High risk of 5th harmonic resonance!")
            results["recommendations"].append("Use 5.67% detuned reactors")
        elif 3 < resonance_harmonic <= 4:
            results["risk_level"] = "Medium"
            results["recommendations"].append("Potential 3rd/4th harmonic issues")
            results["recommendations"].append("Consider 7% detuned reactors")
        elif resonance_harmonic >= 7:
            results["risk_level"] = "Low"
            results["recommendations"].append("Standard capacitor bank acceptable")
        
        return results

# ==================== CHART WIDGET ====================
class PowerTriangleChart(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        if MATPLOTLIB_AVAILABLE:
            self.figure = Figure(figsize=(6, 4), dpi=100)
            self.canvas = FigureCanvas(self.figure)
            layout.addWidget(self.canvas)
        else:
            self.label = QLabel("Matplotlib not installed. Install with: pip install matplotlib")
            self.label.setStyleSheet("color: #ef4444; padding: 20px;")
            self.label.setAlignment(Qt.AlignCenter)
            layout.addWidget(self.label)
        
        self.setLayout(layout)
    
    def update_chart(self, P, Q1, Q2, Qc):
        if not MATPLOTLIB_AVAILABLE:
            return
        
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        # Scale for better visualization
        max_val = max(P, Q1, Q2, Qc) * 1.2
        
        # Draw vectors
        ax.arrow(0, 0, P, Q1, head_width=max_val*0.05, head_length=max_val*0.05, 
                fc='red', ec='red', linewidth=2, label=f'Before (PF={Q1/P:.2f})')
        ax.arrow(0, 0, P, Q2, head_width=max_val*0.05, head_length=max_val*0.05, 
                fc='green', ec='green', linewidth=2, label=f'After (PF={Q2/P:.2f})')
        ax.arrow(0, Q2, 0, Qc, head_width=max_val*0.05, head_length=max_val*0.05, 
                fc='blue', ec='blue', linewidth=2, label=f'Correction ({Qc:.1f} kVAR)')
        
        ax.set_xlabel('Active Power (kW)')
        ax.set_ylabel('Reactive Power (kVAR)')
        ax.set_title('Power Triangle Visualization')
        ax.legend(loc='upper left')
        ax.grid(True, alpha=0.3)
        ax.set_xlim(0, max_val)
        ax.set_ylim(0, max_val)
        ax.axhline(y=0, color='black', linewidth=0.5)
        ax.axvline(x=0, color='black', linewidth=0.5)
        
        self.canvas.draw()

# ==================== PF SIMULATOR THREAD ====================
class PFSimulatorThread(QThread):
    pf_updated = Signal(float)
    finished = Signal()
    
    def __init__(self, target_pf=0.95, initial_pf=0.65):
        super().__init__()
        self.target_pf = target_pf
        self.current_pf = initial_pf
        self.running = True
    
    def run(self):
        step = 0.01
        while self.running and self.current_pf < self.target_pf:
            self.current_pf = min(self.current_pf + step, self.target_pf)
            self.pf_updated.emit(self.current_pf)
            self.msleep(100)
        self.finished.emit()
    
    def stop(self):
        self.running = False

# ==================== MAIN CALCULATOR WINDOW ====================
class PowerFactorCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.loads = []
        self.equipment_db = EquipmentDatabase()
        self.cost_estimator = CostEstimator()
        self.harmonic_analyzer = HarmonicAnalyzer()
        self.simulator_thread = None
        
        self.setup_ui()
        self.setWindowTitle("⚡ Power Factor Correction Calculator")
        self.setGeometry(100, 60, 1200, 800)
        
        self.setStyleSheet("""
            QMainWindow { background-color: #0f0f17; }
            QLabel { color: #e4e4e7; }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #2a2a38;
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 12px;
                color: #60a5fa;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 8px; }
            QLineEdit, QDoubleSpinBox, QComboBox, QSpinBox {
                background-color: #1a1a26;
                color: #e4e4e7;
                border: 1px solid #2a2a38;
                border-radius: 6px;
                padding: 8px;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #2563eb; }
            QPushButton[reset="true"] { background-color: #6b7280; }
            QPushButton[reset="true"]:hover { background-color: #4b5563; }
            QTableWidget {
                background-color: #1a1a26;
                color: #e4e4e7;
                border: 1px solid #2a2a38;
                border-radius: 6px;
                gridline-color: #2a2a38;
            }
            QTableWidget::item { padding: 5px; }
            QHeaderView::section {
                background-color: #2d2d3a;
                color: #60a5fa;
                padding: 8px;
                border: none;
            }
            QProgressBar {
                border: 1px solid #2a2a38;
                border-radius: 5px;
                text-align: center;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #22c55e;
                border-radius: 4px;
            }
            QTextEdit {
                background-color: #111118;
                color: #e4e4e7;
                border: 1px solid #2a2a38;
                border-radius: 6px;
                font-family: monospace;
            }
        """)
    
    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        title = QLabel("⚡ Power Factor Correction Calculator")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        title.setStyleSheet("color: #60a5fa;")
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)
        
        # Tabs
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabWidget::pane { background-color: #1a1a26; border-radius: 8px; }
            QTabBar::tab {
                background-color: #2d2d3a;
                color: #9ca3af;
                padding: 8px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected { background-color: #3b82f6; color: white; }
        """)
        
        # Tab 1: Basic Calculator
        basic_tab = self.create_basic_tab()
        tabs.addTab(basic_tab, "📊 Basic Calculator")
        
        # Tab 2: Multiple Loads
        loads_tab = self.create_loads_tab()
        tabs.addTab(loads_tab, "🏭 Multiple Loads")
        
        # Tab 3: Equipment Database
        equipment_tab = self.create_equipment_tab()
        tabs.addTab(equipment_tab, "📚 Equipment Database")
        
        # Tab 4: Cost Analysis
        cost_tab = self.create_cost_tab()
        tabs.addTab(cost_tab, "💰 Cost Analysis")
        
        # Tab 5: Charts & Visualization
        charts_tab = self.create_charts_tab()
        tabs.addTab(charts_tab, "📈 Charts")
        
        main_layout.addWidget(tabs)
        
        # Status Bar
        self.statusBar().setStyleSheet("background-color: #111118; color: #6b7280;")
        self.status_label = QLabel("Ready")
        self.statusBar().addWidget(self.status_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumWidth(200)
        self.statusBar().addPermanentWidget(self.progress_bar)
    
    def create_basic_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Input Group
        input_group = QGroupBox("Input Parameters")
        input_layout = QGridLayout()
        input_layout.setSpacing(12)
        
        # Power
        input_layout.addWidget(QLabel("Active Power (P):"), 0, 0)
        self.power_input = QDoubleSpinBox()
        self.power_input.setRange(0, 10000)
        self.power_input.setValue(100)
        self.power_input.setSuffix(" kW")
        self.power_input.setDecimals(1)
        input_layout.addWidget(self.power_input, 0, 1)
        
        # Voltage
        input_layout.addWidget(QLabel("Voltage:"), 0, 2)
        self.voltage_input = QDoubleSpinBox()
        self.voltage_input.setRange(100, 10000)
        self.voltage_input.setValue(415)
        self.voltage_input.setSuffix(" V")
        input_layout.addWidget(self.voltage_input, 0, 3)
        
        # Current PF
        input_layout.addWidget(QLabel("Current PF (cosφ₁):"), 1, 0)
        self.pf_current = QDoubleSpinBox()
        self.pf_current.setRange(0.5, 0.99)
        self.pf_current.setValue(0.75)
        self.pf_current.setSingleStep(0.01)
        self.pf_current.setDecimals(2)
        input_layout.addWidget(self.pf_current, 1, 1)
        
        # Target PF
        input_layout.addWidget(QLabel("Target PF (cosφ₂):"), 1, 2)
        self.pf_target = QDoubleSpinBox()
        self.pf_target.setRange(0.8, 0.99)
        self.pf_target.setValue(0.95)
        self.pf_target.setSingleStep(0.01)
        self.pf_target.setDecimals(2)
        input_layout.addWidget(self.pf_target, 1, 3)
        
        # Frequency
        input_layout.addWidget(QLabel("Frequency:"), 2, 0)
        self.frequency = QComboBox()
        self.frequency.addItems(["50 Hz", "60 Hz"])
        input_layout.addWidget(self.frequency, 2, 1)
        
        # Connection
        input_layout.addWidget(QLabel("Connection:"), 2, 2)
        self.connection = QComboBox()
        self.connection.addItems(["Delta (Δ)", "Star (Y)"])
        input_layout.addWidget(self.connection, 2, 3)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.calc_btn = QPushButton("📐 Calculate")
        self.calc_btn.clicked.connect(self.calculate_basic)
        self.reset_btn = QPushButton("🔄 Reset")
        self.reset_btn.clicked.connect(self.reset_basic)
        self.reset_btn.setProperty("reset", True)
        self.simulate_btn = QPushButton("🎬 Simulate PF Correction")
        self.simulate_btn.clicked.connect(self.start_simulation)
        
        btn_layout.addWidget(self.calc_btn)
        btn_layout.addWidget(self.reset_btn)
        btn_layout.addWidget(self.simulate_btn)
        layout.addLayout(btn_layout)
        
        # Results Group
        results_group = QGroupBox("Results")
        results_layout = QGridLayout()
        results_layout.setSpacing(10)
        
        self.reactive_power_label = QLabel()
        self.reactive_power_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #60a5fa;")
        results_layout.addWidget(QLabel("Current Reactive Power:"), 0, 0)
        results_layout.addWidget(self.reactive_power_label, 0, 1)
        
        self.reactive_target_label = QLabel()
        results_layout.addWidget(QLabel("Target Reactive Power:"), 0, 2)
        results_layout.addWidget(self.reactive_target_label, 0, 3)
        
        self.capacitor_kvar_label = QLabel()
        self.capacitor_kvar_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #22c55e;")
        results_layout.addWidget(QLabel("Required Capacitor Bank:"), 1, 0)
        results_layout.addWidget(self.capacitor_kvar_label, 1, 1)
        
        self.capacitor_current_label = QLabel()
        results_layout.addWidget(QLabel("Capacitor Current:"), 1, 2)
        results_layout.addWidget(self.capacitor_current_label, 1, 3)
        
        self.capacitance_label = QLabel()
        results_layout.addWidget(QLabel("Capacitance per Phase:"), 2, 0)
        results_layout.addWidget(self.capacitance_label, 2, 1)
        
        self.savings_label = QLabel()
        results_layout.addWidget(QLabel("Estimated Annual Savings:"), 2, 2)
        results_layout.addWidget(self.savings_label, 2, 3)
        
        results_group.setLayout(results_layout)
        layout.addWidget(results_group)
        
        # Detailed Report
        report_group = QGroupBox("Detailed Report")
        report_layout = QVBoxLayout()
        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        self.detail_text.setMaximumHeight(200)
        report_layout.addWidget(self.detail_text)
        report_group.setLayout(report_layout)
        layout.addWidget(report_group)
        
        # Simulation Group
        sim_group = QGroupBox("PF Correction Simulation")
        sim_layout = QVBoxLayout()
        self.sim_pf_label = QLabel("Current PF: 0.65")
        self.sim_pf_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        sim_layout.addWidget(self.sim_pf_label)
        
        self.sim_progress = QProgressBar()
        self.sim_progress.setRange(0, 100)
        sim_layout.addWidget(self.sim_progress)
        
        sim_group.setLayout(sim_layout)
        layout.addWidget(sim_group)
        
        return widget
    
    def create_loads_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Load table
        self.loads_table = QTableWidget()
        self.loads_table.setColumnCount(5)
        self.loads_table.setHorizontalHeaderLabels(["Load Name", "Power (kW)", "Current PF", "Target PF", "Actions"])
        self.loads_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.loads_table)
        
        # Buttons
        btn_layout = QHBoxLayout()
        add_load_btn = QPushButton("+ Add Load")
        add_load_btn.clicked.connect(self.add_load)
        remove_load_btn = QPushButton("- Remove Selected")
        remove_load_btn.clicked.connect(self.remove_load)
        calc_loads_btn = QPushButton("Calculate Total")
        calc_loads_btn.clicked.connect(self.calculate_loads)
        
        btn_layout.addWidget(add_load_btn)
        btn_layout.addWidget(remove_load_btn)
        btn_layout.addWidget(calc_loads_btn)
        layout.addLayout(btn_layout)
        
        # Total results
        self.loads_result_label = QLabel()
        self.loads_result_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #22c55e; padding: 10px; background-color: #111118; border-radius: 6px;")
        layout.addWidget(self.loads_result_label)
        
        return widget
    
    def create_equipment_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Category selector
        cat_layout = QHBoxLayout()
        cat_layout.addWidget(QLabel("Equipment Category:"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(self.equipment_db.get_categories())
        self.category_combo.currentTextChanged.connect(self.update_equipment_list)
        cat_layout.addWidget(self.category_combo)
        cat_layout.addStretch()
        layout.addLayout(cat_layout)
        
        # Equipment list
        self.equipment_list = QTableWidget()
        self.equipment_list.setColumnCount(3)
        self.equipment_list.setHorizontalHeaderLabels(["Equipment", "Power (kW)", "Typical PF"])
        self.equipment_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        layout.addWidget(self.equipment_list)
        
        # Use selected button
        use_btn = QPushButton("Use Selected Equipment")
        use_btn.clicked.connect(self.use_selected_equipment)
        layout.addWidget(use_btn)
        
        self.update_equipment_list()
        
        return widget
    
    def create_cost_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Rate settings
        rates_group = QGroupBox("Utility Rates")
        rates_layout = QGridLayout()
        
        rates_layout.addWidget(QLabel("Demand Rate ($/kW):"), 0, 0)
        self.demand_rate = QDoubleSpinBox()
        self.demand_rate.setRange(0, 100)
        self.demand_rate.setValue(12.50)
        self.demand_rate.setSuffix(" $/kW")
        rates_layout.addWidget(self.demand_rate, 0, 1)
        
        rates_layout.addWidget(QLabel("Energy Rate ($/kWh):"), 0, 2)
        self.energy_rate = QDoubleSpinBox()
        self.energy_rate.setRange(0, 1)
        self.energy_rate.setValue(0.12)
        self.energy_rate.setSingleStep(0.01)
        self.energy_rate.setDecimals(3)
        self.energy_rate.setSuffix(" $/kWh")
        rates_layout.addWidget(self.energy_rate, 0, 3)
        
        rates_layout.addWidget(QLabel("PF Penalty Threshold:"), 1, 0)
        self.pf_threshold = QDoubleSpinBox()
        self.pf_threshold.setRange(0.7, 0.95)
        self.pf_threshold.setValue(0.85)
        self.pf_threshold.setSingleStep(0.01)
        self.pf_threshold.setDecimals(2)
        rates_layout.addWidget(self.pf_threshold, 1, 1)
        
        rates_group.setLayout(rates_layout)
        layout.addWidget(rates_group)
        
        # Cost calculation
        self.cost_result_text = QTextEdit()
        self.cost_result_text.setReadOnly(True)
        layout.addWidget(self.cost_result_text)
        
        return widget
    
    def create_charts_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        self.chart_widget = PowerTriangleChart()
        layout.addWidget(self.chart_widget)
        
        return widget
    
    # ==================== CALCULATION METHODS ====================
    def calculate_basic(self):
        try:
            P = self.power_input.value()
            pf1 = self.pf_current.value()
            pf2 = self.pf_target.value()
            V = self.voltage_input.value()
            freq = 50 if self.frequency.currentText() == "50 Hz" else 60
            
            # Calculate angles
            theta1 = math.acos(pf1)
            theta2 = math.acos(pf2)
            
            # Reactive power
            Q1 = P * math.tan(theta1)
            Q2 = P * math.tan(theta2)
            Qc = Q1 - Q2
            
            if Qc < 0:
                Qc = 0
            
            # Calculate current and capacitance
            I_c = (Qc * 1000) / (math.sqrt(3) * V) if Qc > 0 else 0
            C = (Qc * 1000 * 1000) / (2 * math.pi * freq * V * V) if Qc > 0 else 0
            
            # Calculate savings
            savings = Qc * 0.02 * 8760 * 0.12  # Simplified annual savings
            
            # Update UI
            self.reactive_power_label.setText(f"{Q1:.2f} kVAR")
            self.reactive_target_label.setText(f"{Q2:.2f} kVAR")
            self.capacitor_kvar_label.setText(f"{Qc:.2f} kVAR")
            self.capacitor_current_label.setText(f"{I_c:.2f} A")
            self.capacitance_label.setText(f"{C:.2f} µF")
            self.savings_label.setText(f"${savings:.0f}/year")
            
            # Generate report
            report = self.generate_report(P, V, pf1, pf2, Q1, Q2, Qc, I_c, C, savings, freq)
            self.detail_text.setText(report)
            
            # Update chart
            if hasattr(self, 'chart_widget'):
                self.chart_widget.update_chart(P, Q1, Q2, Qc)
            
            # Harmonic analysis
            harmonic = self.harmonic_analyzer.analyze(Qc, P)
            if harmonic["risk_level"] == "High":
                self.status_label.setText(f"⚠️ {harmonic['risk_level']} harmonic risk detected!")
            else:
                self.status_label.setText(f"✓ Calculation complete. {harmonic['risk_level']} harmonic risk.")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Calculation error: {str(e)}")
    
    def calculate_loads(self):
        if not self.loads:
            QMessageBox.warning(self, "Warning", "No loads added. Click 'Add Load' to add equipment.")
            return
        
        total_power = sum(l.power for l in self.loads)
        # Weighted average of power factors
        total_kvar = sum(l.power * math.tan(math.acos(l.pf_current)) for l in self.loads)
        
        if total_power > 0:
            avg_pf = math.cos(math.atan(total_kvar / total_power))
        else:
            avg_pf = 0
        
        # Target PF (use the most common target or default)
        target_pf = max(set(l.pf_target for l in self.loads), key=list(l.pf_target for l in self.loads).count) if self.loads else 0.95
        
        theta1 = math.acos(avg_pf)
        theta2 = math.acos(target_pf)
        
        Q1 = total_power * math.tan(theta1)
        Q2 = total_power * math.tan(theta2)
        Qc = Q1 - Q2
        
        if Qc < 0:
            Qc = 0
        
        result_text = f"""
╔══════════════════════════════════════════════════════════════╗
║                    TOTAL LOAD CALCULATION                    ║
╚══════════════════════════════════════════════════════════════╝

Total Active Power: {total_power:.2f} kW
Average Power Factor: {avg_pf:.3f}
Target Power Factor: {target_pf:.3f}

Required Capacitor Bank: {Qc:.2f} kVAR

Load Breakdown:
"""
        for i, load in enumerate(self.loads, 1):
            result_text += f"\n{i}. {load.name}: {load.power:.1f} kW @ PF {load.pf_current:.2f}"
        
        self.loads_result_label.setText(f"Total Required: {Qc:.2f} kVAR | Total Power: {total_power:.2f} kW")
        
        # Update basic calculator values
        self.power_input.setValue(total_power)
        self.pf_current.setValue(avg_pf)
        self.pf_target.setValue(target_pf)
        
        # Calculate basic
        self.calculate_basic()
    
    def generate_report(self, P, V, pf1, pf2, Q1, Q2, Qc, I_c, C, savings, freq):
        # Get harmonic analysis
        harmonic = self.harmonic_analyzer.analyze(Qc, P)
        
        # Get cost analysis
        costs = self.cost_estimator.calculate_costs(Qc, Qc > 50)
        savings_calc = self.cost_estimator.calculate_savings(P, pf1, pf2, Qc)
        roi = self.cost_estimator.calculate_roi(costs["total_initial"], savings_calc["total_annual"])
        
        report = f"""
╔══════════════════════════════════════════════════════════════╗
║              POWER FACTOR CORRECTION REPORT                  ║
║                    Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
╚══════════════════════════════════════════════════════════════╝

INPUT PARAMETERS:
───────────────────────────────────────────────────────────────
Active Power (P)           : {P:.2f} kW
System Voltage (V)         : {V:.0f} V
Frequency                  : {freq} Hz
Connection Type            : {self.connection.currentText()}
Current Power Factor (PF₁) : {pf1:.3f}
Target Power Factor (PF₂)  : {pf2:.3f}

CALCULATED VALUES:
───────────────────────────────────────────────────────────────
Current Reactive Power (Q₁) : {Q1:.2f} kVAR
Target Reactive Power (Q₂)  : {Q2:.2f} kVAR
Required Capacitor Bank (Qc): {Qc:.2f} kVAR

CAPACITOR SPECIFICATIONS:
───────────────────────────────────────────────────────────────
Capacitor Current (I_c)     : {I_c:.2f} A
Capacitance per Phase (C)   : {C:.2f} µF
Recommended Rating          : {math.ceil(Qc / 5) * 5:.0f} kVAR (standard size)

HARMONIC ANALYSIS:
───────────────────────────────────────────────────────────────
Resonance Risk Level        : {harmonic["risk_level"]}
"""
        
        for rec in harmonic["recommendations"]:
            report += f"• {rec}\n"
        
        report += f"""
COST ANALYSIS:
───────────────────────────────────────────────────────────────
Capacitor Bank Cost         : ${costs["capacitor_cost"]:.0f}
Installation Cost           : ${costs["installation_cost"]:.0f}
Controller Cost (if needed) : ${costs["controller_cost"]:.0f}
Total Initial Investment    : ${costs["total_initial"]:.0f}

ANNUAL SAVINGS:
───────────────────────────────────────────────────────────────
Demand Charge Savings       : ${savings_calc["demand_savings"]:.0f}
Penalty Avoidance           : ${savings_calc["penalty_savings"]:.0f}
Energy Savings              : ${savings_calc["energy_savings"]:.0f}
Total Annual Savings        : ${savings_calc["total_annual"]:.0f}

RETURN ON INVESTMENT:
───────────────────────────────────────────────────────────────
Payback Period              : {roi["months"]:.1f} months
ROI Percentage              : {roi["roi_percentage"]:.1f}% per year

RECOMMENDATIONS:
───────────────────────────────────────────────────────────────
"""
        
        if Qc > 0:
            report += f"• Install {Qc:.2f} kVAR capacitor bank\n"
            report += f"• Use {math.ceil(Qc / 25)} steps for automatic control\n"
        else:
            report += "• No capacitor bank required - PF already at or above target\n"
        
        report += "• Verify capacitor switching device ratings\n"
        report += "• Consider automatic power factor controller for varying loads\n"
        
        return report
    
    # ==================== LOAD MANAGEMENT ====================
    def add_load(self):
        dialog = AddLoadDialog(self)
        if dialog.exec():
            load = dialog.get_load()
            self.loads.append(load)
            self.update_loads_table()
    
    def remove_load(self):
        current_row = self.loads_table.currentRow()
        if current_row >= 0:
            del self.loads[current_row]
            self.update_loads_table()
    
    def update_loads_table(self):
        self.loads_table.setRowCount(len(self.loads))
        for i, load in enumerate(self.loads):
            self.loads_table.setItem(i, 0, QTableWidgetItem(load.name))
            self.loads_table.setItem(i, 1, QTableWidgetItem(f"{load.power:.1f} kW"))
            self.loads_table.setItem(i, 2, QTableWidgetItem(f"{load.pf_current:.2f}"))
            self.loads_table.setItem(i, 3, QTableWidgetItem(f"{load.pf_target:.2f}"))
            
            # Remove button
            remove_btn = QPushButton("Remove")
            remove_btn.clicked.connect(lambda checked, row=i: self.remove_load_at(row))
            self.loads_table.setCellWidget(i, 4, remove_btn)
    
    def remove_load_at(self, row):
        del self.loads[row]
        self.update_loads_table()
    
    # ==================== EQUIPMENT DATABASE ====================
    def update_equipment_list(self):
        category = self.category_combo.currentText()
        equipment = self.equipment_db.get_equipment(category)
        
        self.equipment_list.setRowCount(len(equipment))
        for i, item in enumerate(equipment):
            self.equipment_list.setItem(i, 0, QTableWidgetItem(item["name"]))
            self.equipment_list.setItem(i, 1, QTableWidgetItem(f"{item['power']:.1f} kW"))
            self.equipment_list.setItem(i, 2, QTableWidgetItem(f"{item['typical_pf']:.2f}"))
    
    def use_selected_equipment(self):
        current_row = self.equipment_list.currentRow()
        if current_row >= 0:
            category = self.category_combo.currentText()
            equipment = self.equipment_db.get_equipment(category)[current_row]
            
            self.power_input.setValue(equipment["power"])
            self.pf_current.setValue(equipment["typical_pf"])
            self.status_label.setText(f"Selected: {equipment['name']}")
    
    # ==================== SIMULATION ====================
    def start_simulation(self):
        if self.simulator_thread and self.simulator_thread.isRunning():
            self.simulator_thread.stop()
        
        target_pf = self.pf_target.value()
        initial_pf = self.pf_current.value()
        
        self.simulator_thread = PFSimulatorThread(target_pf, initial_pf)
        self.simulator_thread.pf_updated.connect(self.update_simulation)
        self.simulator_thread.finished.connect(self.simulation_finished)
        self.simulator_thread.start()
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(int(initial_pf * 100), int(target_pf * 100))
        self.simulate_btn.setEnabled(False)
    
    def update_simulation(self, pf):
        self.sim_pf_label.setText(f"Current PF: {pf:.3f}")
        self.sim_progress.setValue(int(pf * 100))
        
        # Update basic calculator
        self.pf_current.setValue(pf)
        self.calculate_basic()
    
    def simulation_finished(self):
        self.progress_bar.setVisible(False)
        self.simulate_btn.setEnabled(True)
        self.status_label.setText("Simulation complete!")
    
    # ==================== RESET METHODS ====================
    def reset_basic(self):
        self.power_input.setValue(100)
        self.voltage_input.setValue(415)
        self.pf_current.setValue(0.75)
        self.pf_target.setValue(0.95)
        self.frequency.setCurrentIndex(0)
        self.connection.setCurrentIndex(0)
        
        self.reactive_power_label.setText("---")
        self.reactive_target_label.setText("---")
        self.capacitor_kvar_label.setText("---")
        self.capacitor_current_label.setText("---")
        self.capacitance_label.setText("---")
        self.savings_label.setText("---")
        self.detail_text.clear()
        self.status_label.setText("Reset to default values")

# ==================== MAIN ====================
def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = PowerFactorCalculator()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()