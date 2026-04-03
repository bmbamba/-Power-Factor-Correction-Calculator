# ⚡ Power Factor Correction Calculator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![PySide6](https://img.shields.io/badge/PySide6-6.0+-green.svg)](https://pypi.org/project/PySide6/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)](https://github.com/yourusername/power-factor-calculator)

A professional, feature-rich Power Factor Correction Calculator with comprehensive analysis tools for electrical engineers, facility managers, and energy consultants.

## 📋 Table of Contents
- [Features](#features)
- [Screenshots](#screenshots)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Guide](#usage-guide)
- [Technical Details](#technical-details)
- [Contributing](#contributing)
- [License](#license)

## ✨ Features

### Core Features
- **Basic Power Factor Calculator** - Calculate required capacitor bank size (kVAR), capacitance, and current ratings
- **Multiple Load Analysis** - Add and manage multiple electrical loads with weighted average calculations
- **Equipment Database** - Pre-loaded library with typical power factors for motors, pumps, compressors, etc.
- **Cost Analysis** - ROI calculations, demand charge savings, energy cost reduction, payback period
- **Harmonic Analysis** - Resonance risk assessment and harmonic filter recommendations
- **Visualization** - Interactive power triangle charts with before/after comparison
- **Professional Reports** - Detailed formatted reports with all calculations

### Technical Specifications
- Real-time power triangle visualization
- Support for 50/60 Hz systems
- Delta and Star connection options
- Customizable utility rates
- Export-ready reports

## 📸 Screenshots

### Main Interface
![Main Window](screenshots/main_window.png)

### Calculation Results
![Results](screenshots/results.png)

### Multiple Loads Management
![Loads](screenshots/loads.png)

### Cost Analysis
![Cost Analysis](screenshots/cost_analysis.png)

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- pip package manager

### Quick Install

```bash
# Clone the repository
git clone https://github.com/yourusername/power-factor-calculator.git
cd power-factor-calculator

# Install required dependencies
pip install PySide6

# Optional but recommended
pip install matplotlib pandas openpyxl

# Run the application
python power_factor_calculator.py
