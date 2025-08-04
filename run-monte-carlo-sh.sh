#!/bin/bash

# Monte Carlo Simulation Setup and Run Script
# This script sets up a Python environment, installs dependencies, and runs the consumer harm analysis

set -e  # Exit on error

echo "=================================================="
echo "Consumer Harm Monte Carlo Simulation Setup"
echo "=================================================="
echo

# Color codes for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

# Check if Python 3 is installed
print_status "Checking Python installation..."
if ! command -v python3 &> /dev/null; then
    print_error "Python 3 is not installed. Please install Python 3.8 or higher."
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')
print_status "Python $PYTHON_VERSION detected"

# Create virtual environment
VENV_DIR="monte_carlo_env"
if [ -d "$VENV_DIR" ]; then
    print_warning "Virtual environment already exists. Removing old environment..."
    rm -rf "$VENV_DIR"
fi

print_status "Creating virtual environment..."
python3 -m venv "$VENV_DIR"

# Activate virtual environment
print_status "Activating virtual environment..."
source "$VENV_DIR/bin/activate"

# Upgrade pip
print_status "Upgrading pip..."
pip install --upgrade pip > /dev/null 2>&1

# Create requirements.txt
print_status "Creating requirements.txt..."
cat > requirements.txt << EOF
numpy==1.24.3
pandas==2.0.3
matplotlib==3.7.2
seaborn==0.12.2
scipy==1.11.1
plotly==5.15.0
kaleido==0.2.1
notebook==7.0.2
ipykernel==6.25.0
openpyxl==3.1.2
xlsxwriter==3.1.3
EOF

# Install dependencies
print_status "Installing dependencies..."
print_status "This may take a few minutes..."
pip install -r requirements.txt

# Create output directory
OUTPUT_DIR="monte_carlo_output"
if [ ! -d "$OUTPUT_DIR" ]; then
    print_status "Creating output directory..."
    mkdir -p "$OUTPUT_DIR"
fi

# Check if Python script exists
SCRIPT_NAME="consumer_harm_monte_carlo.py"
EXCEL_SCRIPT="export_to_excel.py"

if [ ! -f "$SCRIPT_NAME" ]; then
    print_error "Python script '$SCRIPT_NAME' not found!"
    print_status "Please ensure the script is in the current directory."
    exit 1
fi

# Run the simulation
print_status "Running Monte Carlo simulation..."
print_status "This will generate:"
echo "  - Statistical analysis of consumer harm"
echo "  - Visualizations (PNG and interactive HTML)"
echo "  - CSV data export"
echo "  - Summary report"
echo

cd "$OUTPUT_DIR"
python ../"$SCRIPT_NAME"

# Check if Excel export script exists and run it
if [ -f "../$EXCEL_SCRIPT" ]; then
    print_status "Generating Excel report..."
    python ../"$EXCEL_SCRIPT"
else
    print_warning "Excel export script not found. Skipping Excel generation."
fi

cd ..

# Check if outputs were generated
print_status "Checking output files..."
if [ -f "$OUTPUT_DIR/consumer_harm_analysis.png" ]; then
    print_status "✓ Static visualizations generated"
else
    print_warning "Static visualizations not found"
fi

if [ -f "$OUTPUT_DIR/consumer_harm_interactive.html" ]; then
    print_status "✓ Interactive dashboard generated"
else
    print_warning "Interactive dashboard not found"
fi

if [ -f "$OUTPUT_DIR/monte_carlo_results.csv" ]; then
    print_status "✓ Raw data exported"
else
    print_warning "Raw data export not found"
fi

if [ -f "$OUTPUT_DIR/simulation_summary.txt" ]; then
    print_status "✓ Summary report generated"
    echo
    print_status "Summary preview:"
    echo "----------------------------------------"
    head -n 20 "$OUTPUT_DIR/simulation_summary.txt"
    echo "----------------------------------------"
else
    print_warning "Summary report not found"
fi

if [ -f "$OUTPUT_DIR/Consumer_Harm_Analysis.xlsx" ]; then
    print_status "✓ Excel report generated"
    print_status "The Excel file contains 7 comprehensive sheets with charts and analysis"
else
    print_warning "Excel report not found"
fi

# Deactivate virtual environment
deactivate

echo
echo "=================================================="
print_status "Simulation complete!"
echo "=================================================="
echo
print_status "Output files are in the '$OUTPUT_DIR' directory:"
echo "  - consumer_harm_analysis.png: Statistical visualizations"
echo "  - consumer_harm_interactive.html: Interactive dashboard"
echo "  - monte_carlo_results.csv: Raw simulation data"
echo "  - monte_carlo_raw_data.csv: Full dataset for analysis"
echo "  - simulation_summary.txt: Statistical summary"
echo "  - Consumer_Harm_Analysis.xlsx: Comprehensive Excel report"
echo
print_status "The Excel file includes:"
echo "  • Executive Summary with key findings"
echo "  • Detailed simulation results (1,000 records)"
echo "  • Percentile analysis of harm distribution"
echo "  • Harm components breakdown"
echo "  • Scenario comparison with cost-benefit analysis"
echo "  • Professional charts and visualizations"
echo "  • Complete parameter documentation"
echo
print_status "To view the interactive dashboard, open:"
echo "  $OUTPUT_DIR/consumer_harm_interactive.html"
echo
print_status "To open the Excel report:"
echo "  $OUTPUT_DIR/Consumer_Harm_Analysis.xlsx"
echo
print_status "To run Jupyter notebook instead:"
echo "  1. source $VENV_DIR/bin/activate"
echo "  2. jupyter notebook consumer_harm_notebook.ipynb"
echo "  3. deactivate (when done)"
echo

# Optional: Open the files
read -p "Would you like to open the Excel report now? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    if command -v open &> /dev/null; then
        # macOS
        open "$OUTPUT_DIR/Consumer_Harm_Analysis.xlsx"
    elif command -v xdg-open &> /dev/null; then
        # Linux
        xdg-open "$OUTPUT_DIR/Consumer_Harm_Analysis.xlsx"
    elif command -v start &> /dev/null; then
        # Windows (WSL)
        start "$OUTPUT_DIR/Consumer_Harm_Analysis.xlsx"
    else
        print_warning "Could not automatically open the Excel file. Please open manually."
    fi
fi