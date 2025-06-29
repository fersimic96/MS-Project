# MS-Project Converter

A Python tool for converting Microsoft Project (.mpp) files to Excel (.xlsx) format with complete task hierarchy, resources, and dependencies.

## Features

- Converts .mpp files to structured Excel workbooks
- Preserves task hierarchy and WBS structure
- Exports task dependencies and predecessor relationships
- Includes resource information and assignments
- Maintains critical path and milestone information
- Supports both command-line and programmatic usage

## Installation

1. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install required dependencies:
```bash
pip install jpype1 mpxj pandas openpyxl
```

## Directory Structure

```
MS-Project/
├── data/                    # Input .mpp files
├── output/                  # Generated Excel files
├── mpp_to_xlsx.py          # CLI conversion tool (flexible paths)
├── ms_project_converter.py  # Analysis tool (with visualization)
└── README.md               # This file
```

## Scripts Available

### 1. `mpp_to_xlsx.py` - Command-line Conversion Tool
- Flexible file paths via command-line arguments
- Optional verbose mode (-v flag)
- Production-ready CLI interface
- **⚠️ Warning**: Has duration conversion bug (factor ~24x)
- Best for: Quick conversions, automation, scripts

### 2. `ms_project_converter.py` - Analysis and Visualization Tool
- Shows detailed project summary and statistics
- Displays task hierarchy with indentation
- Shows project properties (title, manager, dates)
- **⚠️ Warning**: Has duration conversion bug (factor ~24x)
- Best for: Project analysis, understanding structure

### 3. `corrected_converter.py` - ✅ CORRECTED Conversion Tool
- **Fixes duration conversion bug** by comparing with native export
- Intelligent correction using native MS Project data
- Detailed correction reporting and statistics
- Produces accurate durations for project planning
- **Recommended for production use**

### 4. `gantt_visualizer.py` - Interactive Gantt Chart Generator
- Creates interactive HTML Gantt charts from Excel exports
- **✅ Auto-detects** corrected converter files and standard exports
- Color-coded by task status and critical path
- Milestone markers and resource charts
- Shows corrected duration data and source information
- Best for: Project visualization and presentations

## Usage

### ✅ Recommended: Corrected Conversion (corrected_converter.py)

Convert with accurate durations:
```bash
python corrected_converter.py data/your_project.mpp -o output/corrected_project.xlsx -v
```

**Note**: Requires native MS Project export (`gant FCC.xls`) in same directory for correction.

### Quick Conversion (mpp_to_xlsx.py) ⚠️

Convert a single .mpp file:
```bash
python mpp_to_xlsx.py data/your_project.mpp -o output/converted_project.xlsx
```

**Warning**: Durations will be ~24x smaller than actual values.

### Detailed Analysis (ms_project_converter.py) ⚠️

Edit the file paths in the script, then run:
```bash
python ms_project_converter.py
```

**Warning**: Durations will be ~24x smaller than actual values.

### Create Interactive Gantt Chart

From corrected converter file (recommended):
```bash
python gantt_visualizer.py output/proyecto_corregido.xlsx -o output/gantt_corregido.html -t "Project Title"
```

With resource charts:
```bash
python gantt_visualizer.py output/proyecto_corregido.xlsx --resources -o output/full_report.html
```

From any Excel export (auto-detects format):
```bash
python gantt_visualizer.py output/any_project.xlsx -o output/gantt_chart.html
```

### Python Script Usage

```python
from mpp_to_xlsx import read_ms_project, export_to_xlsx
import pandas as pd

# Read project file
tasks_data, project = read_ms_project('data/your_project.mpp')

# Convert to DataFrame
tasks_df = pd.DataFrame(tasks_data)

# Export to Excel
export_to_xlsx(tasks_df, project, 'output/converted_project.xlsx')
```

## Output Format

The generated Excel file contains:

### Tasks Sheet
- **ID**: Task ID number
- **WBS**: Work Breakdown Structure code
- **Name**: Task name
- **Duration**: Task duration
- **Start/Finish**: Planned dates
- **Percent Complete**: Progress percentage
- **Predecessors**: Dependencies (e.g., "3FS", "5SS+2d")
- **Resource Names**: Assigned resources
- **Cost**: Task cost
- **Work**: Work hours
- **Critical**: Critical path indicator
- **Milestone**: Milestone indicator
- **Summary**: Summary task indicator
- **Notes**: Task notes
- **Outline Level**: Hierarchy level

### Resources Sheet
- **ID**: Resource ID
- **Name**: Resource name
- **Type**: Resource type
- **Cost**: Total cost
- **Standard Rate**: Hourly/daily rate

## Requirements

- Python 3.8+
- Java Runtime Environment (JRE) 8+
- Dependencies:
  - jpype1
  - mpxj
  - pandas
  - openpyxl

## License

This project is open source and available under the MIT License.
