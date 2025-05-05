# Excel MCP Server

An MCP server that provides comprehensive Excel file management and data analysis capabilities.

## Features

- **Excel File Operations**
  - Read multiple Excel formats (XLSX, XLS, CSV, TSV, JSON)
  - Write and update Excel files
  - Get file information and sheet names

- **Data Analysis**
  - Summary statistics and descriptive analysis
  - Data quality assessment
  - Pivot tables
  - Filtering and querying data

- **Visualization**
  - Generate charts and plots from Excel data
  - Create data previews
  - Export visualizations as images

## Installation

1. Create a new Python environment (recommended):

```bash
# Using uv (recommended)
uv init excel-mcp-server
cd excel-mcp-server
uv venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Or using pip
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

2. Install dependencies:

```bash
# Using uv
uv pip install -e .
```

## Integration with Claude Desktop

1. Install [Claude Desktop](https://claude.ai/download)
2. Open Settings and go to the Developer tab
3. Edit `claude_desktop_config.json`:

```json
{
  "mcpServers": {
      "command": "uvx",
      "args": [
        "mcp-excel-server"
      ],
      "env": {
        "PYTHONPATH": "/path/to/your/python"
      }
  }
}
```

## Available Tools

### File Reading
- `read_excel`: Read Excel files
- `get_excel_info`: Get file details
- `get_sheet_names`: List worksheet names

### Data Analysis
- `analyze_excel`: Perform statistical analysis
- `filter_excel`: Filter data by conditions
- `pivot_table`: Create pivot tables
- `data_summary`: Generate comprehensive data summary

### Data Visualization
- `export_chart`: Generate charts
  - Supports line charts, bar charts, scatter plots, histograms

### File Operations
- `write_excel`: Write new Excel files
- `update_excel`: Update existing Excel files

## Available Resources

- `excel://{file_path}`: Get file content
- `excel://{file_path}/info`: Get file structure information
- `excel://{file_path}/preview`: Generate data preview image

## Prompt Templates

- `analyze_excel_data`: Guided template for Excel data analysis
- `create_chart`: Help create data visualizations
- `data_cleaning`: Assist with data cleaning

## Usage Examples

- "Analyze my sales_data.xlsx file"
- "Create a bar chart for product_sales.csv"
- "Filter employees over 30 in employees.xlsx"
- "Generate a pivot table of department sales"

## Security Considerations

- Read files only from specified paths
- Limit file size
- Prevent accidental file overwriting
- Strictly control data transformation operations

## Dependencies

- pandas
- numpy
- matplotlib
- seaborn

## License

MIT License