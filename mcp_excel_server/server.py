import os
import io
import json
import pandas as pd
import numpy as np
from typing import Optional, Dict, List, Union, Tuple, Any
from dataclasses import dataclass
import base64
from datetime import datetime
from mcp.server.fastmcp import FastMCP, Context, Image

# Create the MCP server
mcp = FastMCP("Excel Data Manager")

# Helper functions
def _read_excel_file(file_path: str) -> Tuple[pd.DataFrame, str]:
    """
    Read an Excel file and return a DataFrame and the file extension.
    Supports .xlsx, .xls, .csv, and other formats pandas can read.
    """
    # Check if file exists
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    # Get file extension
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    # Read based on file extension
    if ext in ['.xlsx', '.xls', '.xlsm']:
        df = pd.read_excel(file_path)
    elif ext == '.csv':
        df = pd.read_csv(file_path)
    elif ext == '.tsv':
        df = pd.read_csv(file_path, sep='\t')
    elif ext == '.json':
        df = pd.read_json(file_path)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")
    
    return df, ext

def _get_dataframe_info(df: pd.DataFrame) -> Dict[str, Any]:
    """Generate summary information about a DataFrame."""
    # Basic info
    info = {
        "shape": df.shape,
        "columns": list(df.columns),
        "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
        "missing_values": df.isnull().sum().to_dict(),
        "total_memory_usage": df.memory_usage(deep=True).sum(),
    }
    
    # Sample data (first 5 rows)
    info["sample"] = df.head(5).to_dict(orient='records')
    
    # Numeric column stats
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        info["numeric_stats"] = {}
        for col in numeric_cols:
            info["numeric_stats"][col] = {
                "min": float(df[col].min()) if not pd.isna(df[col].min()) else None,
                "max": float(df[col].max()) if not pd.isna(df[col].max()) else None,
                "mean": float(df[col].mean()) if not pd.isna(df[col].mean()) else None,
                "median": float(df[col].median()) if not pd.isna(df[col].median()) else None,
                "std": float(df[col].std()) if not pd.isna(df[col].std()) else None
            }
    
    return info

# Resource Handlers

@mcp.resource("excel://{file_path}")
def get_excel_file(file_path: str) -> str:
    """
    Retrieve content of an Excel file as a formatted text representation.
    
    Args:
        file_path: Path to the Excel file to read
        
    Returns:
        String representation of the Excel data
    """
    df, _ = _read_excel_file(file_path)
    return df.to_string(index=False)

@mcp.resource("excel://{file_path}/info")
def get_excel_info(file_path: str) -> str:
    """
    Retrieve information about an Excel file including structure and stats.
    
    Args:
        file_path: Path to the Excel file to analyze
        
    Returns:
        JSON string with information about the Excel file
    """
    df, ext = _read_excel_file(file_path)
    info = _get_dataframe_info(df)
    info["file_path"] = file_path
    info["file_type"] = ext
    return json.dumps(info, indent=2, default=str)

@mcp.resource("excel://{file_path}/sheet_names")
def get_sheet_names(file_path: str) -> str:
    """
    Get the names of all sheets in an Excel workbook.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        JSON string with sheet names
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    if ext not in ['.xlsx', '.xls', '.xlsm']:
        return json.dumps({"error": "File is not an Excel workbook"})
    
    xls = pd.ExcelFile(file_path)
    return json.dumps({"sheet_names": xls.sheet_names})

@mcp.resource("excel://{file_path}/preview")
def get_excel_preview(file_path: str) -> Image:
    """
    Generate a visual preview of an Excel file.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Image of the data visualization
    """
    import matplotlib.pyplot as plt
    import seaborn as sns
    
    df, _ = _read_excel_file(file_path)
    
    # Create a styled preview
    plt.figure(figsize=(10, 6))
    
    # If DataFrame is small enough, show as a table
    if df.shape[0] <= 10 and df.shape[1] <= 10:
        plt.axis('tight')
        plt.axis('off')
        table = plt.table(cellText=df.values,
                          colLabels=df.columns,
                          cellLoc='center',
                          loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1.2, 1.2)
    else:
        # For larger DataFrames, show a heatmap of the first 10x10 section
        preview_df = df.iloc[:10, :10]
        sns.heatmap(preview_df.select_dtypes(include=['number']), 
                    cmap='viridis', 
                    annot=False,
                    linewidths=.5)
        plt.title(f"Preview of {os.path.basename(file_path)}")
    
    # Save to bytes buffer
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    
    # Convert to Image
    plt.close()
    return Image(data=buf.getvalue(), format="png")

# Tool Handlers

@mcp.tool()
def read_excel(file_path: str, sheet_name: Optional[str] = None, 
             nrows: Optional[int] = None, header: Optional[int] = 0) -> str:
    """
    Read an Excel file and return its contents as a string.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to read (only for .xlsx, .xls)
        nrows: Maximum number of rows to read
        header: Row to use as header (0-indexed)
        
    Returns:
        String representation of the Excel data
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    read_params = {"header": header}
    if nrows is not None:
        read_params["nrows"] = nrows
    
    if ext in ['.xlsx', '.xls', '.xlsm']:
        if sheet_name is not None:
            read_params["sheet_name"] = sheet_name
        df = pd.read_excel(file_path, **read_params)
    elif ext == '.csv':
        df = pd.read_csv(file_path, **read_params)
    elif ext == '.tsv':
        df = pd.read_csv(file_path, sep='\t', **read_params)
    elif ext == '.json':
        df = pd.read_json(file_path)
    else:
        return f"Unsupported file extension: {ext}"
    
    return df.to_string(index=False)

@mcp.tool()
def write_excel(file_path: str, data: str, sheet_name: Optional[str] = "Sheet1", 
              format: Optional[str] = "csv") -> str:
    """
    Write data to an Excel file.
    
    Args:
        file_path: Path to save the Excel file
        data: Data in CSV or JSON format
        sheet_name: Name of the sheet (for Excel files)
        format: Format of the input data ('csv' or 'json')
        
    Returns:
        Confirmation message
    """
    try:
        if format.lower() == 'csv':
            df = pd.read_csv(io.StringIO(data))
        elif format.lower() == 'json':
            df = pd.read_json(io.StringIO(data))
        else:
            return f"Unsupported data format: {format}"
        
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        if ext in ['.xlsx', '.xls', '.xlsm']:
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
        elif ext == '.csv':
            df.to_csv(file_path, index=False)
        elif ext == '.tsv':
            df.to_csv(file_path, sep='\t', index=False)
        elif ext == '.json':
            df.to_json(file_path, orient='records')
        else:
            return f"Unsupported output file extension: {ext}"
        
        return f"Successfully wrote data to {file_path}"
    except Exception as e:
        return f"Error writing data: {str(e)}"

@mcp.tool()
def update_excel(file_path: str, data: str, sheet_name: Optional[str] = "Sheet1",
               format: Optional[str] = "csv") -> str:
    """
    Update an existing Excel file with new data.
    
    Args:
        file_path: Path to the Excel file to update
        data: New data in CSV or JSON format
        sheet_name: Name of the sheet to update (for Excel files)
        format: Format of the input data ('csv' or 'json')
        
    Returns:
        Confirmation message
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            return f"File not found: {file_path}"
        
        # Load new data
        if format.lower() == 'csv':
            new_df = pd.read_csv(io.StringIO(data))
        elif format.lower() == 'json':
            new_df = pd.read_json(io.StringIO(data))
        else:
            return f"Unsupported data format: {format}"
        
        # Get file extension
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        # Read existing file
        if ext in ['.xlsx', '.xls', '.xlsm']:
            # For Excel files, we need to read all sheets
            excel_file = pd.ExcelFile(file_path)
            with pd.ExcelWriter(file_path) as writer:
                # Copy all existing sheets
                for sheet in excel_file.sheet_names:
                    if sheet != sheet_name:
                        df = pd.read_excel(excel_file, sheet_name=sheet)
                        df.to_excel(writer, sheet_name=sheet, index=False)
                
                # Write new data to specified sheet
                new_df.to_excel(writer, sheet_name=sheet_name, index=False)
        elif ext == '.csv':
            new_df.to_csv(file_path, index=False)
        elif ext == '.tsv':
            new_df.to_csv(file_path, sep='\t', index=False)
        elif ext == '.json':
            new_df.to_json(file_path, orient='records')
        else:
            return f"Unsupported file extension: {ext}"
        
        return f"Successfully updated {file_path}"
    except Exception as e:
        return f"Error updating file: {str(e)}"

@mcp.tool()
def analyze_excel(file_path: str, columns: Optional[str] = None, 
                sheet_name: Optional[str] = None) -> str:
    """
    Perform statistical analysis on Excel data.
    
    Args:
        file_path: Path to the Excel file
        columns: Comma-separated list of columns to analyze (analyzes all numeric columns if None)
        sheet_name: Name of the sheet to analyze (for Excel files)
        
    Returns:
        JSON string with statistical analysis
    """
    try:
        # Read file
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        read_params = {}
        if ext in ['.xlsx', '.xls', '.xlsm'] and sheet_name is not None:
            read_params["sheet_name"] = sheet_name
            
        if ext in ['.xlsx', '.xls', '.xlsm']:
            df = pd.read_excel(file_path, **read_params)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext == '.tsv':
            df = pd.read_csv(file_path, sep='\t')
        elif ext == '.json':
            df = pd.read_json(file_path)
        else:
            return f"Unsupported file extension: {ext}"
            
        # Filter columns if specified
        if columns:
            column_list = [c.strip() for c in columns.split(',')]
            df = df[column_list]
        
        # Select only numeric columns for analysis
        numeric_df = df.select_dtypes(include=['number'])
        
        if numeric_df.empty:
            return json.dumps({"error": "No numeric columns found for analysis"})
        
        # Perform analysis
        analysis = {
            "descriptive_stats": numeric_df.describe().to_dict(),
            "correlation": numeric_df.corr().to_dict(),
            "missing_values": numeric_df.isnull().sum().to_dict(),
            "unique_values": {col: int(numeric_df[col].nunique()) for col in numeric_df.columns}
        }
        
        return json.dumps(analysis, indent=2, default=str)
    except Exception as e:
        return json.dumps({"error": str(e)})

@mcp.tool()
def filter_excel(file_path: str, query: str, sheet_name: Optional[str] = None) -> str:
    """
    Filter Excel data using a pandas query string.
    
    Args:
        file_path: Path to the Excel file
        query: Pandas query string (e.g., "Age > 30 and Department == 'Sales'")
        sheet_name: Name of the sheet to filter (for Excel files)
        
    Returns:
        Filtered data as string
    """
    try:
        # Read file
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        read_params = {}
        if ext in ['.xlsx', '.xls', '.xlsm'] and sheet_name is not None:
            read_params["sheet_name"] = sheet_name
            
        if ext in ['.xlsx', '.xls', '.xlsm']:
            df = pd.read_excel(file_path, **read_params)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext == '.tsv':
            df = pd.read_csv(file_path, sep='\t')
        elif ext == '.json':
            df = pd.read_json(file_path)
        else:
            return f"Unsupported file extension: {ext}"
        
        # Apply filter
        filtered_df = df.query(query)
        
        # Return results
        if filtered_df.empty:
            return "No data matches the filter criteria."
        
        return filtered_df.to_string(index=False)
    except Exception as e:
        return f"Error filtering data: {str(e)}"

@mcp.tool()
def pivot_table(file_path: str, index: str, columns: Optional[str] = None, 
              values: str = None, aggfunc: str = "mean", 
              sheet_name: Optional[str] = None) -> str:
    """
    Create a pivot table from Excel data.
    
    Args:
        file_path: Path to the Excel file
        index: Column to use as the pivot table index
        columns: Optional column to use as the pivot table columns
        values: Column to use as the pivot table values
        aggfunc: Aggregation function ('mean', 'sum', 'count', etc.)
        sheet_name: Name of the sheet to pivot (for Excel files)
        
    Returns:
        Pivot table as string
    """
    try:
        # Read file
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        read_params = {}
        if ext in ['.xlsx', '.xls', '.xlsm'] and sheet_name is not None:
            read_params["sheet_name"] = sheet_name
            
        if ext in ['.xlsx', '.xls', '.xlsm']:
            df = pd.read_excel(file_path, **read_params)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext == '.tsv':
            df = pd.read_csv(file_path, sep='\t')
        elif ext == '.json':
            df = pd.read_json(file_path)
        else:
            return f"Unsupported file extension: {ext}"
        
        # Configure pivot table params
        pivot_params = {"index": index}
        if columns:
            pivot_params["columns"] = columns
        if values:
            pivot_params["values"] = values
            
        # Map string aggfunc to actual function
        if aggfunc == "mean":
            pivot_params["aggfunc"] = np.mean
        elif aggfunc == "sum":
            pivot_params["aggfunc"] = np.sum
        elif aggfunc == "count":
            pivot_params["aggfunc"] = len
        elif aggfunc == "min":
            pivot_params["aggfunc"] = np.min
        elif aggfunc == "max":
            pivot_params["aggfunc"] = np.max
        else:
            return f"Unsupported aggregation function: {aggfunc}"
        
        # Create pivot table
        pivot = pd.pivot_table(df, **pivot_params)
        
        return pivot.to_string()
    except Exception as e:
        return f"Error creating pivot table: {str(e)}"

@mcp.tool()
def export_chart(file_path: str, x_column: str, y_column: str, 
               chart_type: str = "line", sheet_name: Optional[str] = None) -> Image:
    """
    Create a chart from Excel data and return as an image.
    
    Args:
        file_path: Path to the Excel file
        x_column: Column to use for x-axis
        y_column: Column to use for y-axis
        chart_type: Type of chart ('line', 'bar', 'scatter', 'hist')
        sheet_name: Name of the sheet to chart (for Excel files)
        
    Returns:
        Chart as image
    """
    import matplotlib.pyplot as plt
    import seaborn as sns
    
    try:
        # Read file
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        read_params = {}
        if ext in ['.xlsx', '.xls', '.xlsm'] and sheet_name is not None:
            read_params["sheet_name"] = sheet_name
            
        if ext in ['.xlsx', '.xls', '.xlsm']:
            df = pd.read_excel(file_path, **read_params)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext == '.tsv':
            df = pd.read_csv(file_path, sep='\t')
        elif ext == '.json':
            df = pd.read_json(file_path)
        else:
            raise ValueError(f"Unsupported file extension: {ext}")
        
        # Create chart
        plt.figure(figsize=(10, 6))
        
        if chart_type == "line":
            sns.lineplot(data=df, x=x_column, y=y_column)
        elif chart_type == "bar":
            sns.barplot(data=df, x=x_column, y=y_column)
        elif chart_type == "scatter":
            sns.scatterplot(data=df, x=x_column, y=y_column)
        elif chart_type == "hist":
            df[y_column].hist()
            plt.xlabel(y_column)
        else:
            raise ValueError(f"Unsupported chart type: {chart_type}")
        
        plt.title(f"{chart_type.capitalize()} Chart: {y_column} by {x_column}")
        plt.tight_layout()
        
        # Save to bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        
        # Convert to Image
        plt.close()
        return Image(data=buf.getvalue(), format="png")
    except Exception as e:
        # Return error image
        plt.figure(figsize=(8, 2))
        plt.text(0.5, 0.5, f"Error creating chart: {str(e)}", 
                 horizontalalignment='center', fontsize=12, color='red')
        plt.axis('off')
        
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        return Image(data=buf.getvalue(), format="png")

@mcp.tool()
def data_summary(file_path: str, sheet_name: Optional[str] = None) -> str:
    """
    Generate a comprehensive summary of the data in an Excel file.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to summarize (for Excel files)
        
    Returns:
        Comprehensive data summary as string
    """
    try:
        # Read file
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        read_params = {}
        if ext in ['.xlsx', '.xls', '.xlsm'] and sheet_name is not None:
            read_params["sheet_name"] = sheet_name
            
        if ext in ['.xlsx', '.xls', '.xlsm']:
            df = pd.read_excel(file_path, **read_params)
        elif ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext == '.tsv':
            df = pd.read_csv(file_path, sep='\t')
        elif ext == '.json':
            df = pd.read_json(file_path)
        else:
            return f"Unsupported file extension: {ext}"
        
        # Basic file info
        file_info = {
            "file_name": os.path.basename(file_path),
            "file_type": ext,
            "file_size": f"{os.path.getsize(file_path) / 1024:.2f} KB",
            "last_modified": datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Data structure
        data_structure = {
            "rows": df.shape[0],
            "columns": df.shape[1],
            "column_names": list(df.columns),
            "column_types": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "memory_usage": f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB"
        }
        
        # Data quality
        data_quality = {
            "missing_values": {col: int(count) for col, count in df.isnull().sum().items()},
            "missing_percentage": {col: f"{count/len(df)*100:.2f}%" for col, count in df.isnull().sum().items()},
            "duplicate_rows": int(df.duplicated().sum()),
            "unique_values": {col: int(df[col].nunique()) for col in df.columns}
        }
        
        # Statistical summary
        numeric_cols = df.select_dtypes(include=['number']).columns
        categorical_cols = df.select_dtypes(include=['object', 'category']).columns
        datetime_cols = df.select_dtypes(include=['datetime', 'datetime64']).columns
        
        statistics = {}
        if len(numeric_cols) > 0:
            statistics["numeric"] = df[numeric_cols].describe().to_dict()
        
        if len(categorical_cols) > 0:
            statistics["categorical"] = {
                col: {
                    "unique_values": int(df[col].nunique()),
                    "top_values": df[col].value_counts().head(5).to_dict()
                } for col in categorical_cols
            }
        
        if len(datetime_cols) > 0:
            statistics["datetime"] = {
                col: {
                    "min": df[col].min().strftime('%Y-%m-%d') if pd.notna(df[col].min()) else None,
                    "max": df[col].max().strftime('%Y-%m-%d') if pd.notna(df[col].max()) else None,
                    "range_days": (df[col].max() - df[col].min()).days if pd.notna(df[col].min()) and pd.notna(df[col].max()) else None
                } for col in datetime_cols
            }
        
        # Combine all info
        summary = {
            "file_info": file_info,
            "data_structure": data_structure,
            "data_quality": data_quality,
            "statistics": statistics
        }
        
        return json.dumps(summary, indent=2, default=str)
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# Add prompt templates for common Excel operations
@mcp.prompt()
def analyze_excel_data(file_path: str) -> str:
    """
    Create a prompt for analyzing Excel data
    """
    return f"""
I have an Excel file at {file_path} that I'd like to analyze. 
Could you help me understand the data structure, perform basic statistical analysis, 
and identify any patterns or insights in the data?
"""

@mcp.prompt()
def create_chart(file_path: str) -> str:
    """
    Create a prompt for generating charts from Excel data
    """
    return f"""
I have an Excel file at {file_path} and I want to create some visualizations. 
Could you suggest some appropriate charts based on the data and help me create them?
"""

@mcp.prompt()
def data_cleaning(file_path: str) -> str:
    """
    Create a prompt for cleaning and preprocessing Excel data
    """
    return f"""
I have an Excel file at {file_path} that needs some cleaning and preprocessing. 
Could you help me identify and fix issues like missing values, outliers, 
inconsistent formatting, and other data quality problems?
"""

def main():
    mcp.run()

# Main function to run server
if __name__ == "__main__":
    main()