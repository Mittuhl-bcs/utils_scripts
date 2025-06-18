import ollama
import pandas as pd
import os
from datetime import datetime
import json


def load_excel_data(file_path, sheet_name=None):
    """
    Load data from Excel file into a DataFrame
    
    Args:
        file_path (str): Path to Excel file
        sheet_name (str, optional): Name of sheet to read. If None, reads first sheet
        
    Returns:
        pd.DataFrame: Loaded data
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        print(f"Error loading Excel file: {str(e)}")
        return None


def format_dataframe_for_model(df):
    """
    Format DataFrame into a clear, structured text format for the AI model
    
    Args:
        df (pd.DataFrame): DataFrame containing the dashboard data
        
    Returns:
        str: Formatted string representation of the data
    """
    if df is None or df.empty:
        return "No data available for analysis"
    
    # Create a structured text format
    formatted_text = "DASHBOARD DATA TABLE:\n"
    formatted_text += "=" * 80 + "\n\n"
    
    # Add column headers with clear separation
    columns = df.columns.tolist()
    formatted_text += "COLUMNS: " + " | ".join(columns) + "\n"
    formatted_text += "-" * 80 + "\n"
    
    # Add each row with clear category identification
    for index, row in df.iterrows():
        formatted_text += f"\nROW {index + 1}:\n"
        for col in columns:
            value = row[col]
            # Handle different data types
            if pd.isna(value):
                value_str = "N/A"
            elif isinstance(value, (int, float)):
                value_str = str(int(value)) if isinstance(value, float) and value.is_integer() else str(value)
            else:
                value_str = str(value)
            
            formatted_text += f"  {col}: {value_str}\n"
        formatted_text += "-" * 40 + "\n"
    
    # Add summary statistics
    formatted_text += f"\nDATA SUMMARY:\n"
    formatted_text += f"Total Categories: {len(df)}\n"
    formatted_text += f"Total Columns: {len(df.columns)}\n"
    
    # Add key metrics summary if they exist
    numeric_columns = df.select_dtypes(include=['number']).columns
    if len(numeric_columns) > 0:
        formatted_text += f"\nKEY METRICS OVERVIEW:\n"
        for col in numeric_columns:
            if 'difference' in col.lower() or 'diff' in col.lower():
                positive_count = (df[col] > 0).sum()
                negative_count = (df[col] < 0).sum()
                zero_count = (df[col] == 0).sum()
                formatted_text += f"  {col}: {positive_count} increases, {negative_count} decreases, {zero_count} unchanged\n"
    
    return formatted_text


def create_analysis_prompt(formatted_data):
    """
    Create a comprehensive prompt for analysis
    
    Args:
        formatted_data (str): Formatted data string
        
    Returns:
        str: Complete analysis prompt
    """
    prompt = f"""Please analyze the following dashboard data. Each ROW represents a different category with its associated metrics.

If it is decreasing, then it is positive, if it is increasing, then it is concerning.

{formatted_data}

ANALYSIS REQUIREMENTS:
1. For each category, analyze the key performance indicators
2. Focus on these critical metrics:
   - daily_difference: Day-over-day changes (increases are concerning, decreasing(-) is good)
   - week_max_difference: Weekly maximum changes
   - week_total_difference: Weekly total changes
   - last_two_weeks_total_diff: Two-week trend analysis

3. Identify patterns and trends:
   - Which categories show concerning increases?
   - Which categories show positive decreases or stability?
   - Are there any anomalies or outliers?

4. Provide category-specific insights with exact numbers from the data.

Please structure your response in the format specified in the system prompt."""

    return prompt


def analyze_dashboard_dataframe(df, model="gemma3"):
    """
    Analyze DataFrame data using the specified model
    
    Args:
        df (pd.DataFrame): DataFrame containing dashboard data
        model (str): Model name to use for analysis
        
    Returns:
        str: Analysis results
    """
    # System prompt for structured analysis
    system_prompt = """You are a specialized data dashboard analyst. Remember increasing numbers are a concern, if the numbers are declining then it is positive. When analyzing dashboard data:

Structure your analysis as follows:

## Executive Summary
[Brief overview of overall data health]

## Category-by-Category Analysis
For each category, provide:
- **[Category Name]**:
  - Current Status: [latest count vs previous count]
  - Daily Trend: [daily difference with direction] (increasing is concerning, decreasing(-) is good)
  - Weekly Trend: [weekly changes] (increasing is concerning, decreasing(-) is good)
  - Assessment: [Positive/Concerning/Neutral with reasoning] (increasing is concerning, decreasing(-) is good)

## Weekly Performance Overview
- **Max Count Changes**: [Analysis of week_max_difference across categories] (increasing is concerning, decreasing(-) is good)
- **Total Count Changes**: [Analysis of week_total_difference across categories] (increasing is concerning, decreasing(-) is good)

## Key Findings 
- Most Concerning: [Categories with worrying increases] (increasing is concerning, decreasing(-) is good)
- Most Improved: [Categories with positive decreases] 
- Stable Categories: [Categories with minimal changes]

Always reference specific numbers from the data and explain why trends are positive or concerning."""

    # Format the DataFrame for the model
    formatted_data = format_dataframe_for_model(df)
    
    # Create the analysis prompt
    user_prompt = create_analysis_prompt(formatted_data)
    
    try:
        # Send to Ollama
        response = ollama.chat(
            model=model,
            messages=[
                {
                    'role': 'system',
                    'content': system_prompt
                },
                {
                    'role': 'user',
                    'content': user_prompt
                }
            ]
        )
        
        return response['message']['content']
    
    except Exception as e:
        return f"Error during analysis: {str(e)}\nPlease ensure Ollama is running and the {model} model is available."


def create_sample_dataframe():
    """
    Create a sample DataFrame from the provided data for testing
    
    Returns:
        pd.DataFrame: Sample data as DataFrame
    """
    data = {
        'category': [
            'Blue',
            'Orange - Price',
            'Orange - Standard (primary)',
            'Orange - Standard (secondary)',
            'Purple',
            'duplicates'
        ],
        'latest_date': ['########'] * 6,
        'previous_date': ['########'] * 6,
        'latest_count': [12870, 77209, 31181, 403429, 22086, 6],
        'previous_count': [12845, 77084, 54905, 403430, 16968, 279],
        'daily_difference': [25, 125, -23724, -1, 5118, -273],
        'current_week_max': [12870, 77209, 31181, 403429, 22086, 6],
        'last_week_max': [12873, 77084, 79075, 403430, 16954, 279],
        'week_max_difference': [-3, 125, -47894, -1, 5132, -273],
        'current_week_total': [25715, 77209, 31181, 403429, 39054, 6],
        'last_week_total': [25717, 225165, 133980, 806817, 33775, 279],
        'week_total_difference': [-2, -147956, -102799, -403388, 5279, -273],
        'second_last_week_total': [25162, 195449, 153834, 805121, 33455, None],
        'last_two_weeks_total_diff': [555, 29716, -19854, 1696, 320, None]
    }
    
    return pd.DataFrame(data)


def save_analysis_report(analysis, df_info="", filename_prefix="dashboard_analysis"):
    """
    Save analysis to a markdown file with timestamp and data info
    
    Args:
        analysis (str): Analysis content
        df_info (str): Information about the DataFrame
        filename_prefix (str): Prefix for the filename
        
    Returns:
        str: Generated filename
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.md"
    
    with open(filename, "w", encoding='utf-8') as f:
        f.write(f"# Dashboard Analysis Report\n")
        f.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        if df_info:
            f.write(f"## Data Information\n{df_info}\n\n")
        
        f.write(f"## Analysis Results\n")
        f.write(analysis)
    
    return filename


# Main execution functions
def analyze_excel_file(file_path, sheet_name=None, model="gemma3"):
    """
    Complete workflow: Load Excel file and analyze
    
    Args:
        file_path (str): Path to Excel file
        sheet_name (str, optional): Sheet name to read
        model (str): Model to use for analysis
        
    Returns:
        tuple: (analysis_result, dataframe, saved_filename)
    """
    print(f"Loading Excel file: {file_path}")
    
    # Load the data
    df = load_excel_data(file_path, sheet_name)
    
    if df is None:
        return "Failed to load Excel file", None, None
    
    print(f"Data loaded successfully. Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print(f"Categories: {df['category'].tolist() if 'category' in df.columns else 'No category column found'}")
    
    # Perform analysis
    print("\nAnalyzing data with AI model...")
    analysis = analyze_dashboard_dataframe(df, model)
    
    # Save results
    df_info = f"Shape: {df.shape}\nColumns: {list(df.columns)}"
    saved_file = save_analysis_report(analysis, df_info)
    
    return analysis, df, saved_file


# Example usage
if __name__ == "__main__":
    print("\n=== DASHBOARD ANALYSIS SYSTEM ===\n")
    
    # Option 1: Use sample data for testing
    print("1. Testing with sample data...")
    sample_df = create_sample_dataframe()
    
    print(f"Sample data shape: {sample_df.shape}")
    print("Sample categories:", sample_df['category'].tolist())
    
    analysis = analyze_dashboard_dataframe(sample_df)
    print("\n" + "="*60)
    print("ANALYSIS RESULTS:")
    print("="*60)
    print(analysis)
    
    # Save sample analysis
    saved_file = save_analysis_report(analysis, f"Sample data shape: {sample_df.shape}")
    print(f"\nAnalysis saved to: {saved_file}")
    
    print("\n" + "="*60)
    
    # Option 2: Uncomment to use with actual Excel file
    """
    excel_file_path = "path/to/your/dashboard_data.xlsx"  # Replace with your file path
    
    if os.path.exists(excel_file_path):
        print("2. Analyzing Excel file...")
        analysis, df, saved_file = analyze_excel_file(excel_file_path)
        
        print("\n" + "="*60)
        print("EXCEL FILE ANALYSIS RESULTS:")
        print("="*60)
        print(analysis)
        print(f"\nResults saved to: {saved_file}")
    else:
        print(f"Excel file not found: {excel_file_path}")
    """
    
   