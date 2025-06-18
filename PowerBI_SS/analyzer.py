import ollama
import os
import base64
from PIL import Image
import io
import requests
import json


def encode_image_to_base64(image_path):
    """Convert an image to base64 encoding"""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def analyze_dashboard(model="gemma3"):
    """
    Analyze the given data and return structured insights
    
    Args:
        model (str): The model name to use
        
    Returns:
        str: Formatted analysis of the dashboard
    """
    # System prompt to guide the model's analysis
    system_prompt = """You are a specialized data dashboard analyst. When given table of data metrics:
    
1. Analyze the metrics for total and max count_of_rows comparison for various categories
2. Focus on identifying decreases or stagnation in data rows across categories
3. Pay special attention to any increasing trends, as these are concerning
4. Structure your analysis consistently in the following format:

## Weekly Metrics Analysis
- **Total Count of Rows (Weekly)**: [Provide analysis with specific numbers and percentage changes]
- **Max Count of Rows (Weekly)**: [Provide analysis with specific numbers and percentage changes]

## Daily Metrics Analysis
- **Daily Count of Rows**: [Provide analysis with specific numbers and day-over-day changes]

## Category Performance
- [Category Name]: [Trend direction] by [X%] - [Brief insight]
- [Category Name]: [Trend direction] by [X%] - [Brief insight]

## Overall Assessment
[Provide a brief summary of the dashboard health, highlighting any concerns or positive indicators]

Remember to extract specific numbers when possible and always indicate whether trends are positive, negative, or neutral from a data health perspective.
"""

    # User prompt focused on dashboard analysis
    user_prompt = "Analyze these data given in a tabular format. Focus on total and max count of rows comparison on weekly basis, and daily count of rows comparison. Identify any concerning trends (increases) or positive indicators (decreases or stability)."
    
    # Send to Ollama with system prompt
    response = ollama.chat(
        model=model,
        messages=[
            {a
                'role': 'system',
                'content': system_prompt
            },
            {
                'role': 'user',
                'content': user_prompt,
            }
        ]
    )
    
    return response['message']['content']

# Example usage
if __name__ == "__main__":
    image_path = "D:\\Utils_scripts\\screenshots\\powerbi_dashboard_20250416_134459.png"
    print("\n=== DASHBOARD ANALYSIS REPORT ===\n")
    analysis = analyze_dashboard(image_path)
    print(analysis)
    print("\n================================\n")
    
    # Optionally save the analysis to a file
    with open("dashboard_analysis.md", "w") as f:
        f.write(analysis)
    print("Analysis saved to dashboard_analysis.md")