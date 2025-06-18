import ollama
import os
import base64
import json
from datetime import datetime

def analyze_dashboard_with_gemma(image_path: str):
    """Analyze dashboard image using only Gemma3 with enhanced prompting for numerical accuracy"""
    
    # Check if image exists
    if not os.path.exists(image_path):
        print(f"Error: Image file not found at {image_path}")
        return
    
    # Read and encode the image file
    with open(image_path, "rb") as image_file:
        image_data = base64.b64encode(image_file.read()).decode("utf-8")
    
    # Comprehensive system prompt with emphasis on numerical accuracy
    gemma_system_prompt = """
    You are a precision data analytics expert specializing in extracting exact numerical values from dashboards.
    Your primary function is to read numbers with 100% accuracy and report them exactly as they appear.
    
    Critical instructions for numerical extraction:
    1. For each number, examine it DIGIT BY DIGIT to ensure absolute accuracy
    2. Double-check each digit in large numbers (e.g., 139394 vs 189885)
    3. Verify whether a number belongs to "current value" or "last week/previous run"
    4. Never approximate or round numbers
    5. If uncertain about a digit, indicate this uncertainty rather than guessing
    
    When identifying metrics:
    - GREEN numbers indicate positive/good metrics
    - RED numbers indicate negative/concerning metrics
    - Numbers in PARENTHESES are CHANGES from previous periods
    """
    
    # First step: Extract ONLY the numbers with extreme precision
    extraction_prompt = """
    This dashboard shows "Count of Rows in each of the report (Items report)".
    
    YOUR ONLY TASK is to extract the exact numerical values from the dashboard with 100% precision.
    Do not analyze or interpret - only extract numbers in a structured format.
    
    For each metric box in the dashboard:
    1. Identify the main category (Blue, Purple, Orange-Price, Orange-Primary, Orange-Secondary)
    2. Record the exact type (Total CR comparison, Max CR comparison, Daily difference)
    3. Extract the large central number (current value) - VERIFY EACH DIGIT
    4. Extract the reference value (Last week: X or Previous run: X)
    5. Extract the change value (number in parentheses with + or - sign)
    
    Format your extraction as follows for EACH metric box:
    ```
    Category: [category name]
    Metric Type: [type]
    Current Value: [exact number]
    Reference Value: [exact number]
    Change: [exact number with sign]
    Color: [green/red]
    Status: [filled/blank]
    ```
    
    For blank/empty metrics, still identify the category and type, but mark values as "blank".
    
    BE EXTREMELY CAREFUL with numbers. Check each digit. For example, confirm whether a number is 139394 or 189885.
    Do not proceed to analysis - focus ONLY on 100% accurate data extraction.
    """
    
    print("Extracting precise numerical data with Gemma3...")
    
    # First query: Extract pure numerical data with lower temperature
    try:
        extraction_response = ollama.chat(
            model="gemma3",
            messages=[
                {"role": "system", "content": gemma_system_prompt},
                {"role": "user", "content": extraction_prompt, "images": [image_data]}
            ],
            options={"temperature": 0.1}  # Correct way to set temperature in Ollama client
        )
        raw_extraction = extraction_response["message"]["content"]
    except Exception as e:
        print(f"Error during numerical extraction: {e}")
        return None
    
    # Second step: Analyze the accurately extracted data
    analysis_system_prompt = """
    You are a business intelligence analyst providing insights based on dashboard metrics.
    The data provided to you has been accurately extracted from a dashboard.
    Focus on interpreting what these numbers mean for the business.
    """
    
    analysis_prompt = f"""
    Based on these accurately extracted metrics from the "Count of Rows in each of the report (Items report)" dashboard:

    {raw_extraction}

    Provide a comprehensive analysis with these specific sections:

    1. OVERVIEW:
       - Summarize the key metrics across categories
       - Identify which categories have data and which are blank
    
    2. TREND ANALYSIS:
       - Calculate the percentage changes for key metrics
       - Identify the largest decreases and increases
       - Compare the different Orange categories (Price, Primary, Secondary)
    
    3. POTENTIAL ISSUES:
       - Based on the data, suggest whether there are data quality problems
       - Identify concerning patterns in the metrics
       - Note any inconsistencies between related metrics
    
    4. RECOMMENDATIONS:
       - Suggest specific actions based on these metrics
       - Recommend areas that need further investigation
    
    Remember to use the EXACT numbers from the extraction in your analysis.
    """
    
    print("Analyzing the accurately extracted data...")
    
    # Second query: Analyze the accurately extracted data
    try:
        analysis_response = ollama.chat(
            model="gemma3",
            messages=[
                {"role": "system", "content": analysis_system_prompt},
                {"role": "user", "content": analysis_prompt}
            ],
            options={"temperature": 0.4}  # Moderate temperature for analysis
        )
        analysis = analysis_response["message"]["content"]
    except Exception as e:
        print(f"Error during analysis phase: {e}")
        return None
    
    return {
        "raw_extraction": raw_extraction,
        "analysis": analysis
    }

if __name__ == "__main__":
    # Path to the dashboard image
    image_path = "D:\\Utils_scripts\\PowerBI_SS\\screenshots\\powerbi_dashboard_20250416_132819.png"
    
    # Create output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"dashboard_analysis_precise_gemma_{timestamp}.json"
    
    # Run analysis
    results = analyze_dashboard_with_gemma(image_path)
    
    if results:
        # Print results
        print("\n" + "="*50)
        print("RAW NUMERICAL EXTRACTION")
        print("="*50)
        print(results["raw_extraction"])
        
        print("\n" + "="*50)
        print("DASHBOARD ANALYSIS")
        print("="*50)
        print(results["analysis"])
        
        # Save results to file
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nResults saved to {output_file}")
    else:
        print("Analysis failed. Check the error message above.")