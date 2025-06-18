import cv2
import numpy as np
import pytesseract
import pandas as pd
import re
from PIL import Image
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import io
import os

def detect_grid(image):
    """
    Automatically detect grid cells in the dashboard image
    
    Args:
        image: OpenCV image object
        
    Returns:
        List of cell rectangles [x, y, w, h]
    """
    # Convert to grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Apply adaptive thresholding
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                  cv2.THRESH_BINARY_INV, 11, 2)
    
    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    
    # Filter contours by area to find the cells
    min_area = image.shape[0] * image.shape[1] / 30  # Adjust this threshold as needed
    max_area = image.shape[0] * image.shape[1] / 15
    
    cell_contours = []
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if min_area < area < max_area:
            x, y, w, h = cv2.boundingRect(cnt)
            cell_contours.append([x, y, w, h])
    
    # Sort contours by position (top-to-bottom, left-to-right)
    cell_contours.sort(key=lambda x: (x[1], x[0]))
    
    return cell_contours

def extract_text_from_cell(image, cell):
    """
    Extract text from a single cell using OCR
    
    Args:
        image: Full dashboard image
        cell: Cell coordinates [x, y, w, h]
        
    Returns:
        Extracted text from the cell
    """
    x, y, w, h = cell
    cell_img = image[y:y+h, x:x+w]
    
    # Preprocess for better OCR
    gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    
    # Use Tesseract for OCR
    text = pytesseract.image_to_string(thresh)
    return text

def extract_cell_color(image, cell):
    """
    Determine the dominant color in a cell
    
    Args:
        image: Full dashboard image
        cell: Cell coordinates [x, y, w, h]
        
    Returns:
        Dominant color name
    """
    x, y, w, h = cell
    cell_img = image[y:y+h, x:x+w]
    
    # Convert to HSV for better color detection
    hsv = cv2.cvtColor(cell_img, cv2.COLOR_BGR2HSV)
    
    # Define color ranges
    red_lower1 = np.array([0, 100, 100])
    red_upper1 = np.array([10, 255, 255])
    red_lower2 = np.array([160, 100, 100])
    red_upper2 = np.array([180, 255, 255])
    
    green_lower = np.array([40, 100, 100])
    green_upper = np.array([80, 255, 255])
    
    # Create masks
    red_mask1 = cv2.inRange(hsv, red_lower1, red_upper1)
    red_mask2 = cv2.inRange(hsv, red_lower2, red_upper2)
    red_mask = red_mask1 + red_mask2
    
    green_mask = cv2.inRange(hsv, green_lower, green_upper)
    
    # Count pixels
    red_pixels = cv2.countNonZero(red_mask)
    green_pixels = cv2.countNonZero(green_mask)
    
    # Return dominant color
    if red_pixels > green_pixels and red_pixels > 100:
        return "red"
    elif green_pixels > red_pixels and green_pixels > 100:
        return "green"
    else:
        return "neutral"

def extract_main_number(text):
    """
    Extract the main number from cell text
    
    Args:
        text: Text extracted from cell
        
    Returns:
        Main number as integer
    """
    # Try to find large numbers first
    numbers = re.findall(r'\d{5,}', text)
    if numbers:
        return int(numbers[0])
    
    # If no large numbers, try any numbers
    numbers = re.findall(r'\d+', text)
    if numbers:
        return int(numbers[0])
    
    return None

def extract_comparison_values(text):
    """
    Extract comparison values (like last week, 2W ago)
    
    Args:
        text: Text extracted from cell
        
    Returns:
        Dictionary of comparison values
    """
    values = {}
    
    # Look for patterns like "Last week: 12255 (+62)"
    last_week_match = re.search(r'Last week[:\s]+(\d+)\s*\(([+-]\d+)\)', text)
    if last_week_match:
        values['last_week'] = int(last_week_match.group(1))
        values['change'] = int(last_week_match.group(2))
    
    # Look for patterns like "2W ago: 24377 (-12122)"
    weeks_ago_match = re.search(r'(\d+)W ago[:\s]+(\d+)\s*\(([+-]\d+)\)', text)
    if weeks_ago_match:
        weeks = int(weeks_ago_match.group(1))
        values[f'{weeks}w_ago'] = int(weeks_ago_match.group(2))
        values[f'{weeks}w_change'] = int(weeks_ago_match.group(3))
    
    # Look for patterns like "Previous run: (Blank) (+12317)"
    prev_run_match = re.search(r'Previous run[:\s]+\(([^)]+)\)\s*\(([+-]\d+)\)', text)
    if prev_run_match:
        values['prev_run'] = prev_run_match.group(1)
        values['prev_change'] = int(prev_run_match.group(2))
    
    return values

def extract_trend_indicator(text, cell_color):
    """
    Determine if the trend is up or down
    
    Args:
        text: Text extracted from cell
        cell_color: Dominant color of cell
        
    Returns:
        Trend indicator (up, down, or neutral)
    """
    if '↑' in text or '↗' in text:
        return 'up'
    elif '↓' in text or '↘' in text:
        return 'down'
    elif cell_color == 'red':
        return 'up'  # Red usually indicates an increase (concern)
    elif cell_color == 'green':
        return 'down'  # Green usually indicates a decrease (good)
    else:
        return 'neutral'

def process_dashboard_image_enhanced(image_path):
    """
    Process dashboard image using automatic grid detection
    
    Args:
        image_path: Path to dashboard image
        
    Returns:
        Structured data from the dashboard
    """
    # Load image
    if isinstance(image_path, str):
        img = cv2.imread(image_path)
    else:
        # Handle case where image is already loaded
        img = image_path
        
    if img is None:
        raise ValueError("Could not load image")
    
    # Detect grid cells
    cells = detect_grid(img)
    
    # If automatic detection fails, fall back to fixed grid
    if len(cells) < 15:  # We expect at least 15 cells
        height, width, _ = img.shape
        num_rows, num_cols = 5, 4
        cell_height = height // num_rows
        cell_width = width // num_cols
        
        cells = []
        for row in range(num_rows):
            for col in range(num_cols):
                x = col * cell_width
                y = row * cell_height
                cells.append([x, y, cell_width, cell_height])
    
    # Categories and comparison types
    categories = ["Blue", "Purple", "Orange-Price", "Orange-Primary", "Orange-Secondary"]
    comparison_types = [
        "LW vs 2nd LW", 
        "Total CR comparison", 
        "Max CR comparison", 
        "Daily difference"
    ]
    
    # Process each cell
    dashboard_data = []
    
    # Visualize the cells
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    plt.figure(figsize=(15, 10))
    plt.imshow(img_rgb)
    
    cell_index = 0
    for i, cell in enumerate(cells):
        if i >= len(categories) * len(comparison_types):
            break
            
        x, y, w, h = cell
        
        # Draw rectangle around cell
        plt.gca().add_patch(Rectangle((x, y), w, h, linewidth=1, edgecolor='r', facecolor='none'))
        
        # Extract information from cell
        cell_text = extract_text_from_cell(img, cell)
        cell_color = extract_cell_color(img, cell)
        
        # Determine category and comparison type
        category_idx = i // len(comparison_types)
        comp_type_idx = i % len(comparison_types)
        
        if category_idx < len(categories) and comp_type_idx < len(comparison_types):
            category = categories[category_idx]
            comp_type = comparison_types[comp_type_idx]
            
            # Extract main number
            main_number = extract_main_number(cell_text)
            
            # Extract comparison values
            comparison_values = extract_comparison_values(cell_text)
            
            # Determine trend
            trend = extract_trend_indicator(cell_text, cell_color)
            
            # Add to dataset
            dashboard_data.append({
                'Category': category,
                'Comparison Type': comp_type,
                'Main Value': main_number,
                'Comparison Values': comparison_values,
                'Trend': trend,
                'Color': cell_color,
                'Raw Text': cell_text
            })
            
            # Annotate cell with index
            plt.text(x + w/2, y + h/2, str(cell_index), 
                     color='white', fontsize=12, 
                     ha='center', va='center',
                     bbox=dict(facecolor='black', alpha=0.5))
            cell_index += 1
    
    # Save visualization
    plt.tight_layout()
    plt.savefig('grid_detection.png')
    plt.close()
    
    # Convert to DataFrame
    df = pd.DataFrame(dashboard_data)
    
    # Create pivot tables
    main_values = df.pivot(index='Category', columns='Comparison Type', values='Main Value')
    trends = df.pivot(index='Category', columns='Comparison Type', values='Trend')
    colors = df.pivot(index='Category', columns='Comparison Type', values='Color')
    
    # Create a comprehensive table
    detailed_data = []
    for _, row in df.iterrows():
        comparison_dict = row['Comparison Values']
        
        # Format comparison info
        comparison_info = []
        for key, value in comparison_dict.items():
            comparison_info.append(f"{key}: {value}")
        
        detailed_data.append({
            'Category': row['Category'],
            'Comparison Type': row['Comparison Type'],
            'Main Value': row['Main Value'],
            'Trend': row['Trend'],
            'Status': 'Concern' if row['Color'] == 'red' else 'Good' if row['Color'] == 'green' else 'Neutral',
            'Comparison Details': '; '.join(comparison_info) if comparison_info else 'N/A'
        })
    
    detailed_df = pd.DataFrame(detailed_data)
    
    return {
        'main_values': main_values,
        'trends': trends,
        'colors': colors,
        'raw_data': df,
        'detailed_analysis': detailed_df
    }

def format_dashboard_output(results):
    """
    Format the dashboard data into a nicely formatted table
    
    Args:
        results: Results from process_dashboard_image_enhanced
        
    Returns:
        Formatted DataFrame for display/export
    """
    # Create a comprehensive table
    formatted_df = results['detailed_analysis'].copy()
    
    # Format the main value column to include trend indicators
    formatted_df['Formatted Value'] = formatted_df.apply(
        lambda row: f"{row['Main Value']} {'↑' if row['Trend'] == 'up' else '↓' if row['Trend'] == 'down' else '-'}", 
        axis=1
    )
    
    # Create a pivot table with categories as rows and comparison types as columns
    pivot_df = formatted_df.pivot(
        index='Category', 
        columns='Comparison Type', 
        values=['Formatted Value', 'Status', 'Comparison Details']
    )
    
    # Format for better display
    summary_df = formatted_df.pivot(
        index='Category',
        columns='Comparison Type',
        values='Formatted Value'
    )
    
    return {
        'detailed': formatted_df,
        'pivot': pivot_df,
        'summary': summary_df
    }

def analyze_dashboard_image(image_path):
    """
    Complete analysis pipeline for dashboard image
    
    Args:
        image_path: Path to dashboard image
        
    Returns:
        Tuple of (raw results, formatted results)
    """
    # Process image
    results = process_dashboard_image_enhanced(image_path)
    
    # Format output
    formatted = format_dashboard_output(results)
    
    # Print summary
    print("=== DASHBOARD ANALYSIS SUMMARY ===")
    print(formatted['summary'])
    
    print("\n=== DETAILED ANALYSIS ===")
    print(formatted['detailed'].to_string())
    
    return results, formatted

# Example usage
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        image_path = sys.argv[1]
        results, formatted = analyze_dashboard_image(image_path)
        
        # Save to CSV
        formatted['detailed'].to_csv("dashboard_detailed_analysis.csv", index=False)
        formatted['summary'].to_csv("dashboard_summary.csv")
        print("\nAnalysis saved to CSV files")
    else:
        print("Please provide an image path as argument")