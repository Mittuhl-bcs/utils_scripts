from pymongo import MongoClient
from datetime import datetime
import pandas as pd
from collections import Counter
from datetime import datetime
import calendar
from pymongo import MongoClient
import os




def analyze_and_store(csv_paths, mongo_uri, db_name, collection_name):
    
    # Step 3: Generate time-related metadata
    now = datetime.now()
    metadata = {
        "time_created": now.isoformat(),
        "date": now.strftime("%Y-%m-%d"),
        "week": now.strftime("%U"),
        "weekday": calendar.day_name[now.weekday()],
        "month": now.strftime("%B"),
        "year": now.year
    }

    # Step 1: Read CSV
    client = MongoClient(mongo_uri)
    db = client[db_name]
    collection = db[collection_name]

    # Step 3: Initialize discrepancies dict
    discrepancies = {}

    for i in csv_paths:
        df = pd.read_csv(i)

        item_name = item_checker(i)

        if "discrepancy_type" not in df.columns:
            raise ValueError("The column 'discrepancy_type' was not found in the CSV.")

        # Step 2: Process discrepancy_type values
        all_types = []
        for item in df["discrepancy_type"].dropna():
            all_types.extend(part.strip() for part in str(item).split("-"))

        counts = Counter(all_types)

        
        # Step 5: Add to discrepancies dict
        discrepancies[item_name] = {
            "dis_type": dict(counts)
        }

    # Step 6: Prepare final document
    document = {
        **metadata,
        "discrepancies": discrepancies
    }

    # Step 7: Insert into MongoDB
    collection.insert_one(document)


def item_checker(csv_path):

    # Extract the filename from the file path
    former_filename = os.path.basename(csv_path)
    base_filename = former_filename.split('{')[0]  # Split on '{' and take the first part
    base_filename = base_filename.rstrip('_')  # Remove any trailing underscores

    # Get current time
    current_time = datetime.now()
    formatted_time = current_time.strftime("%Y-%m-%d %H:%M:%S")

    filename = base_filename  # get the i and then extract just the filename

    if "Blue" in base_filename:
        category = "Blue"
    elif "Purple" in base_filename:
        category = "Purple"
    elif "Orange" in base_filename:
        if "Orange - Price" in base_filename:
            category = "Orange - Price"
        elif "Secondary Orange items - standard" in base_filename:
            category = "Orange - Standard (secondary)"
        elif "Orange items - standard" in base_filename:
            category = "Orange - Standard (primary)"

    else:
        raise ValueError
    
    return category

# Function to get all the files generated today
def get_files_from_folders(folders, extensions=[".csv"]):
    today = datetime.now().date()
    file_list = []
    
    for folder in folders:
        for root, dirs, files in os.walk(folder):
            for file in files:
                if any(file.endswith(ext) for ext in extensions):
                    file_path = os.path.join(root, file)
                    file_modified_date = datetime.fromtimestamp(os.path.getmtime(file_path)).date()
                    
                    if file_modified_date == today:
                        file_list.append(file_path)
    return file_list


def main():

    folders = [
        "D:\\Temp_items_discrepancy_reports",
        "D:\\Price_mapping_discrepancies"
    ]


    file_list = get_files_from_folders(folders)

    # Filter to keep only files that don't have "formatted" in their names
    filtered_paths = [path for path in file_list if "Formatted" not in path]

    # Example usage
    mongo_uri = "mongodb://localhost:27017/"
    db_name = "item_discrepancies"
    collection_name = "discrepancy_types"
    

    # time created, date, week, weekday, month, year, category, D_type, count
    analyze_and_store(filtered_paths, mongo_uri, db_name, collection_name)

