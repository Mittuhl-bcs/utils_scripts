import pandas as pd
from sqlalchemy import create_engine
from urllib.parse import quote_plus
from datetime import datetime

# --- Database Credentials ---
dbname = 'BCS_items'
user = 'postgres'
raw_password = 'post@BCS'
host = 'localhost'
port = '5432'
password = quote_plus(raw_password)

# --- SQLAlchemy connection ---
conn_str = f"postgresql+psycopg2://{user}:{password}@{host}:{port}/{dbname}"
engine = create_engine(conn_str)

# --- Load Data ---
query = "SELECT * FROM latest_data"
df = pd.read_sql(query, engine)

# --- Convert date to datetime and derive week if needed ---
# --- Convert date to datetime and filter for current year ---
df['date_'] = pd.to_datetime(df['date_'])
current_year = datetime.now().year
df = df[df['date_'].dt.year == current_year]

# Derive week number if not already in the data
if 'week' not in df.columns:
    df['week'] = df['date_'].dt.isocalendar().week

# --- Compute metrics per category ---
results = []

for category, group in df.groupby('category'):
    group = group.copy()
    group.sort_values('date_', inplace=True)
    
    # --- Latest vs Previous Date ---
    latest_date = group['date_'].max()
    latest_count = group.loc[group['date_'] == latest_date, 'count_of_rows'].max()
    prev_df = group[group['date_'] < latest_date]
    
    if not prev_df.empty:
        previous_date = prev_df['date_'].max()
        previous_count = prev_df.loc[prev_df['date_'] == previous_date, 'count_of_rows'].max()
    else:
        previous_date = None
        previous_count = 0

    date_diff = latest_count - previous_count

    # --- Weekly Aggregations ---
    group.sort_values('week', inplace=True)
    unique_weeks = sorted(group['week'].unique())

    current_week_max = current_week_total = last_week_max = last_week_total = second_last_week_total = None
    week_diff_max = week_diff_total = last_two_weeks_diff = None

    if len(unique_weeks) >= 1:
        current_week = unique_weeks[-1]
        current_week_data = group[group['week'] == current_week]
        current_week_max = current_week_data['count_of_rows'].max()
        current_week_total = current_week_data['count_of_rows'].sum()

    if len(unique_weeks) >= 2:
        last_week = unique_weeks[-2]
        last_week_data = group[group['week'] == last_week]
        last_week_max = last_week_data['count_of_rows'].max()
        last_week_total = last_week_data['count_of_rows'].sum()

        week_diff_max = (current_week_max or 0) - (last_week_max or 0)
        week_diff_total = (current_week_total or 0) - (last_week_total or 0)

    if len(unique_weeks) >= 3:
        second_last_week = unique_weeks[-3]
        second_last_week_data = group[group['week'] == second_last_week]
        second_last_week_total = second_last_week_data['count_of_rows'].sum()

        last_two_weeks_diff = (last_week_total or 0) - (second_last_week_total or 0)

    results.append({
        'category': category,
        'latest_date': latest_date,
        'previous_date': previous_date,
        'latest_count': latest_count,
        'previous_count': previous_count,
        'daily_difference': date_diff,
        'current_week_max': current_week_max,
        'last_week_max': last_week_max,
        'week_max_difference': week_diff_max,
        'current_week_total': current_week_total,
        'last_week_total': last_week_total,
        'week_total_difference': week_diff_total,
        'second_last_week_total': second_last_week_total,
        'last_two_weeks_total_diff': last_two_weeks_diff
    })

# --- Convert to DataFrame and show results ---
result_df = pd.DataFrame(results)
print(result_df)
result_df.to_excel("results.xlsx")
