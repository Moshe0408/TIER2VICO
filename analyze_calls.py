import pandas as pd
import os
import glob
import json

def process_duration(dur_str):
    try:
        parts = dur_str.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        return 0
    except:
        return 0

def consolidate_data(directory):
    all_files = glob.glob(os.path.join(directory, "*.csv"))
    df_list = []
    
    for filename in all_files:
        try:
            # Handling potential encoding issues with Hebrew filenames or content
            df = pd.read_csv(filename, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(filename, encoding='latin1')
        
        # Add month-year based on filename if not in data, but it is in 'Start Time'
        df_list.append(df)
        
    master_df = pd.concat(df_list, ignore_index=True)
    
    # Convert 'Start Time' to datetime
    master_df['Start Time'] = pd.to_datetime(master_df['Start Time'], dayfirst=True)
    
    # Extract month and day of week
    master_df['Month'] = master_df['Start Time'].dt.month
    master_df['MonthName'] = master_df['Start Time'].dt.strftime('%B')
    master_df['DayOfWeek'] = master_df['Start Time'].dt.strftime('%A')
    master_df['Hour'] = master_df['Start Time'].dt.hour
    
    # Convert durations to seconds
    master_df['Duration_Sec'] = master_df['Interaction Duration'].apply(process_duration)
    master_df['Hold_Sec'] = master_df['Interaction Total Hold Time'].apply(process_duration)
    
    return master_df

def generate_stats(df):
    stats = {}
    
    # Mapping phone numbers to names
    # Dialed From (ANI)
    df['Dialed From (ANI)'] = df['Dialed From (ANI)'].astype(str).str.replace('+', '', regex=False)
    
    # Dialed To (DNIS)
    df['Dialed To (DNIS)'] = df['Dialed To (DNIS)'].astype(str).str.replace('+', '', regex=False)
    
    # Specific counts based on requirements
    vico_outgoing = df[df['Dialed From (ANI)'] == '97239029740']
    tier1_calls = df[df['Dialed To (DNIS)'] == '97239029740']
    verticals_calls = df[df['Dialed To (DNIS)'] == '972732069574']
    shufersal_calls = df[df['Dialed To (DNIS)'] == '972732069576']
    
    # Total stats
    stats['total_calls'] = len(df)
    stats['vico_outgoing_count'] = len(vico_outgoing)
    stats['tier1_count'] = len(tier1_calls)
    stats['verticals_count'] = len(verticals_calls)
    stats['shufersal_count'] = len(shufersal_calls)
    stats['tier1_plus_vico_count'] = len(tier1_calls) + len(vico_outgoing)
    
    stats['avg_duration_sec'] = round(df['Duration_Sec'].mean(), 2)
    stats['avg_duration_min'] = round(df['Duration_Sec'].mean() / 60, 2)
    
    # Monthly trends
    monthly_counts = df.groupby(['Month', 'MonthName']).size().reset_index(name='count')
    stats['monthly_trends'] = monthly_counts.sort_values('Month').to_dict(orient='records')
    
    # Employee stats
    employee_stats = df.groupby('Employee').agg({
        'Start Time': 'count',
        'Duration_Sec': 'mean'
    }).rename(columns={'Start Time': 'count', 'Duration_Sec': 'avg_duration'}).reset_index()
    
    # Round avg_duration to 2 decimal places and convert to minutes for readability
    employee_stats['avg_duration_min'] = (employee_stats['avg_duration'] / 60).round(2)
    stats['employee_performance'] = employee_stats.to_dict(orient='records')
    
    # Hourly distribution
    hourly_dist = df.groupby('Hour').size().reset_index(name='count')
    stats['hourly_distribution'] = hourly_dist.to_dict(orient='records')
    
    # Day of week distribution
    dow_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    dow_dist = df.groupby('DayOfWeek').size().reindex(dow_order).reset_index(name='count').dropna()
    stats['dow_distribution'] = dow_dist.to_dict(orient='records')

    # FCR Calculation (Refined: Monthly One-Time Callers)
    # We calculate FCR per month to avoid penalizing customers returning in different months
    fcr_filtered_df = df[df['Dialed To (DNIS)'] != '972732069576']
    
    monthly_fcr_rates = []
    for month in fcr_filtered_df['Month'].unique():
        month_data = fcr_filtered_df[fcr_filtered_df['Month'] == month]
        if len(month_data) == 0: continue
        
        # Count how many unique callers called only ONCE in this month
        caller_counts = month_data.groupby('Dialed From (ANI)').size()
        one_time_callers = len(caller_counts[caller_counts == 1])
        total_unique_callers = len(caller_counts)
        
        month_rate = (one_time_callers / total_unique_callers) * 100
        monthly_fcr_rates.append(month_rate)
    
    if monthly_fcr_rates:
        stats['fcr_rate'] = round(sum(monthly_fcr_rates) / len(monthly_fcr_rates), 1)
    else:
        stats['fcr_rate'] = 0

    return stats

if __name__ == "__main__":
    dir_path = r"c:\Users\Moshei1\OneDrive - Verifone\Desktop\ישיבת צוות 2026\טיר 2 - 2025"
    master_df = consolidate_data(dir_path)
    stats = generate_stats(master_df)
    
    with open('call_stats.json', 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=4)
    
    print("Data consolidation and analysis complete. Stats saved to call_stats.json")
