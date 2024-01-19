import pandas as pd

def analyze_excel_file(file_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Print the name and position of employees who have worked for 7 consecutive days
    print("Employees who have worked for 7 consecutive days:")
    
    if 'Employee Name' in df.columns and 'Pay Cycle Start Date' in df.columns:
        consecutive_days = df.groupby('Employee Name')['Pay Cycle Start Date'].transform('count') >= 7
        result_consecutive_days = df[consecutive_days]
        print(result_consecutive_days[['Employee Name', 'Position ID']])
    else:
        print("Columns 'Employee Name' or 'Pay Cycle Start Date' not found in the Excel file.")

    # Print the name and position of employees with less than 10 hours between shifts but greater than 1 hour
    print("\nEmployees with less than 10 hours between shifts but greater than 1 hour:")
    
    if 'Time' in df.columns:
        time_diff = pd.to_datetime(df['Time']).diff().dt.total_seconds() / 3600
        short_breaks = (time_diff < 10) & (time_diff > 1)
        result_short_breaks = df[short_breaks]
        print(result_short_breaks[['Employee Name', 'Position ID']])
    else:
        print("Column 'Time' not found in the Excel file.")

    # Print the name and position of employees who have worked for more than 14 hours in a single shift
    print("\nEmployees who have worked for more than 14 hours in a single shift:")
    
    if 'Timecard Hours (as Time)' in df.columns:
        # Convert 'Timecard Hours (as Time)' to total hours in numeric format
        df['Timecard Hours (as Time)'] = pd.to_datetime(df['Timecard Hours (as Time)'], format='%H:%M')
        df['Total Hours'] = df['Timecard Hours (as Time)'].dt.hour + df['Timecard Hours (as Time)'].dt.minute / 60
        long_shifts = df[df['Total Hours'] > 14]
        print(long_shifts[['Employee Name', 'Position ID']])
    else:
        print("Column 'Timecard Hours (as Time)' not found in the Excel file.")

if __name__ == "__main__":
    excel_file_path = r"G:\assignment\bluejay\excel.xlsx"  
    analyze_excel_file(excel_file_path)
