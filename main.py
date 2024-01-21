import pandas as pd
import sys


input_file_path = 'Assignment_Timecard.xlsx'

df = pd.read_excel(input_file_path)

df['Time'] = pd.to_datetime(df['Time'])
df['Time Out'] = pd.to_datetime(df['Time Out'])

df.sort_values(by='Time', inplace=True)

consecutive_days = df.groupby('Employee Name')['Time'].diff().dt.days
employees_7_consecutive_days = df[consecutive_days == 1].groupby('Employee Name').size()
employees_7_consecutive_days = employees_7_consecutive_days[employees_7_consecutive_days >= 6].index.tolist()



original_stdout = sys.stdout
output_file_path = 'output.txt'


try:
    
    with open(output_file_path, 'w') as file:
        
        sys.stdout = file
        
        print("Employees who have worked for 7 consecutive days:")
        print(employees_7_consecutive_days)
        
        time_between_shifts = df.groupby('Employee Name')['Time'].diff().dt.total_seconds() / 3600
        employees_less_than_10_hours = df[(time_between_shifts < 10) & (time_between_shifts > 1)]['Employee Name'].unique()

        print("\nEmployees who have less than 10 hours between shifts but greater than 1 hour:")
        print(employees_less_than_10_hours)
        
        df['Shift Duration'] = (df['Time Out'] - df['Time']).dt.total_seconds() / 3600
        employees_more_than_14_hours = df[df['Shift Duration'] > 14]['Employee Name'].unique()

        print("\nEmployees who have worked for more than 14 hours in a single shift:")
        print(employees_more_than_14_hours)
        # print(data)

finally:
    
    sys.stdout = original_stdout
