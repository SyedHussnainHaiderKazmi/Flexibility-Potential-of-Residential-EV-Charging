import pandas as pd

# Load the Excel file
file_path = 'BFlexLM_7_days_2024-11-24_21-10-32_(modified).xlsx'  # Replace with your file path
data = pd.read_excel(file_path, parse_dates=['Timeseries'])

# Set the 'Timeseries' column as the index for easier resampling
data.set_index('Timeseries', inplace=True)

# Resample to 15-minute intervals, taking the average power consumption
data_15min = data.resample('15T').mean()

# Save the resampled data to a new Excel file
output_file_path = 'BFlexLM_7_days_2024-11-24_21-10-32_(modified)_15_min.xlsx'  # Replace with your desired path
data_15min.to_excel(output_file_path)

print("Resampling complete. File saved to:", output_file_path)
