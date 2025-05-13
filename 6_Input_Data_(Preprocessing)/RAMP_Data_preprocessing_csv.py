import pandas as pd

# Load the CSV file
file_path = 'BFlexLM_10_Hosuehold(s)_Winter_Week.csv'  # Replace with your file path
data = pd.read_csv(file_path, parse_dates=['Timeseries'])

# Set the 'Timeseries' column as the index for easier resampling
data.set_index('Timeseries', inplace=True)

# Resample to 15-minute intervals, taking the average power consumption
data_15min = data.resample('15T').mean()

# Save the resampled data to a new CSV file
output_file_path = 'Flexible_Loads_10_Hosuehold(s)_Winter_Week_15_min.csv'  # Replace with your desired path
data_15min.to_csv(output_file_path)

print("Resampling complete. File saved to:", output_file_path)