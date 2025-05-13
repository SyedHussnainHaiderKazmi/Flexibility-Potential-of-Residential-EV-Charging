import pandas as pd
from datetime import datetime, timedelta

# Function to round to the nearest 15 minutes
def round_to_nearest_15_minutes(dt):
    new_minute = (dt.minute // 15) * 15
    if dt.minute % 15 >= 8:  # Round up if halfway or more
        new_minute += 15
    return dt.replace(minute=0, second=0) + timedelta(minutes=new_minute)

# Step 1: Read the trip data from an Excel file
input_path = "ROW-E 397E_Winter_Week_input.xlsx"  # Replace with your actual file path
trips_df = pd.read_excel(input_path)

# Convert departure and arrival times to datetime and round them to the nearest 15 minutes
trips_df['departure_time'] = pd.to_datetime(trips_df['departure_time']).apply(round_to_nearest_15_minutes)
trips_df['arrival_time'] = pd.to_datetime(trips_df['arrival_time']).apply(round_to_nearest_15_minutes)

# Step 2: Define the time range and create the time series with 15-minute intervals
start_time = datetime(2024, 1, 15, 0, 0, 0)
end_time = datetime(2024, 1, 21, 23, 45, 0)
time_index = pd.date_range(start=start_time, end=end_time, freq='15T')

# Create a DataFrame to hold the time series, initializing "at_home" as 1 and "distance" as None
time_series_df = pd.DataFrame({'time': time_index, 'at_home': 1, 'distance': None})

# Step 3: Mark "0" (away from home) during each rounded trip duration and add the distance value
for _, row in trips_df.iterrows():
    mask = (time_series_df['time'] >= row['departure_time']) & (time_series_df['time'] <= row['arrival_time'])
    time_series_df.loc[mask, 'at_home'] = 0
    
    # Only set the distance at the very first time step of the away interval
    first_away_index = time_series_df[mask].index[0]
    time_series_df.at[first_away_index, 'distance'] = row['distance']

# Step 4: Write the modified trips and the output time series to Excel
output_path = "ROW-E 397E_Winter_Week_output.xlsx"  # Specify your desired output file path
with pd.ExcelWriter(output_path) as writer:
    trips_df.to_excel(writer, sheet_name="Rounded_Trips", index=False)  # Save rounded trip data
    time_series_df.to_excel(writer, sheet_name="Time_Series", index=False)  # Save time series data

print("Time series file with distance data created successfully at:", output_path)
