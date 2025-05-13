# ComfficientShare Building_Fixed_Loads_Model
# This script simulates the power consumption of a ComfficientSahre Building with 10 Household(s) (each Household contains multiple fixed appliances)!

# Importing necessary libraries from the RAMP framework and other useful Python modules
from ramp import User, UseCase
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import os

# Create a user category for the household
household = User(
    user_name="Household",  # The name of the user category
    num_users=10,           # Number of users in this category (10 households)
)

# Add an appliance (Indoor Light Bulb) to the household user category
indoor_bulb = household.add_appliance(
    name="Indoor Light Bulb",                   # The name of the appliance
    number=20,                                  # Number of bulbs in the apartment
    power=0.01,                                 # Power usage of each appliance in Kilowatt (kW)
    num_windows=2,                              # Number of time windows throughout the day when this appliance is used
    func_time=300,                              # Total usage time of the appliance in minutes per day
    func_cycle=15,                              # Minimum usage time after a switch-on event in minutes
    window_1=[480, 840],                        # First usage window (8:00 AM to 2:00 PM)
    window_2=[1080, 1440],                      # Second usage window (6:00 PM to midnight)
    random_var_w=0.25,                          # Variability of the windows in percentage
    time_fraction_random_variability=0.1,       # Randomness in total usage time (between 0 and 1)
)

# Add additional fixed appliances to the household user category

# Add a Television to the household
television = household.add_appliance(
    name="Television",                          # The name of the appliance
    number=2,                                   # Number of TVs in the apartment
    power=0.15,                                 # Power usage of each appliance in Kilowatt (kW)
    num_windows=2,                              # Number of time windows throughout the day when this appliance is used
    func_time=240,                              # Total usage time of the appliance in minutes per day
    func_cycle=30,                              # Minimum usage time after a switch-on event in minutes
    window_1=[720, 840],                        # First usage window (12:00 PM to 2:00 PM)
    window_2=[1080, 1440],                      # Second usage window (6:00 PM to midnight)
    random_var_w=0.2,                           # Variability of the windows in percentage
    time_fraction_random_variability=0.1,       # Randomness in total usage time (between 0 and 1)
)

# Add a Radio to the household
radio = household.add_appliance(
    name="Radio",                               # The name of the appliance
    number=1,                                   # Number of radios in the apartment
    power=0.01,                                 # Power usage of each appliance in Kilowatt (kW)
    num_windows=2,                              # Number of time windows throughout the day when this appliance is used
    func_time=120,                              # Total usage time of the appliance in minutes per day
    func_cycle=15,                              # Minimum usage time after a switch-on event in minutes
    window_1=[420, 600],                        # First usage window (7:00 AM to 10:00 AM)
    window_2=[1080, 1260],                      # Second usage window (6:00 PM to 9:00 PM)
    random_var_w=0.3,                           # Variability of the windows in percentage
    time_fraction_random_variability=0.15,      # Randomness in total usage time (between 0 and 1)
)

# Add a Fridge to the household
fridge = household.add_appliance(
    name="Fridge",                              # The name of the appliance
    number=1,                                   # Number of fridges in the apartment
    power=0.15,                                 # Power usage of each appliance in Kilowatt (kW)
    num_windows=1,                              # Number of time windows throughout the day when this appliance is used
    func_time=1440,                             # Total usage time of the appliance in minutes per day (24 hours)
    func_cycle=30,                              # Minimum usage time after a switch-on event in minutes
    window_1=[0, 1440],                         # Continuous usage window (24/7)
    window_2=None,                              # No second window
    random_var_w=0.05,                          # Variability of the windows in percentage
    time_fraction_random_variability=0.05,      # Randomness in total usage time (between 0 and 1)
)

# Add a Household Computer to the household
household_computer = household.add_appliance(
    name="Household Computer",                  # The name of the appliance
    number=2,                                   # Number of computers in the apartment
    power=0.25,                                 # Power usage of each appliance in Kilowatt (kW)
    num_windows=2,                              # Number of time windows throughout the day when this appliance is used
    func_time=300,                              # Total usage time of the appliance in minutes per day
    func_cycle=60,                              # Minimum usage time after a switch-on event in minutes
    window_1=[540, 780],                        # First usage window (9:00 AM to 1:00 PM)
    window_2=[840, 1140],                       # Second usage window (2:00 PM to 7:00 PM)
    random_var_w=0.2,                           # Variability of the windows in percentage
    time_fraction_random_variability=0.15,      # Randomness in total usage time (between 0 and 1)
)

# # Add Mobile Phones to the household
# mobile_phones = household.add_appliance(
#     name="Mobile Phones",                     # The name of the appliance
#     number=4,                                 # Number of mobile phones in the apartment
#     power=0.005,                              # Power usage of each appliance in Kilowatt (kW) while charging
#     num_windows=2,                            # Number of time windows throughout the day when this appliance is used
#     func_time=180,                            # Total usage time of the appliance in minutes per day
#     func_cycle=60,                            # Minimum usage time after a switch-on event in minutes
#     window_1=[0, 360],                        # First usage window (12:00 AM to 6:00 AM)
#     window_2=[1080, 1440],                    # Second usage window (6:00 PM to midnight)
#     random_var_w=0.3,                         # Variability of the windows in percentage
#     time_fraction_random_variability=0.2,     # Randomness in total usage time (between 0 and 1)
# )

# Add Mobile Chargers to the household
mobile_chargers = household.add_appliance(
    name="Mobile Chargers",                     # The name of the appliance
    number=4,                                   # Number of mobile chargers in the apartment
    power=0.005,                                # Power usage of each appliance in Kilowatt (kW) while in use
    num_windows=2,                              # Number of time windows throughout the day when this appliance is used
    func_time=180,                              # Total usage time of the appliance in minutes per day
    func_cycle=60,                              # Minimum usage time after a switch-on event in minutes
    window_1=[0, 360],                          # First usage window (12:00 AM to 6:00 AM)
    window_2=[1080, 1440],                      # Second usage window (6:00 PM to midnight)
    random_var_w=0.3,                           # Variability of the windows in percentage
    time_fraction_random_variability=0.2,       # Randomness in total usage time (between 0 and 1)
)

# Add a Coffee Maker to the household
coffee_maker = household.add_appliance(
    name="Coffee Maker",                        # The name of the appliance
    number=1,                                   # Number of coffee makers in the apartment
    power=1.2,                                  # Power usage of each appliance in Kilowatt (kW)
    num_windows=2,                              # Number of time windows throughout the day when this appliance is used
    func_time=20,                               # Total usage time of the appliance in minutes per day
    func_cycle=5,                               # Minimum usage time after a switch-on event in minutes
    window_1=[360, 420],                        # First usage window (6:00 AM to 7:00 AM)
    window_2=[900, 960],                        # Second usage window (3:00 PM to 4:00 PM)
    random_var_w=0.2,                           # Variability of the windows in percentage
    time_fraction_random_variability=0.1,       # Randomness in total usage time (between 0 and 1)
)

# Prompt the user for the number of days to simulate
num_days = int(input("Please indicate the number of days to be generated: "))
start_date = "2024-01-15"                                                               # The starting date for the simulation [should be adjusted for Summer/ Winter Week]
end_date = pd.to_datetime(start_date) + pd.Timedelta(days=num_days - 1)                 # Adjusted to ensure correct end date

# Create a use case for the simulation with the household user category
use_case = UseCase(
    users=[household],                                                                  # List of all user categories included in the simulation
    date_start=start_date,                                                              # Starting date of the simulation
    date_end=end_date.strftime("%Y-%m-%d"),                                             # Ending date of the simulation
)

# Generate the daily load profiles for the specified number of days
daily_load_profiles = use_case.generate_daily_load_profiles()

# Convert the daily load profiles to a pandas DataFrame
daily_load_profiles_df = pd.DataFrame(
    daily_load_profiles, columns=["Power Consumption (kW)"], index=use_case.datetimeindex
)

# Add the 'Timeseries' column using the index
daily_load_profiles_df.reset_index(inplace=True)
daily_load_profiles_df.rename(columns={"index": "Timeseries"}, inplace=True)


# Create a results folder with a timestamp
current_time = datetime.datetime.now()
folder_name = current_time.strftime("BFixLM_results_WINTER_%Y-%m-%d_%H-%M-%S")
os.makedirs(f"Results_Comfficientshare/{folder_name}", exist_ok=True)

# Save the generated load profiles to an Excel file
filename = current_time.strftime(f"BFixLM_{num_days}_days_%Y-%m-%d_%H-%M-%S.xlsx")
output_file_path = os.path.join('Results_Comfficientshare', folder_name, filename)
daily_load_profiles_df.to_excel(output_file_path, index=False)
print(f"Load profiles have been saved to: {output_file_path}")


# Plot the results: First graph - Power consumption over the simulation period in minutes
plt.figure(figsize=(10, 6))
plt.plot(daily_load_profiles_df["Timeseries"], daily_load_profiles_df["Power Consumption (kW)"])
plt.title(f"Power Consumption Over {num_days} Days")
plt.xlabel("Time (Minutes)")
plt.ylabel("Power (kW)")
plt.grid(True)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Power Consumption over {num_days} days.png")
plt.close()

# Plot the results: Second graph - Average power consumption over a single day in hours
# Add a "Time" column extracted from the "Timeseries" column
daily_load_profiles_df["Time"] = pd.to_datetime(daily_load_profiles_df["Timeseries"]).dt.time

# Group by time to calculate the average power consumption
average_profile = daily_load_profiles_df.groupby("Time")["Power Consumption (kW)"].mean()

plt.figure(figsize=(10, 6))
average_profile.plot(ax=plt.gca())
plt.title("Average Power Consumption Over a Single Day")
plt.xlabel("Time (Hours)")
plt.ylabel("Power (kW)")
plt.xticks(
    pd.date_range("00:00", "23:59", freq="2h").time,  # Evenly spaced time intervals every 2 hours
    rotation=45, ha="right"
)
plt.grid(True)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Average Power Consumption Over a Single Day.png")
plt.close()
