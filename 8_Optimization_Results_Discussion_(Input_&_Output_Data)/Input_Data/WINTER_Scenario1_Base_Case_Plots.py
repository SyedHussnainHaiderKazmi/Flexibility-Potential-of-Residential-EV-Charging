import pandas as pd
import matplotlib.pyplot as plt
import os
import datetime

# Define input Excel file name
input_file = "examples/WINTER_Scenario1_Base_Case_Plots.xlsx"

# Read the data from the Excel file
df = pd.read_excel(input_file)

# Convert 'Timeseries' column to datetime format if necessary
df['Timeseries'] = pd.to_datetime(df['Timeseries'])

# Create a results folder with a timestamp
current_time = datetime.datetime.now()
folder_name = current_time.strftime("WINTER_Scenario_1_BASE_CASE_Plots_%Y-%m-%d_%H-%M-%S")
os.makedirs(f"Results_Comfficientshare/{folder_name}", exist_ok=True)

# 1. Flexible Load Plot
plt.figure(figsize=(14, 8))
plt.plot(df['Timeseries'], df['P_flexible'], label='P_flexible Before Shifting', color='deepskyblue', linestyle='dashed', lw=2.5)
plt.xlabel('Winter Week Time', fontsize=14)
plt.ylabel('Flexible Load (kW)', fontsize=14)
plt.title('Flexible Load Before Shifting (Base Case)', fontsize=16)
plt.legend(fontsize=12)
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_1.png", dpi=300)
plt.close()

# 2. Power Profiles Plot
plt.figure(figsize=(14, 8))
plt.plot(df['Timeseries'], df['P_fixed'], label='Fixed Load (kW)', color='saddlebrown', lw=1.5)
plt.plot(df['Timeseries'], df['P_flexible'], label='P_flexible Before Shifting (kW)', color='deepskyblue', linestyle='dashed', lw=1.5)
plt.fill_between(df['Timeseries'], df['P_pv (kW)'], color='goldenrod', alpha=0.7, label='PV Generation (kW)')
plt.fill_between(df['Timeseries'], df['P_cars_total_Base (kW)'], color='darkgreen', alpha=1.0, label='Car Charging Total (kW)')
plt.xlabel('Winter Week Time', fontsize=14)
plt.ylabel('Power (kW)', fontsize=14)
plt.title('Power Profiles Over Time (Base Case)', fontsize=16)
plt.legend(fontsize=12, loc='upper left')
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_2.png", dpi=300, bbox_inches='tight')
plt.close()

# 3. Total Car Charging Power Plot
plt.figure(figsize=(14, 8))
plt.plot(df['Timeseries'], df['P_cars_total_Base (kW)'], label='Car Charging Total (kW)', color='royalblue', lw=2.5)
plt.xlabel('Winter Week Time', fontsize=14)
plt.ylabel('Power (kW)', fontsize=14)
plt.title('Total Car Charging Power (5 Cars) (Base Case)', fontsize=16)
plt.legend(fontsize=12)
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_3.png", dpi=300)
plt.close()

# 4. Total Power Demand Over Time without PV System Plot
plt.figure(figsize=(14, 8))
plt.plot(df['Timeseries'], df['P_total_Base_without PV (kW)'], label='Total Power Demand (Without PV)', color='darkred', lw=2.5)
plt.fill_between(df['Timeseries'], df['P_total_Base_without PV (kW)'], color='darkred', alpha=1.0)
plt.xlabel('Winter Week Time', fontsize=14)
plt.ylabel('Power (kW)', fontsize=14)
plt.title('Total Power Demand without PV System (Base Case)', fontsize=16)
plt.legend(fontsize=12)
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_4.png", dpi=300)
plt.close()

# 5. Total Power Demand Over Time with PV System Plot
plt.figure(figsize=(14, 8))
plt.plot(df['Timeseries'], df['P_total_Base_with PV (kW)'], label='Total Power Demand (With PV)', color='mediumblue', lw=2.5)
plt.fill_between(df['Timeseries'], df['P_total_Base_with PV (kW)'], color='mediumblue', alpha=1.0)
plt.xlabel('Winter Week Time', fontsize=14)
plt.ylabel('Power (kW)', fontsize=14)
plt.title('Total Power Demand with PV System (Base Case)', fontsize=16)
plt.legend(fontsize=12)
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_5.png", dpi=300)
plt.close()

print("Plots have been successfully saved!")