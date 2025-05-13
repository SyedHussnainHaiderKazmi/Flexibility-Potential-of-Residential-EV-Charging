import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import datetime

# Define input Excel file name
input_file = 'examples/Section_6.3_Summer_Week_Plot_modified.xlsx'

# Read the data from the Excel file
df = pd.read_excel(input_file, parse_dates=['Timeseries'])

# Create a results folder with the current timestamp
current_time = datetime.datetime.now().strftime("Section_6.3_Summer_Plot_modified_%Y-%m-%d_%H-%M-%S")
output_folder = f"Results_Comfficientshare/{current_time}"
os.makedirs(output_folder, exist_ok=True)

# Plot settings
plt.figure(figsize=(14, 8))

# Define colors and labels for each column
plot_data = {
    'Peak_Power_Base_Case_(%)': ('Base Case', '#1f77b4'), #'#d62728'
    'Peak_Power_20 Percent Load Shift_(%)': ('20 Percent Load Shift', '#2ca02c'), #'#ff7f0e'
    'Peak_Power_30 Percent Load Shift_(%)': ('30 Percent Load Shift', '#8c564b'), #'#8c564b'
    'Peak_Power_40 Percent Load Shift_(%)': ('40 Percent Load Shift', '#ff7f0e'), #'#2ca02c'
    'Peak_Power_50 Percent Load Shift_(%)': ('50 Percent Load Shift', '#d62728')  #'#1f77b4'
}

# Plot each line (Removed markers for smooth curves)
for column, (label, color) in plot_data.items():
    plt.plot(df['Timeseries'], df[column], label=label, color=color, lw=2)

# Format X-axis
plt.xticks(rotation=30, fontsize=12)
xlabel = plt.xlabel('Summer Week Time', fontsize=18)
xlabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))  # Box around X-label

# Format Y-axis
ylabel = plt.ylabel('Peak Power Percentage (ref. to Base Case)', fontsize=18)
ylabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))  # Box around Y-label
plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x}%'))
plt.yticks(fontsize=12)

# Title and legend
plt.title('Change in Peak Power (Without PV System, 6 Hours Limit)', fontsize=18, fontweight='bold')
plt.legend(fontsize=12, frameon=True, edgecolor='black')

# Grid and layout
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.tight_layout()

# Save plot
output_path = f"{output_folder}/Plot_2_Summer_Week_Peak_Power.png"
plt.savefig(output_path, dpi=200, bbox_inches='tight')
plt.close()
