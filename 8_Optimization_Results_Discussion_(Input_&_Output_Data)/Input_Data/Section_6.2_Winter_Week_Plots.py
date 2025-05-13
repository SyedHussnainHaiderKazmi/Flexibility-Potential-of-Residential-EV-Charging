import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import datetime
import re

# Define input Excel file name
input_file = 'examples/Section_6.2_Winter_Week_Plots.xlsx'

# Read the data from the Excel file
df = pd.read_excel(input_file)

df['Cost Reduction Percentage with PV (with reference to Base Case Scenario)'] = pd.to_numeric(
    df['Cost Reduction Percentage with PV (with reference to Base Case Scenario)'], errors='coerce')

df['Cost Reduction Percentage without PV (with reference to Base Case Scenario)'] = pd.to_numeric(
    df['Cost Reduction Percentage without PV (with reference to Base Case Scenario)'], errors='coerce')

# Convert 'Amount of Hours Shifted (50% Load Shifted)' to string to avoid plotting errors
df['Amount of Hours Shifted (50% Load Shifted)'] = df['Amount of Hours Shifted (50% Load Shifted)'].astype(str).fillna("")

def extract_percentage(value):
    if value.strip() == "Base Case":
        return 0  # Assign zero to 'Base Case'
    if 'Hours' in value:
        return float(value.split()[0]) / 24  # Convert hours to percentage (out of 24)
    return 0  # If Base Case, return 0 as default

df['Numeric Amount of Hours Shifted'] = df['Amount of Hours Shifted (50% Load Shifted)'].apply(extract_percentage)

# Create a results folder
current_time = datetime.datetime.now()
folder_name = current_time.strftime("Section_6.2_Winter_Plots_%Y-%m-%d_%H-%M-%S")
os.makedirs(f"Results_Comfficientshare/{folder_name}", exist_ok=True)

# Plot with PV
plt.figure(figsize=(14, 8))

# Plotting the data
plt.plot(df['Numeric Amount of Hours Shifted'], 
         df['Cost Reduction Percentage with PV (with reference to Base Case Scenario)'],  
         marker='o', 
         color='#2ca02c',  # Professional green color
         lw=2.5, 
         markersize=8, 
         markeredgewidth=2, 
         markeredgecolor='#003366',  
         label='Cost Reduction with PV')

# Create formatted labels
xticks_labels = df['Amount of Hours Shifted (50% Load Shifted)']

# Set x-ticks with proper font size
plt.xticks(df['Numeric Amount of Hours Shifted'], xticks_labels, fontsize=15)
plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f'{x}%'))
plt.yticks(fontsize=15)  

# Set x and y labels
xlabel = plt.xlabel('Amount of Hours Shifted (50% Load Shift)', fontsize=20)
ylabel = plt.ylabel('Cost Reduction Percentage (ref. to Base Case)', fontsize=20)

# Add frames around xlabel and ylabel
xlabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))
ylabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))

plt.title('Cost Reduction with PV System (Winter Week Scenarios)', fontsize=22, fontweight='bold')  
plt.legend(fontsize=18, loc='best', frameon=True, framealpha=0.7, edgecolor='black')  
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.tight_layout()
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_1_Winter_Week_PV.png", dpi=200, bbox_inches='tight')
plt.close()

# Plot without PV
plt.figure(figsize=(14, 8))

# Plotting the data
plt.plot(df['Numeric Amount of Hours Shifted'], 
         df['Cost Reduction Percentage without PV (with reference to Base Case Scenario)'],  
         marker='o', 
         color='#2ca02c',  # Professional green color
         lw=2.5, 
         markersize=8, 
         markeredgewidth=2, 
         markeredgecolor='#003366',  
         label='Cost Reduction without PV')

# Set x-ticks with proper font size
plt.xticks(df['Numeric Amount of Hours Shifted'], xticks_labels, fontsize=15)
plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f'{x}%'))
plt.yticks(fontsize=15)  

# Set x and y labels
xlabel = plt.xlabel('Amount of Hours Shifted (50% Load Shift)', fontsize=20)
ylabel = plt.ylabel('Cost Reduction Percentage (ref. to Base Case)', fontsize=20)

# Add frames around xlabel and ylabel
xlabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))
ylabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))

plt.title('Cost Reduction without PV System (Winter Week Scenarios)', fontsize=22, fontweight='bold')  
plt.legend(fontsize=18, loc='best', frameon=True, framealpha=0.7, edgecolor='black')  
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.tight_layout()
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_2_Winter_Week_without_PV.png", dpi=200, bbox_inches='tight')
plt.close()
