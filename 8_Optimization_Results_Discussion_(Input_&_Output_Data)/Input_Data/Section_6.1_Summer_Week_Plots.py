import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import datetime
import re

# Define input Excel file name
input_file = 'examples/Section_6.1_Summer_Week_Plots.xlsx'

# Read the data from the Excel file
df = pd.read_excel(input_file)


df['Cost Reduction Percentage with PV (with reference to Base Case Scenario)'] = pd.to_numeric(
    df['Cost Reduction Percentage with PV (with reference to Base Case Scenario)'], errors='coerce')

df['Cost Reduction Percentage without PV (with reference to Base Case Scenario)'] = pd.to_numeric(
    df['Cost Reduction Percentage without PV (with reference to Base Case Scenario)'], errors='coerce')

# Convert 'Percentage of Flexible Load Shifted (6H Limit)' to string to avoid plotting errors
df['Percentage of Flexible Load Shifted (6H Limit)'] = df['Percentage of Flexible Load Shifted (6H Limit)'].astype(str).fillna("")

def extract_percentage(value):
    if value.strip() == "Base Case":
        return 0  # Assign zero to 'Base Case'
    return float(value.replace('%', '').strip())  # Remove '%' and convert to float

df['Numeric Percentage of Flexible Load Shifted'] = df['Percentage of Flexible Load Shifted (6H Limit)'].apply(extract_percentage)

# Create a results folder
current_time = datetime.datetime.now()
folder_name = current_time.strftime("Section_6.1_Summer_Plots_%Y-%m-%d_%H-%M-%S")
os.makedirs(f"Results_Comfficientshare/{folder_name}", exist_ok=True)

# Plot with PV
plt.figure(figsize=(14, 8))

# Plotting the data
plt.plot(df['Numeric Percentage of Flexible Load Shifted'], 
         df['Cost Reduction Percentage with PV (with reference to Base Case Scenario)'],  
         marker='o', 
         color='#1f77b4',  
         lw=2.5, 
         markersize=8, 
         markeredgewidth=2, 
         markeredgecolor='#003366',  
         label='Cost Reduction with PV')

# Create formatted labels
xticks_labels = [f"{int(val*100)}%" if val != 0 else "Base Case" for val in df['Numeric Percentage of Flexible Load Shifted']]


# Set x-ticks with proper font size
plt.xticks(df['Numeric Percentage of Flexible Load Shifted'], xticks_labels, fontsize=15)
plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f'{x}%'))
plt.yticks(fontsize=15)  

plt.xlabel('Percentage of Flexible Load Shifted (6 Hours Limit)', fontsize=20)
plt.ylabel('Cost Reduction Percentage (ref. to Base Case)', fontsize=20)
xlabel = plt.xlabel('Percentage of Flexible Load Shifted (6 Hours Limit)', fontsize=20)
ylabel = plt.ylabel('Cost Reduction Percentage (ref. to Base Case)', fontsize=20)
xlabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))
ylabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))
plt.title('Cost Reduction with PV System (Summer Week Scenarios)', fontsize=22, fontweight='bold')  
plt.legend(fontsize=18, loc='best', frameon=True, framealpha=0.7, edgecolor='black')  
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.tight_layout()
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_1_Summer_Week_PV.png", dpi=200, bbox_inches='tight')
plt.close()


# Plot without PV
plt.figure(figsize=(14, 8))

# Plotting the data
plt.plot(df['Numeric Percentage of Flexible Load Shifted'], 
         df['Cost Reduction Percentage without PV (with reference to Base Case Scenario)'],  
         marker='o', 
         color='#1f77b4',  
         lw=2.5, 
         markersize=8, 
         markeredgewidth=2, 
         markeredgecolor='#003366',  
         label='Cost Reduction without PV')

# Create formatted labels
xticks_labels = [f"{int(val*100)}%" if val != 0 else "Base Case" for val in df['Numeric Percentage of Flexible Load Shifted']]

# Set x-ticks with proper font size
plt.xticks(df['Numeric Percentage of Flexible Load Shifted'], xticks_labels, fontsize=15)
plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: f'{x}%'))
plt.yticks(fontsize=15)  

plt.xlabel('Percentage of Flexible Load Shifted (6 Hours Limit)', fontsize=20)  
plt.ylabel('Cost Reduction Percentage (ref. to Base Case)', fontsize=20)
xlabel = plt.xlabel('Percentage of Flexible Load Shifted (6 Hours Limit)', fontsize=20)
ylabel = plt.ylabel('Cost Reduction Percentage (ref. to Base Case)', fontsize=20)
xlabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))
ylabel.set_bbox(dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.1'))
plt.title('Cost Reduction without PV System (Summer Week Scenarios)', fontsize=22, fontweight='bold')  
plt.legend(fontsize=18, loc='best', frameon=True, framealpha=0.7, edgecolor='black')  
plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
plt.tight_layout()
plt.savefig(f"Results_Comfficientshare/{folder_name}/Plot_2_Summer_Week_without_PV.png", dpi=200, bbox_inches='tight')
plt.close()
