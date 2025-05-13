# ===========================================
# Section 1: Libraries and Environment Setup
# ===========================================

# Import necessary libraries
import pyomo.environ as pyo
from pyomo.opt import SolverFactory
import gurobipy as gp
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime
import os

# Ensure Gurobi solver is being used
solver = SolverFactory('gurobi')

# =====================================================
# Section 2: Data Import (Time Series Input from Excel)
# =====================================================

# Load the Excel file
file_path = 'examples/Comfficientshare_v9_Summer.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load 'Building_data' sheet
building_data = pd.read_excel(excel_data, 'Building_data')

# Parse Building Data
# -------------------
# Timeseries
building_timeseries = pd.to_datetime(building_data['Timeseries'])
# Fixed power consumption (kW)
P_fixed = building_data['P_fixed (kW)'].values
# Flexible power consumption (kW)
P_flexible = building_data['P_flexible (kW)'].values
# PV generation (kW)
P_pv = building_data['P_pv (kW)'].values
# Electricity cost (€/kWh)
C_t = building_data['C_t (€/kWh)'].values

# Load 'Cars_location' sheet
cars_location = pd.read_excel(excel_data, 'Cars_location')

# Parse Car Location Data
# -----------------------
# Timeseries
car_location_timeseries = pd.to_datetime(cars_location['Timeseries'])
# Car location status (binary values indicating car presence at home)
car_204E_location = cars_location['204E'].values
car_213E_location = cars_location['213E'].values
car_288E_location = cars_location['288E'].values
car_349E_location = cars_location['349E'].values
car_397E_location = cars_location['397E'].values

# Load 'Cars_trips_distance' sheet
cars_trips_distance = pd.read_excel(excel_data, 'Cars_trips_distance')

# Parse Car Trip Distance Data
# ----------------------------
# Timeseries
car_trip_timeseries = pd.to_datetime(cars_trips_distance['Timeseries'])
# Car trip distances (in km, values appear at the end of each trip)
car_204E_trip_distance = cars_trips_distance['204E'].fillna(0).values
car_213E_trip_distance = cars_trips_distance['213E'].fillna(0).values
car_288E_trip_distance = cars_trips_distance['288E'].fillna(0).values
car_349E_trip_distance = cars_trips_distance['349E'].fillna(0).values
car_397E_trip_distance = cars_trips_distance['397E'].fillna(0).values

# Check Data Integrity and Alignment
# -----------------------------------
# Ensure that timeseries data across sheets align for consistent indexing
assert (building_timeseries.equals(car_location_timeseries) and 
        building_timeseries.equals(car_trip_timeseries)), \
    "Timeseries data across sheets are misaligned!"

# Organize the imported data in a dictionary for easy reference within Pyomo
input_data = {
    'P_fixed': P_fixed,
    'P_flexible': P_flexible,
    'P_pv': P_pv,
    'C_t': C_t,
    'car_location': {
        '204E': car_204E_location,
        '213E': car_213E_location,
        '288E': car_288E_location,
        '349E': car_349E_location,
        '397E': car_397E_location
    },
    'car_trip_distance': {
        '204E': car_204E_trip_distance,
        '213E': car_213E_trip_distance,
        '288E': car_288E_trip_distance,
        '349E': car_349E_trip_distance,
        '397E': car_397E_trip_distance
    },
    'timeseries': building_timeseries
}

# =================================
# Section 3: Define Model and Sets
# =================================

# Define the optimization model
model = pyo.ConcreteModel()

# Set of time intervals T, corresponding to each 15-minute interval in the week
model.T = pyo.Set(initialize=range(len(input_data['P_fixed'])), ordered=True)

# Set of cars based on available car identifiers from the input data
car_ids = ['204E', '213E', '288E', '349E', '397E']
model.C = pyo.Set(initialize=car_ids, ordered=True)

# Define subsets for car-specific data at each time interval
# Binary sets for car location status, where '1' indicates the car is at home, '0' if away
# and distance set indicating trip distance at the end of each trip
model.car_location = pyo.Param(model.C, model.T, initialize={(car, t): input_data['car_location'][car][t] for car in car_ids for t in range(len(input_data['car_location'][car]))}, within=pyo.Binary)
model.car_distance = pyo.Param(model.C, model.T, initialize={(car, t): input_data['car_trip_distance'][car][t] if not pd.isnull(input_data['car_trip_distance'][car][t]) else 0 for car in car_ids for t in range(len(input_data['car_trip_distance'][car]))}, within=pyo.NonNegativeReals)

# =================================
# Section 4: Define Parameters
# =================================

# Define static parameters within the Pyomo model
model.SOC_min = pyo.Param(initialize=20)  # Minimum allowable SOC (Percentage)
model.SOC_max = pyo.Param(initialize=100) # Maximum allowable SOC (Percentage)

# Define parameters within the Pyomo model
model.P_fixed = pyo.Param(model.T, initialize={t: input_data['P_fixed'][t] for t in model.T}, within=pyo.NonNegativeReals)   # Fixed power consumption at each time interval (kW)

model.P_flexible = pyo.Param(model.T, initialize={t: input_data['P_flexible'][t] for t in model.T}, within=pyo.NonNegativeReals)   # Flexible power consumption at each time interval (kW)

model.P_pv = pyo.Param(model.T, initialize={t: input_data['P_pv'][t] for t in model.T}, within=pyo.NonNegativeReals)   # PV generation at each time interval (kW)

model.C_t = pyo.Param(model.T, initialize={t: input_data['C_t'][t] for t in model.T}, within=pyo.NonNegativeReals)   # Electricity cost at each time interval (€/kWh)

model.eta = pyo.Param(model.C, within=pyo.NonNegativeReals, initialize=0.95)   # Charging efficiency per Car (Percentage)

model.Max_Shifting_Capability = pyo.Param(model.T, within=pyo.NonNegativeReals, initialize=0.50)   # Max percentage of flexible load that can be shifted (Percentage)

model.Upper_Power_Limit = pyo.Param(model.T, within=pyo.NonNegativeReals, initialize=65)   # Upper power limit (kW)

model.Lower_Power_Limit = pyo.Param(model.T, within=pyo.NonNegativeReals, initialize=0)   # Lower power limit (kW)

model.Battery_Capacity = pyo.Param(model.C, within=pyo.NonNegativeReals, initialize=84)   # Battery capacity per Car (kWh) 

model.P_car_max = pyo.Param(model.C, within=pyo.NonNegativeReals, initialize=11)   # Max Charging Power per Car (11kW for CUPRA Born Charger)

model.Car_Mileage = pyo.Param(model.C, within=pyo.NonNegativeReals, initialize=6.28)   # Car mileage per Car (km/kWh)


# Target SOC, assuming all cars at 100% SOC
Target_SOC_value = 100 # Percentage
model.SOC_Target = pyo.Param(model.C, initialize={car: Target_SOC_value for car in model.C})


# =================================================
# Section 5: Define Variables (Decision Variables)
# =================================================

# Charging power for each car at each time interval (kW)
# The domain is non-negative as charging power cannot be negative
model.P_car_charge = pyo.Var(model.C, model.T, domain=pyo.NonNegativeReals)

# State of Charge (SOC) for each car at each time interval (%)
# Bounds are defined by SOC_min and SOC_max parameters
model.SOC = pyo.Var(model.C, model.T, bounds=(model.SOC_min, model.SOC_max))

# Shifted flexible load at each time interval (kW)
# The domain is non-negative as load shifting cannot reduce the overall load below zero
model.P_shift = pyo.Var(model.T, domain=pyo.NonNegativeReals)     # Non-negative for fractional flexible loads

# Modify Delta_P_shift to allow forward/backward range of ±24 intervals
# Allowing shifts up to ±24 intervals (6 hours)
model.Delta_P_shift = pyo.Var(model.T, domain=pyo.Integers, bounds=(-72, 72))     # Integer shifts within ±6 hours

# Binary variable y_shift[t, delta] that determines whether P_shift[t] is moved to interval t+delta
model.y_shift = pyo.Var(model.T, range(-72, 73), domain=pyo.Binary)

# Introduce a new binary variable z_shift[t] that is 1 if P_shift[t] > 0, and 0 otherwise
model.z_shift = pyo.Var(model.T, domain=pyo.Binary)


# ========================================================================
# Section 6: Define Objective Function: Minimize Overall Electricity Costs
# ========================================================================

# Define the objective function to minimize electricity costs
def objective_rule(model):
    return sum(
        model.C_t[t] * (model.P_fixed[t] + model.P_flexible[t] - model.P_shift[t] 
            + sum(model.P_shift[t - delta] * model.y_shift[t - delta, delta]  
                  for delta in range(-72, 73) if 0 <= t - delta < len(model.T))
            + sum(model.P_car_charge[c, t] for c in model.C) - model.P_pv[t]) for t in model.T)

# Add the objective to the model
model.objective = pyo.Objective(rule=objective_rule, sense=pyo.minimize)


# ==================================================
# Section 7: Define Constraints (All 5 Constraints)
# ==================================================

# ==========================================
# Section 7.1: Car Charging Power Constraint
# ==========================================

def car_charging_power_rule(model, c, t):
    # If the car is available for charging (binary 1 in Cars_location sheet)
    if model.car_location[c, t] == 1:
        return model.P_car_charge[c, t] <= model.P_car_max[c]
    # If the car is not available for charging (binary 0 in Cars_location sheet)
    else:
        return model.P_car_charge[c, t] == 0

# Add the constraint to the model
model.car_charging_power_constraint = pyo.Constraint(model.C, model.T, rule=car_charging_power_rule)

# ===========================================================================
# Section 7.1.1: No charging at the start of the week if SOC is already 100%
# ===========================================================================

def no_charging_initial_rule(model, c):
    # At the first time interval, if SOC starts at 100%, P_car_charge should be 0
    if model.SOC_Target[c] == 100:
        return model.P_car_charge[c, model.T.first()] == 0
    else:
        return pyo.Constraint.Skip

# Add the constraint to the model
model.no_charging_initial_constraint = pyo.Constraint(model.C, rule=no_charging_initial_rule)

# ================================
# Section 7.2: Car SOC Constraints
# ================================

# ===========================================================
# Car SOC Constraint 7.2.1: Initial SOC at Start of the Week
# ===========================================================

def initial_soc_rule(model, c):
    # At the first time interval, all cars start with 100% SOC
    return model.SOC[c, model.T.first()] == model.SOC_Target[c]

# Add the constraint to the model
model.initial_soc_constraint = pyo.Constraint(model.C, rule=initial_soc_rule)

# =======================================================
# Car SOC Constraint 7.2.2: SOC at Departure (SOC_target)
# =======================================================

def soc_target_rule(model, c, t):
    # Check if car switches from "at building" (1) to "not at building" (0)
    if t < model.T.last() and model.car_location[c, t] == 1 and model.car_location[c, t + 1] == 0:
        return model.SOC[c, t] >= model.SOC_Target[c]
    else:
        return pyo.Constraint.Skip

# Add the constraint to the model
model.soc_target_constraint = pyo.Constraint(model.C, model.T, rule=soc_target_rule)

# =======================================================
# Car SOC Constraints 7.2.3: SOC at Arrival (SOC_arrival)
# =======================================================

def soc_arrival_rule(model, c, t):
    # Check if car switches from "not at building" (0) to "at building" (1)
    if t < model.T.last() and model.car_location[c, t] == 0 and model.car_location[c, t + 1] == 1:
        # SOC at arrival based on trip distance; must remain above SOC_min
        return model.SOC[c, t + 1] == model.SOC_Target[c] - (model.car_distance[c, t] / model.Car_Mileage[c]) * (100 / model.Battery_Capacity[c])
    else:
        return pyo.Constraint.Skip

# Additional Constraint: Enforce SOC_min after arrival
def enforce_minimum_soc_rule(model, c, t):
    # Ensure SOC does not drop below SOC_min after arriving
    if t < model.T.last() and model.car_location[c, t] == 0 and model.car_location[c, t + 1] == 1:
        return model.SOC[c, t + 1] >= model.SOC_min
    else:
        return pyo.Constraint.Skip

# Add the constraints to the model
model.soc_arrival_constraint = pyo.Constraint(model.C, model.T, rule=soc_arrival_rule)
model.enforce_minimum_soc_constraint = pyo.Constraint(model.C, model.T, rule=enforce_minimum_soc_rule)

# ======================================================
# Car SOC Constraints 7.2.4: SOC During Charging Periods
# ======================================================

def soc_during_charging_rule(model, c, t):
    # Only during the time car is at the building for charging
    if t > model.T.first() and model.car_location[c, t] == 1 and model.car_location[c, t - 1] == 1:
        # SOC at time t is equal to SOC at (t-1) plus charging increment
        return model.SOC[c, t] == model.SOC[c, t - 1] + (model.P_car_charge[c, t] * model.eta[c] * 100 / model.Battery_Capacity[c])
    else:
        return pyo.Constraint.Skip

# Additional Constraint: Enforce SOC_max during charging
def enforce_maximum_soc_rule(model, c, t):
    # Ensure SOC does not exceed SOC_max during charging
    if t > model.T.first() and model.car_location[c, t] == 1:
        return model.SOC[c, t] <= model.SOC_max
    else:
        return pyo.Constraint.Skip

# Add the constraints to the model
model.soc_during_charging_constraint = pyo.Constraint(model.C, model.T, rule=soc_during_charging_rule)
model.enforce_maximum_soc_constraint = pyo.Constraint(model.C, model.T, rule=enforce_maximum_soc_rule)

# ======================================================
# Car SOC Constraint 7.2.5: Final SOC at End of the Week
# ======================================================

def final_soc_rule(model, c):
    # At the last time interval, ensure the car reaches 100% SOC if it is at home
    if model.car_location[c, model.T.last()] == 1:
        return model.SOC[c, model.T.last()] >= model.SOC_Target[c]
    else:
        return pyo.Constraint.Skip

# Add the constraint to the model
model.final_soc_constraint = pyo.Constraint(model.C, rule=final_soc_rule)

# ===============================================
# Section 7.3: Upper Connection Limit Constraint
# ===============================================

def upper_power_limit_rule(model, t):
    return (model.P_fixed[t] + model.P_flexible[t] - model.P_shift[t]  
        + sum(model.P_shift[t - delta] * model.y_shift[t - delta, delta] for delta in range(-72, 73) if 0 <= t - delta < len(model.T))  # Add received flexible load
        + sum(model.P_car_charge[c, t] for c in model.C) <= model.Upper_Power_Limit[t])

model.power_limit_upper = pyo.Constraint(model.T, rule=upper_power_limit_rule)

# ==============================================
# Section 7.4: Lower Connection Limit Constraint
# ==============================================

def lower_power_limit_rule(model, t):
    return (model.Lower_Power_Limit[t] <= model.P_fixed[t] + model.P_flexible[t] - model.P_shift[t]  
        + sum(model.P_shift[t - delta] * model.y_shift[t - delta, delta] for delta in range(-72, 73) if 0 <= t - delta < len(model.T))  # Add received flexible load
        + sum(model.P_car_charge[c, t] for c in model.C))

model.power_limit_lower = pyo.Constraint(model.T, rule=lower_power_limit_rule)


# ================================================
# Section 7.5: Flexible Load Shifting Constraints
# ================================================

# ================================================================
# Flexible Load 7.5.1: Restrict P_shift to Available Flexible Load
# ================================================================
# Constraint: Ensure P_shift[t] is between 0 and 20% of P_flexible[t]

# Lower bound: P_shift[t] must be at least 0 kW (shifting is optional)
def flexible_load_limit_lower_rule(model, t):
    return model.P_shift[t] >= 0  # Allows no shifting when it's not needed

model.flexible_load_limit_lower = pyo.Constraint(model.T, rule=flexible_load_limit_lower_rule)

# Upper bound: P_shift[t] cannot exceed 20% of P_flexible[t]
def flexible_load_limit_upper_rule(model, t):
    return model.P_shift[t] <= model.P_flexible[t] * model.Max_Shifting_Capability[t]

model.flexible_load_limit_upper = pyo.Constraint(model.T, rule=flexible_load_limit_upper_rule)

# =========================================================
# Flexible Load 7.5.2: Assign Binary Variables for Shifting
# =========================================================

# This ensures that each P_shift[t] is assigned to exactly one target interval.
# Constraint to ensure that P_shift[t] is assigned to exactly one time interval
def shift_assignment_rule(model, t):
    return sum(model.y_shift[t, delta] for delta in range(-72, 73) if 0 <= t + delta < len(model.T)) == 1

model.shift_assignment = pyo.Constraint(model.T, rule=shift_assignment_rule)

# ================================================================
# Flexible Load 7.5.3: Enforce Load Balance at the Target Interval
# ================================================================

# This ensures that no time interval gets overloaded beyond 120% of its original flexible load.
# Constraint to limit the final flexible load after shifting
def shifted_load_limit_rule(model, t, delta):
    if 0 <= t + delta < len(model.T):
        return model.P_flexible[t + delta] + model.P_shift[t] * model.y_shift[t, delta] <= 1.5 * model.P_flexible[t + delta]
    else:
        return pyo.Constraint.Skip

model.shifted_load_limit = pyo.Constraint(model.T, range(-72, 73), rule=shifted_load_limit_rule)

# ===========================================
# Flexible Load 7.5.4: Enforce Load Shifting
# ===========================================

# This ensures that P_shift[t] moves to exactly one location and does not disappear.
# Constraint to ensure the shifted load is assigned properly
def enforce_shift_rule(model, t):
    return model.P_shift[t] == sum(model.P_shift[t] * model.y_shift[t, delta] for delta in range(-72, 73) if 0 <= t + delta < len(model.T))

model.enforce_shift = pyo.Constraint(model.T, rule=enforce_shift_rule)


# ==========================================================================
# Flexible Load 7.5.5: Add a Constraint to Link z_shift[t] to P_shift[t]
# ==========================================================================

def link_z_shift_rule(model, t):
    return model.P_shift[t] <= model.z_shift[t] * 100  # Big-M method (assuming max P_shift is 100 kW)

model.link_z_shift = pyo.Constraint(model.T, rule=link_z_shift_rule)

# =======================================================================
# Flexible Load 7.5.6: Define Delta_P_shift[t] based on y_shift[t, delta]
# =======================================================================

# This ensures that Delta_P_shift represents the actual shift interval.
# Ensure that Delta_P_shift[t] is 0 when no load is shifted.
def delta_p_shift_definition_rule(model, t):
    return model.Delta_P_shift[t] == sum(delta * model.y_shift[t, delta] for delta in range(-72, 73)) * model.z_shift[t]

model.delta_p_shift_definition = pyo.Constraint(model.T, rule=delta_p_shift_definition_rule)

# ============================================================================
# Flexible Load 7.5.7: Prevent y_shift[t, 0] from being 1 when P_shift[t] > 0
# ============================================================================

# If P_shift[t] > 0, at least one nonzero shift interval must be chosen.
def prevent_zero_shift_rule(model, t):
    return model.y_shift[t, 0] <= 1 - (model.P_shift[t] / (model.P_flexible[t] * model.Max_Shifting_Capability[t] + 1e-6))

model.prevent_zero_shift = pyo.Constraint(model.T, rule=prevent_zero_shift_rule)


# =======================================
# Section 8: Solver Setup (Using Gurobi)
# =======================================

# Solve the optimization problem using Gurobi
results = solver.solve(model, tee=True)
model.write('model.lp', io_options={'symbolic_solver_labels': True})

# Check the solver status
if results.solver.termination_condition == pyo.TerminationCondition.optimal:
    print('Solution is optimal and feasible!')
else:
    print('1_Solver could not find an optimal solution.')

if results.solver.termination_condition != 'optimal':
   print("2_Solver could not find an optimal solution.")
   iis_prob = gp.read('model.lp')
   iis_prob.computeIIS()
   iis_prob.write('model_iis.ilp')

# ======================================================
# Section 9: Generate and Write Output (Including Plots)
# ======================================================

# Check if the optimization problem was solved to optimality
if results.solver.termination_condition == 'optimal':
    print("Optimization completed successfully. Extracting results...")

    # Create a results folder with a timestamp
    current_time = datetime.datetime.now()
    folder_name = current_time.strftime("OptimizationResults_SUMMER_50PShift_18HLimit_%Y-%m-%d_%H-%M-%S")
    os.makedirs(f"Results_Comfficientshare/{folder_name}", exist_ok=True)

    # Initialize dictionaries to store results
    results_dict = {
        'Timeseries': building_timeseries,
        'P_fixed': [P_fixed[t] for t in model.T],
        'P_flexible': [P_flexible[t] for t in model.T],
        'P_shift': [model.P_shift[t]() for t in model.T],
        'Delta_P_shift': [model.Delta_P_shift[t]() for t in model.T],
        'P_flexible_post_shift': [P_flexible[t] - model.P_shift[t]() +  
            sum(model.P_shift[t - delta]() * model.y_shift[t - delta, delta]() for delta in range(-72, 73) if 0 <= t - delta < len(model.T))
            for t in model.T],
        'P_pv': [P_pv[t] for t in model.T],
        'P_total': [P_fixed[t] + P_flexible[t] - model.P_shift[t]() + 
            sum(model.P_shift[t - delta]() * model.y_shift[t - delta, delta]() for delta in range(-72, 73) if 0 <= t - delta < len(model.T)) +
            sum(model.P_car_charge[c, t]() for c in model.C) - P_pv[t]
            for t in model.T],
        'P_total_noPV': [P_fixed[t] + P_flexible[t] - model.P_shift[t]() + 
            sum(model.P_shift[t - delta]() * model.y_shift[t - delta, delta]() for delta in range(-72, 73) if 0 <= t - delta < len(model.T)) +
            sum(model.P_car_charge[c, t]() for c in model.C) for t in model.T]
    }

    # Add SOC results and individual car charging power to the results_dict
    for car in model.C:
        results_dict[f'SOC_{car}'] = [model.SOC[car, t]() for t in model.T]
        results_dict[f'P_car_charge_{car}'] = [model.P_car_charge[car, t]() for t in model.T]

    # Add collective car charging power over all cars
    results_dict['P_cars_total'] = [sum(model.P_car_charge[c, t]() for c in model.C) for t in model.T]

    # Compute total electricity costs
    total_electricity_cost_PV = sum(results_dict['P_total'][t] * C_t[t] for t in range(len(model.T)))  # With PV
    total_electricity_cost_noPV = sum(results_dict['P_total_noPV'][t] * C_t[t] for t in range(len(model.T)))  # Without PV

    # Print the values in the VS Code terminal
    print(f"Total electricity cost (With PV): €{total_electricity_cost_PV:.2f}")
    print(f"Total electricity cost (Without PV): €{total_electricity_cost_noPV:.2f}")

    # Add the total electricity cost values ONLY in the first row of the last two columns
    results_dict['Total_Electricity_Cost_PV'] = [total_electricity_cost_PV] + [None] * (len(model.T) - 1)
    results_dict['Total_Electricity_Cost_noPV'] = [total_electricity_cost_noPV] + [None] * (len(model.T) - 1)
    
    # Compute the total shifted load and the global limit
    total_shifted_load = sum(model.P_shift[t].value for t in model.T)
    global_shifted_load_limit = sum(model.P_flexible[t] * model.Max_Shifting_Capability[t] for t in model.T)

    # Print the values in the VS Code terminal
    print(f"Total shifted load: {total_shifted_load}")
    print(f"Global shifting limit: {global_shifted_load_limit}")

    # Add the values in the first row of the last two columns
    results_dict['Total_Shifted_Load'] = [total_shifted_load] + [None] * (len(model.T) - 1)
    results_dict['Global_Shifting_Limit'] = [global_shifted_load_limit] + [None] * (len(model.T) - 1)


    # Create a DataFrame for numerical results
    results_df = pd.DataFrame(results_dict)

    # Save results to an Excel file
    file_name = current_time.strftime("OptimizationResults_SUMMER_50PShift_18HLimit_%Y-%m-%d_%H-%M-%S.xlsx")
    output_file_path = os.path.join('Results_Comfficientshare', folder_name, file_name)
    results_df.to_excel(output_file_path, index=False)
    print(f"Results saved to Excel file at: {output_file_path}")


    # Visualization: Generate and save plots
    # 1. Plot Total Power Demand Over Time with PV System
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    plt.figure(figsize=(14, 8))
    plt.plot(results_df['Timeseries'], results_df['P_total'], label='Total Power Demand (With PV)', color='mediumblue', lw=2.5)
    plt.fill_between(results_df['Timeseries'], results_df['P_total'], color='mediumblue', alpha=1.0)
    plt.xlabel('Summer Week Time', fontsize=14)
    plt.ylabel('Power (kW)', fontsize=14)
    plt.title(f'Total Power Demand with PV System ({scenario_label})', fontsize=16)
    plt.legend(fontsize=12)
    plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/Total_Power_Demand_with_PV_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
    plt.close()

    # 2. Plot Total Power Demand Over Time without PV System
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    plt.figure(figsize=(14, 8))
    plt.plot(results_df['Timeseries'], results_df['P_total_noPV'], label='Total Power Demand (Without PV)', color='darkred', lw=2.5)
    plt.fill_between(results_df['Timeseries'], results_df['P_total_noPV'], color='darkred', alpha=1.0)
    plt.xlabel('Summer Week Time', fontsize=14)
    plt.ylabel('Power (kW)', fontsize=14)
    plt.title(f'Total Power Demand without PV System ({scenario_label})', fontsize=16)
    plt.legend(fontsize=12)
    plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/Total_Power_Demand_without_PV_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
    plt.close()


    # 3. Plot SOC evolution for each car
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    for car in model.C:
        plt.figure(figsize=(14, 8))
        plt.plot(results_df['Timeseries'], results_df[f'SOC_{car}'], label=f'SOC of Car {car}', color='forestgreen', lw=2.5)
        plt.xlabel('Summer Week Time', fontsize=14)
        plt.ylabel('State of Charge (%)', fontsize=14)
        plt.title(f'SOC Evolution for Car {car} ({scenario_label})', fontsize=16)
        plt.legend(fontsize=12)
        plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
        plt.savefig(f"Results_Comfficientshare/{folder_name}/SOC_Car_{car}_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
        plt.close()
    
    # 4. Power Profiles: Fixed, Flexible After Shifting, PV, and Total Car Charging Power (all cars)
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    plt.figure(figsize=(14, 8))
    plt.plot(results_df['Timeseries'], results_df['P_fixed'], label='Fixed Load (kW)', color='saddlebrown', lw=1.5)
    plt.plot(results_df['Timeseries'], results_df['P_flexible_post_shift'], label='Flexible Load After Shifting (kW)', color='blue', lw=1.5)
    plt.fill_between(results_df['Timeseries'], results_df['P_pv'], color='goldenrod', alpha=0.7, label='PV Generation (kW)')
    plt.fill_between(results_df['Timeseries'], results_df['P_cars_total'], color='darkgreen', alpha=1.0, label='Car Charging Total (kW)')
    plt.xlabel('Summer Week Time', fontsize=14)
    plt.ylabel('Power (kW)', fontsize=14)
    plt.title(f'Power Profiles Over Time ({scenario_label})', fontsize=16)
    plt.legend(fontsize=12, loc='upper right')
    plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/Power_Profiles_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300, bbox_inches='tight')
    plt.close()
    
    # 5. Plot Individual Car Charging Power Profiles
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    for car in model.C:
        plt.figure(figsize=(14, 8))
        plt.plot(results_df['Timeseries'], results_df[f'P_car_charge_{car}'], label=f'Charging Power of Car {car} (kW)', color='dodgerblue', lw=2.5)
        plt.xlabel('Summer Week Time', fontsize=14)
        plt.ylabel('Power (kW)', fontsize=14)
        plt.title(f'Charging Power Profile for Car {car} ({scenario_label})', fontsize=16)
        plt.legend(fontsize=12)
        plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
        plt.savefig(f"Results_Comfficientshare/{folder_name}/Charging_Power_Car_{car}_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
        plt.close()
    
    # 6. Total Car Charging Power (5 cars)
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    plt.figure(figsize=(14, 8))
    plt.plot(results_df['Timeseries'], results_df['P_cars_total'], label='Car Charging Total (kW)', color='royalblue', lw=2.5)
    plt.xlabel('Summer Week Time', fontsize=14)
    plt.ylabel('Power (kW)', fontsize=14)
    plt.title(f'Total Car Charging Power (5 Cars) ({scenario_label})', fontsize=16)
    plt.legend(fontsize=12)
    plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/5_Cars_Charging_Power_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
    plt.close()
    
    # 7. Flexible Load Shifts for DSM
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    plt.figure(figsize=(14, 8))
    plt.plot(results_df['Timeseries'], results_df['P_shift'], label='Shifted Flexible Load (kW)', color='olivedrab', lw=2.5)
    plt.xlabel('Summer Week Time', fontsize=14)
    plt.ylabel('Power (kW)', fontsize=14)
    plt.title(f'Flexible Load Shifting Over Time ({scenario_label})', fontsize=16)
    plt.legend(fontsize=12)
    plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/Flexible_Load_Shifting_Plot_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
    plt.close()

    # 8. Comparison of Flexible Load Before and After Shifting
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    plt.figure(figsize=(14, 8))
    plt.plot(results_df['Timeseries'], results_df['P_flexible'], label='P_flexible Before Shifting', color='deepskyblue', linestyle='dashed', lw=2.5)
    plt.plot(results_df['Timeseries'], results_df['P_flexible_post_shift'], label='P_flexible After Shifting', color='blue', lw=2.5)
    plt.xlabel('Summer Week Time', fontsize=14)
    plt.ylabel('Flexible Load (kW)', fontsize=14)
    plt.title(f'Comparison of Flexible Load Before and After Shifting ({scenario_label})', fontsize=16)
    plt.legend(fontsize=12)
    plt.grid(linestyle='--', linewidth=0.7, alpha=0.5)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/Flexible_Load_Comparison_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
    plt.close()
    
    
    # 9. Flexible_Load_Shifting_Histogram
    # Get the values of Delta_P_shift (number of hours shifted) and P_shift (shifted load)
    # Ensure valid data (replace None values with 0)
    scenario_label = "50% Shift, 18H Limit"  # Update this based on the scenario
    delta_p_shift_values = results_df['Delta_P_shift'].dropna().astype(int) / 4  # Remove NaN and convert to integers [Convert 15-min intervals to hours]
    p_shift_values = results_df['P_shift'].dropna().astype(float)  # Remove NaN and convert to floats

    # Ensure both arrays have the same length (to avoid mismatched weights)
    if len(delta_p_shift_values) != len(p_shift_values):
        raise ValueError("Mismatch: Delta_P_shift and P_shift must have the same length!")

    # Define bin range dynamically based on actual data
    bin_min = delta_p_shift_values.min()
    bin_max = delta_p_shift_values.max()
    bins = np.arange(bin_min - 0.5, bin_max + 1, 1)  # Adjust bins to center correctly

    # Create the histogram weighted by P_shift (shifted load)
    hist, bin_edges = np.histogram(delta_p_shift_values, bins=bins, weights=p_shift_values)

    # Plot the histogram
    plt.figure(figsize=(14, 8))
    plt.bar(bin_edges[:-1], hist, width=1, align='edge', color='olivedrab', edgecolor='black', linewidth=1.5)
    plt.xticks(np.arange(bin_min, bin_max + 1, 2))  # Ensure proper spacing for hours
    plt.xlabel('Shifted Time (Hours)', fontsize=14)
    plt.ylabel('Total Shifted Load (kW)', fontsize=14)
    plt.title(f'Histogram of Flexible Load Shifting ({scenario_label})', fontsize=16)
    plt.grid(axis='y', linestyle='--', linewidth=0.7, alpha=0.5)
    plt.legend(['Total Shifted Load'], fontsize=12)
    plt.savefig(f"Results_Comfficientshare/{folder_name}/Flexible_Load_Shifting_Histogram_{scenario_label.replace('%', 'P').replace(', ', '_')}.png", dpi=300)
    plt.close()



    print(f"All plots saved in the Results_Comfficientshare folder: Results_Comfficientshare/{folder_name}")
    


else:
    print("Optimization was not successful. Please check the model for infeasibility or errors.")

