import openpyxl
import pypsa
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Marginal costs in EUR/MWh
marginal_costs = {
    "brown_Coal": 40,
    "hard_coal": 50,
    "gas": 60,
    "oil": 80,
    "other_non_res": 30,
    "hydro": 0,
    "biomass": 20,
    "offshore_wind": 0,
    "onshore_wind": 0,
    "solar": 0,
}

# Power plant capacities (nominal powers in MW) in each country
power_plant_p_nom = {
    "Germany": {
        "brown_Coal": 15190,
        "hard_coal": 16003,
        "gas": 36664,
        "oil": 4442,
        "other_non_res": 3171,
        "hydro": 6439,
        "biomass": 9057,
        "offshore_wind": 9215,
        "onshore_wind": 62670,
        "solar": 96095,
    }
}

# Load the data from Excel (demand sheet)
file_path = r'C:\Users\user\OneDrive\Desktop\electricity market design of Germany\Phase 3\generator_and_demand.xlsx'
df_demand = pd.read_excel(file_path, sheet_name='demand')

# Print the column names to check for any discrepancies
print("Columns in the demand sheet:", df_demand.columns)

# Strip any leading/trailing spaces from column names (if necessary)
df_demand.columns = df_demand.columns.str.strip()

# Get the number of rows in the 'demand' sheet (this will be the number of snapshots)
num_snapshots = len(df_demand)

# Set the snapshots based on the number of rows in Column A
network_snapshots = range(num_snapshots)
print(f"Number of snapshots: {num_snapshots}")

# Initialize the PyPSA network
country = "Germany"
network = pypsa.Network()

# Set up snapshots in PyPSA
network.set_snapshots(network_snapshots)

# Add the bus for Germany
network.add("Bus", country)

# Load capacity factors for offshore, onshore, and solar from columns C, D, and E
offshore_capacity_factors = df_demand['offshore capacity factor'].tolist()  # Column C
onshore_capacity_factors = df_demand['onshore capacity factor'].tolist()  # Column D
solar_capacity_factors = df_demand['solar capacity factor'].tolist()  # Column E

# Load demand data from Column B
load_data = df_demand['Load (MW)'].tolist()  # Column B for the load values in MW

# Add generators for each technology in Germany
for tech in power_plant_p_nom[country]:
    if tech == "offshore_wind":
        network.add(
            "Generator",
            f"{country} {tech}",
            bus=country,
            p_nom=power_plant_p_nom[country][tech],
            marginal_cost=marginal_costs[tech],
            p_max_pu=offshore_capacity_factors,  # Use offshore capacity factors
        )
    elif tech == "onshore_wind":
        network.add(
            "Generator",
            f"{country} {tech}",
            bus=country,
            p_nom=power_plant_p_nom[country][tech],
            marginal_cost=marginal_costs[tech],
            p_max_pu=onshore_capacity_factors,  # Use onshore capacity factors
        )
    elif tech == "solar":
        network.add(
            "Generator",
            f"{country} {tech}",
            bus=country,
            p_nom=power_plant_p_nom[country][tech],
            marginal_cost=marginal_costs[tech],
            p_max_pu=solar_capacity_factors,  # Use solar capacity factors
        )
    else:
        network.add(
            "Generator",
            f"{country} {tech}",
            bus=country,
            p_nom=power_plant_p_nom[country][tech],
            marginal_cost=marginal_costs[tech],
            p_max_pu=1,  # For other technologies, assume maximum capacity is fully utilized
        )

# Add the load, which varies over the snapshots, using the load data from Column B
network.add(
    "Load",
    f"{country} load",
    bus=country,
    p_set=load_data,  # Use the load data from the Excel sheet
)

# Optimize the network
network.optimize()

#------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------


#RES and Non-RES

# Prepare the data
snapshots = range(len(network.snapshots))  # Integer indices for snapshots
load_curve = network.loads_t.p_set.loc[:, f"{country} load"]

# Separate RES and Non-RES generation
res_generation = network.generators_t.p.loc[:, network.generators.loc[
    network.generators['type'].isin(['solar', 'offshore_wind', 'onshore_wind'])].index].sum(axis=1)
non_res_generation = network.generators_t.p.loc[:, network.generators.loc[
    ~network.generators['type'].isin(['solar', 'offshore_wind', 'onshore_wind'])].index].sum(axis=1)

# Plot: RES Generation vs Non-RES Generation
fig, ax = plt.subplots(figsize=(12, 6))

# Plot RES generation and Non-RES generation
ax.plot(snapshots, res_generation, label="RES Generation (Solar, Wind)", color="green")
ax.plot(snapshots, non_res_generation, label="Non-RES Generation (Coal, Gas, etc.)", color="red")

# Set labels and title
ax.set_xlabel("Snapshots")
ax.set_ylabel("Generation (MW)")
ax.set_title("RES Generation vs Non-RES Generation")
ax.legend()

# Show the plot
plt.grid()
plt.show()

#------------------------------------------------------------------------------





# Prepare the data
snapshots = range(len(network.snapshots))  # Integer indices for snapshots
load_curve = network.loads_t.p_set.loc[:, f"{country} load"]
total_generation = network.generators_t.p.sum(axis=1)
marginal_price = network.buses_t.marginal_price.loc[:, country]

# Plot 1: Demand and Marginal Price
fig1, ax1 = plt.subplots(figsize=(12, 6))

# Plot load curve (demand) on primary y-axis
ax1.plot(snapshots, load_curve, label="Load Curve", color="blue")
ax1.set_xlabel("Snapshots")
ax1.set_ylabel("Load (MW)", color="blue")
ax1.tick_params(axis="y", labelcolor="blue")
ax1.grid()

# Create a secondary y-axis for marginal price
ax2 = ax1.twinx()
ax2.plot(snapshots, marginal_price, label="Marginal Price", color="red")
ax2.set_ylabel("Marginal Price (EUR/MWh)", color="red")
ax2.tick_params(axis="y", labelcolor="red")

# Title and Legend
fig1.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
plt.title("Demand and Marginal Price")

# Plot 2: Demand and Supply
fig2, ax3 = plt.subplots(figsize=(12, 6))

# Plot load curve (demand) and total generation (supply)
ax3.plot(snapshots, load_curve, label="Load Curve", color="blue")
ax3.plot(snapshots, total_generation, label="Total Generation", color="green")
ax3.set_xlabel("Snapshots")
ax3.set_ylabel("Power (MW)")
ax3.tick_params(axis="y")
ax3.grid()

# Title and Legend
fig2.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
plt.title("Demand and Supply")

# Plot 3: Supply and Marginal Price
fig3, ax4 = plt.subplots(figsize=(12, 6))

# Plot total generation (supply)
ax4.plot(snapshots, total_generation, label="Total Generation", color="green")
ax4.set_xlabel("Snapshots")
ax4.set_ylabel("Total Generation (MW)", color="green")
ax4.tick_params(axis="y", labelcolor="green")
ax4.grid()

# Create a secondary y-axis for marginal price
ax5 = ax4.twinx()
ax5.plot(snapshots, marginal_price, label="Marginal Price", color="red")
ax5.set_ylabel("Marginal Price (EUR/MWh)", color="red")
ax5.tick_params(axis="y", labelcolor="red")

# Title and Legend
fig3.legend(loc="upper left", bbox_to_anchor=(0.1, 0.9))
plt.title("Supply and Marginal Price")

# Show all plots
plt.show()

#----------------------------------------------------------------------------

# Plot the load curve
plt.figure(figsize=(10, 6))
plt.plot(network.snapshots, network.loads_t.p_set.loc[:, f"{country} load"], label="Load Curve", color="blue")
plt.xlabel("Snapshots")
plt.ylabel("Load (MW)")
plt.title("Load Curve")
plt.legend()
plt.grid()
plt.show()

# Plot the marginal price of the bus
plt.figure(figsize=(10, 6))
plt.plot(network.snapshots, network.buses_t.marginal_price.loc[:, country], label="Marginal Price", color="red")
plt.xlabel("Snapshots")
plt.ylabel("Marginal Price (EUR/MWh)")
plt.title("Marginal Price at the Bus")
plt.legend()
plt.grid()
plt.show()
#----------------------------------------------------------------------------