import pandas as pd
import matplotlib.pyplot as plt

#Import datasets
base_path = "/Users/luorui/Desktop/UCL/COMP0035 Software Engineering/Coursework/"
file_path = base_path + "dataset.xlsx"

housing_df = pd.read_excel(file_path, sheet_name="housing", skiprows=6, header=[0,1])

# Reading the unemployment data from its sheet
unemployment_df = pd.read_excel(file_path, sheet_name="unemployment", skiprows=6)

# Initial understanding of the datasets
print("\nSummary for Housing Data:\n")
print(housing_df.head())
print("\nData Types:\n", housing_df.dtypes)
print("\nMissing Values:\n", housing_df.isnull().sum())
print("\nDescriptive Statistics:\n", housing_df.describe())

print("\nSummary for Unemployment Data:\n")
print(unemployment_df.head())
print("\nData Types:\n", unemployment_df.dtypes)
print("\nMissing Values:\n", unemployment_df.isnull().sum())
print("\nDescriptive Statistics:\n", unemployment_df.describe())

# Convert "Quarter" column in the unemployment dataset to the last month of the quarter
def convert_to_last_month(quarter_str):
    if isinstance(quarter_str, str) and "-" in quarter_str:
        return quarter_str.split("-")[-1].strip()
    return quarter_str

unemployment_df["Quarter"] = unemployment_df["Quarter"].apply(convert_to_last_month)

# Convert to datetime format
unemployment_df["Quarter"] = pd.to_datetime(unemployment_df["Quarter"], format="%b %Y", errors='coerce')

# Filter the unemployment dataset to cover the same time range as the housing dataset
unemployment_df_filtered = unemployment_df[(unemployment_df["Quarter"] >= "1995-03-01") & (unemployment_df["Quarter"] <= "2023-06-01")]

# Drop rows with missing data
unemployment_df_filtered = unemployment_df_filtered.dropna()

# Aggregate housing data to get the average price for each quarter
housing_df_quarterly = housing_df.resample('Q', on=('Month', 'Unnamed: 0_level_1')).mean()

# Visualization for Housing Prices and Annual Growth Rate in London
fig, ax1 = plt.subplots(figsize=(14, 7))
ax1.plot(housing_df_quarterly.index, housing_df_quarterly[('London', 'Value')], color="b", label="Housing Prices (London)")
ax1.set_xlabel('Quarter')
ax1.set_ylabel('Housing Prices', color="b")
ax1.tick_params(axis='y', labelcolor="b")
ax2 = ax1.twinx()
ax2.plot(housing_df_quarterly.index, housing_df_quarterly[('London', 'Annual growth')], color="g", linestyle='--', label="Annual Growth Rate (London)")
ax2.set_ylabel('Growth Rate (%)', color="g")
ax2.tick_params(axis='y', labelcolor="g")
plt.title("London's Housing Prices and Annual Growth Rate (1995-2023)")
plt.show()

# Visualization for Unemployment Rate in London
fig, ax2 = plt.subplots(figsize=(14, 7))
ax2.plot(unemployment_df_filtered["Quarter"], unemployment_df_filtered["London"], color="r", label="Unemployment Rate (London)")
ax2.set_xlabel('Quarter')
ax2.set_ylabel('Unemployment Rate (%)', color="r")
ax2.tick_params(axis='y', labelcolor="r")
plt.title("London's Unemployment Rate (1995-2023)")
plt.show()

# Visualization for Housing Prices, Annual Growth Rate in London, and Unemployment Rate

fig, ax1 = plt.subplots(figsize=(14, 8))
ax1.plot(housing_df_quarterly.index, housing_df_quarterly[('London', 'Value')], color="b", label="Housing Prices (London)")
ax1.set_xlabel('Quarter')
ax1.set_ylabel('Housing Prices', color="b")
ax1.tick_params(axis='y', labelcolor="b")
ax1.legend(loc="upper left")

ax2 = ax1.twinx()
ax2.plot(housing_df_quarterly.index, housing_df_quarterly[('London', 'Annual growth')], color="g", linestyle='--', label="Annual Growth Rate (London)")
ax2.set_ylabel('Growth Rate (%)', color="g")
ax2.tick_params(axis='y', labelcolor="g")
ax2.legend(loc="upper right")

ax3 = ax1.twinx()
ax3.plot(unemployment_df_filtered["Quarter"], unemployment_df_filtered["London"], color="r", label="Unemployment Rate (London)")
ax3.set_ylabel('Unemployment Rate (%)', color="r")
ax3.tick_params(axis='y', labelcolor="r")
ax3.spines['right'].set_position(('outward', 60))
ax3.legend(loc="upper center")

plt.title("London's Housing Prices, Annual Growth Rate vs Unemployment Rate (1995-2023)")
plt.tight_layout()
plt.show()


#Saving the prepared datasets
file_path = base_path + "dataset_prepared.xlsx"

# Flatten the MultiIndex columns
housing_df_quarterly.columns = ['_'.join(col).strip() for col in housing_df_quarterly.columns.values]

# Save the data
with pd.ExcelWriter(file_path) as writer:
    housing_df_quarterly.to_excel(writer, sheet_name='Housing', index=False)
    unemployment_df_filtered.to_excel(writer, sheet_name='Unemployment', index=False)

