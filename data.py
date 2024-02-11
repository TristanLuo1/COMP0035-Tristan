import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Import datasets
base_path = "/Users/luorui/Desktop/UCL/COMP0035 Software Engineering/Coursework/"
file_path = base_path + "dataset.xlsx"

# Read the Excel file, specifying the sheet name and the row to skip

housing_df = pd.read_excel(file_path, sheet_name="housing", skiprows=7)

# Rename columns based on the structure we identified
housing_df.columns = ['Month', 'London Value', 'London Annual growth', 'UK Value', 'UK Annual growth']

# Convert the 'Month' column to datetime format
housing_df['Month'] = pd.to_datetime(housing_df['Month'], errors='coerce')


# Set the 'Month' column as the index for the DataFrame
housing_df.set_index('Month', inplace=True)

# Aggregate data by quarters and compute the mean for each quarter
housing_df_quarterly = housing_df.resample('Q').mean()

# Generate a 'Quarter' column to explicitly indicate the quarter for each row
housing_df_quarterly['Quarter'] = housing_df_quarterly.index.to_period('Q')

# Assuming 'Quarter' is the name of the column you want to move to the front
cols = ['Quarter'] + [col for col in housing_df_quarterly.columns if col != 'Quarter']
housing_df_quarterly = housing_df_quarterly[cols]

# Reading the unemployment data from its sheet, assuming no adjustments needed here
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

# Identify outliers for Housing Prices in London
Q1_housing = housing_df_quarterly['London Value'].quantile(0.25)
Q3_housing = housing_df_quarterly['London Value'].quantile(0.75)
IQR_housing = Q3_housing - Q1_housing

housing_outliers = housing_df_quarterly[
    (housing_df_quarterly['London Value'] < (Q1_housing - 1.5 * IQR_housing)) |
    (housing_df_quarterly['London Value'] > (Q3_housing + 1.5 * IQR_housing))
]

# Identify outliers for Unemployment Rate in London
Q1_unemployment = unemployment_df_filtered['London'].quantile(0.25)
Q3_unemployment = unemployment_df_filtered['London'].quantile(0.75)
IQR_unemployment = Q3_unemployment - Q1_unemployment

unemployment_outliers = unemployment_df_filtered[
    (unemployment_df_filtered['London'] < (Q1_unemployment - 1.5 * IQR_unemployment)) |
    (unemployment_df_filtered['London'] > (Q3_unemployment + 1.5 * IQR_unemployment))
]

print("Outliers for Housing Prices in London:")
print(housing_outliers)

print("\nOutliers for Unemployment Rate in London:")
print(unemployment_outliers)


# Box-plot for Housing Prices in London
plt.figure(figsize=(14, 6))
plt.subplot(1, 2, 1)
sns.boxplot(y=housing_df_quarterly['London Value'])
plt.title('Box plot of Housing Prices in London')
plt.ylabel('Housing Prices')

# Box-plot for Unemployment Rate in London
plt.subplot(1, 2, 2)
sns.boxplot(y=unemployment_df_filtered['London'])
plt.title('Box plot of Unemployment Rate in London')
plt.ylabel('Unemployment Rate')

plt.tight_layout()
plt.show()


# Visualization for Housing Prices and Annual Growth Rate in London
fig, ax1 = plt.subplots(figsize=(14, 7))
ax1.plot(housing_df_quarterly.index, housing_df_quarterly['London Value'], color="b", label="Housing Prices (London)")
ax1.set_xlabel('Quarter')
ax1.set_ylabel('Housing Prices', color="b")
ax1.tick_params(axis='y', labelcolor="b")
ax2 = ax1.twinx()
ax2.plot(housing_df_quarterly.index, housing_df_quarterly['London Annual growth'], color="g", linestyle='--', label="Annual Growth Rate (London)")
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
ax1.plot(housing_df_quarterly.index, housing_df_quarterly['London Value'], color="b", label="Housing Prices (London)")
ax1.set_xlabel('Quarter')
ax1.set_ylabel('Housing Prices', color="b")
ax1.tick_params(axis='y', labelcolor="b")
ax1.legend(loc="upper left")

ax2 = ax1.twinx()
ax2.plot(housing_df_quarterly.index, housing_df_quarterly['London Annual growth'], color="g", linestyle='--', label="Annual Growth Rate (London)")
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

# Save the data
with pd.ExcelWriter(file_path) as writer:
    housing_df_quarterly.to_excel(writer, sheet_name='Housing', index=False)
    unemployment_df_filtered.to_excel(writer, sheet_name='Unemployment', index=False)

