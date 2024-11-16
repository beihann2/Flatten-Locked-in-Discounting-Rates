import pandas as pd
import numpy as np
from scipy.optimize import newton

# Load data from specific sheets by name

file_path = 'D:/LCKIN_goal_seek/CFS-Reins.xlsm'
sheets_to_load = ['CFS_BEL', 'CFS_PAD']  # List the names of the sheets you want to load
dfs = pd.read_excel(file_path, sheet_name = sheets_to_load, engine = 'openpyxl')

# dfs will be a dictionary with sheet names as keys and dataframes as values
df_BEL = dfs['CFS_BEL']
df_PAD = dfs['CFS_PAD']

agg_dict = {'BEL': 'sum'}
agg_dict.update({f'M{i}': 'sum' for i in range(1272)})  # Adjust range according to your columns

# Group and aggregate based on dictionary
grouped_df_BEL = df_BEL.groupby('CONTRACT_GROUP_CODE').agg(agg_dict).reset_index()
grouped_df_PAD = df_PAD.groupby('CONTRACT_GROUP_CODE').agg(agg_dict).reset_index()

# Merge DataFrames on 'CONTRACT_GROUP_CODE'
merged_df = pd.merge(grouped_df_BEL, grouped_df_PAD, on = 'CONTRACT_GROUP_CODE', suffixes = ('_BEL', '_PAD'))


# Parameters
adjustment_factor = 1 - 0.3447

# List of columns to process (assuming 'M0' to 'M1271' and 'BEL')
columns = ['BEL'] + [f'M{i}' for i in range(1272)]

# Applying the formula for each column
for column in columns:
    merged_df[column] = merged_df[f'{column}_BEL'] + (merged_df[f'{column}_PAD'] - merged_df[f'{column}_BEL']) * adjustment_factor

merged_df['BEL'] = -merged_df['BEL_PAD']

# Drop the original '_BEL' and '_PAD' columns as they are no longer needed
merged_df.drop([f'{column}_BEL' for column in columns] + [f'{column}_PAD' for column in columns], axis=1, inplace=True)

# merged_df['BEL'] = -merged_df['BEL']

# Define the NPV function
def npv(rate, cash_flows):
    # The first two elements (for 'BEL' and 'M0') are at time 0.
    # Remaining elements (from 'M1' to 'M1271') start from time 1 to 1271.
    months = np.array([0, 0] + list(range(1, len(cash_flows) - 1)))
    return np.sum(cash_flows / (1 + rate) ** months)

def npv_derivative(rate, cash_flows):
    months = np.array([0, 0] + list(range(1, len(cash_flows) - 1)))
    return np.sum(-months * cash_flows / (1 + rate) ** (months + 1))

def goal_seek_npv(cash_flows, target_npv = 0, initial_guess = 0.000001):
    try:
        # Using the Newton method to find the root of the npv function minus target_npv
        rate = newton(lambda r: npv(r, cash_flows) - target_npv, x0 = initial_guess, fprime = lambda r: npv_derivative(r, cash_flows))
        return rate
    except RuntimeError:
        return None  # Handling cases where the Newton method does not converge

# Initialize the LCKIN column
merged_df['LCKIN'] = 0

for contract_group in merged_df['CONTRACT_GROUP_CODE'].unique():
    # Filter the DataFrame to get the cash flows for this contract group
    selection = merged_df['CONTRACT_GROUP_CODE'] == contract_group
    cash_flows = merged_df.loc[selection, ['BEL'] + [f'M{i}' for i in range(1272)]].values

    # Apply the goal seek function to each row of the filtered DataFrame
    for index, cash_flows in zip(merged_df[selection].index, cash_flows):
        monthly_rate = goal_seek_npv(cash_flows, initial_guess = 0.000001)
        if monthly_rate is not None:
            # Convert the monthly discount rate to an annualized rate
            annualized_rate = (1 + monthly_rate) ** 12 - 1
            merged_df.at[index, 'LCKIN'] = monthly_rate
        else:
            merged_df.at[index, 'LCKIN'] = np.nan

# Select only the required columns
#result_df = merged_df[['CONTRACT_GROUP_CODE', 'LCKIN']]

# Output path for the new Excel file
output_path = "D:/LCKIN_goal_seek/result-reins.xlsx"

# Write the result DataFrame to an Excel file
merged_df.to_excel(output_path, index=False)

print(f"Results have been written to {output_path}")