import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Function Definitions

def scenario_analysis(cost_category, change_percentage, merged_df):
    original_cost = merged_df.loc[merged_df['Cost Category'] == cost_category, 'Amount_Actual'].iloc[0]
    adjusted_cost = original_cost * (1 + change_percentage / 100)
    variance = adjusted_cost - merged_df.loc[merged_df['Cost Category'] == cost_category, 'Amount_Budget'].iloc[0]
    return adjusted_cost, variance

# Load Data with Error Handling

try:
    budget_df = pd.read_excel('budgeted_costs.xlsx', sheet_name='Budget')
    actual_df = pd.read_excel('actual_costs.xlsx', sheet_name='Actual')
    schedule_df = pd.read_excel('schedule_data.xlsx', sheet_name='Schedule')  # Planned vs actual timelines
    risk_df = pd.read_excel('risk_data.xlsx', sheet_name='Risks')
except FileNotFoundError as e:
    print(f"Error loading files: {e}")
    exit()

# Define Weights
weights = {
    'Labor Costs': 0.20, 'Material Costs': 0.35,
    'Equipment & Machinery Costs': 0.15, 'Subcontractor Costs': 0.10,
    'Indirect Costs': 0.10, 'Contingency Allowances': 0.10
}

# Merge Dataframes with Error Handling
try:
    merged_df = budget_df.merge(actual_df, on='Cost Category', suffixes=('_Budget', '_Actual'))
except KeyError as e:
    print(f"Merge error: {e}")
    exit()

# Variance Analysis
merged_df['Variance'] = merged_df['Amount_Actual'] - merged_df['Amount_Budget']
merged_df['Weight'] = merged_df['Cost Category'].map(weights)
merged_df['Weighted Variance'] = merged_df['Variance'] * merged_df['Weight']

# CPI Calculation
merged_df['CPI'] = merged_df['Amount_Budget'] / merged_df['Amount_Actual']

# SPI Calculation
schedule_df['SPI'] = schedule_df['Planned Duration'] / schedule_df['Actual Duration']

# Trend Analysis (quarterly)
trend_df = actual_df.copy()
trend_df['Quarter'] = pd.to_datetime(trend_df['Date']).dt.to_period('Q')
quarterly_trend = trend_df.groupby(['Quarter', 'Cost Category'])['Amount_Actual'].sum().reset_index()

# Risk Impact Analysis
risk_df = pd.read_excel('risk_data.xlsx', sheet_name='Risks')
risk_merged = merged_df.merge(risk_df, on='Cost Category')
risk_impact = risk_merged.groupby('Risk Factor')['Variance'].sum().reset_index()

# Scenario Analysis with User Input
cost_category_input = 'Material Costs'  # Example default value
change_percentage_input = 10  # Example default value
adjusted_cost, variance = scenario_analysis(cost_category_input, change_percentage_input, merged_df)

print(f"Adjusted cost for {cost_category_input}: {adjusted_cost:,.2f}")
print(f"Variance from Budget: {variance:,.2f}")

# Visual Reporting
# CPI Visualization
fig, ax = plt.subplots()
merged_df.plot.bar(x='Cost Category', y='CPI', legend=False, ax=ax)
ax.axhline(1, color='red', linestyle='--')
ax.set_ylabel('CPI')
ax.set_title('Cost Performance Index (CPI)')
plt.tight_layout()
plt.savefig('CPI_analysis.png')

# Export Analysis
merged_df.to_excel('detailed_cost_analysis.xlsx', index=False)
schedule_df.to_excel('schedule_performance.xlsx', index=False)
quarterly_trend.to_excel('quarterly_trend_analysis.xlsx', index=False)
risk_impact.to_excel('risk_impact_analysis.xlsx', index=False)

print("Analysis complete and reports generated successfully.")
