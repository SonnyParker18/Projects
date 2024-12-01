import pandas as pd
import matplotlib.pyplot as plt

# Path to your Excel file
file_path = r'C:\Users\London\Downloads\ASSESSMENT DATA Portfolio Analyst.xlsx'

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Ensure refdate is in datetime format
df['refdate'] = pd.to_datetime(df['refdate'], format='%d/%m/%Y')

# Filter out rows where 'GICS_sector' is empty to remove currencies and other non-equities
df = df[df['GICS_sector'].notna()]

# Remove weekends from the dataset
df = df[~df['refdate'].dt.weekday.isin([5, 6])]

# Create market cap buckets
bins = [-float('inf'), 250e6, 2e9, 10e9, 200e9, float('inf')]
labels = ['Micro-cap', 'Small-cap', 'Mid-cap', 'Large-cap', 'Mega-cap']
df['Market Cap Bucket'] = pd.cut(df['Market Capitalization (USD)'], bins=bins, labels=labels)

# Sort the DataFrame in order to split out individual securities (including country to split different listings of the same security)
df.sort_values(by=['Asset Name', 'Country', 'refdate'], inplace=True)

# Calculate the daily price change for each security
df['daily_perf'] = df.groupby(['Asset Name', 'Country'])['Price(USD)'].pct_change()

# Shift the Weight (%) by one date in order to get the Open Exposure for contribution calculation
df['Open Exposure'] = df.groupby(['Asset Name', 'Country'])['Weight (%)'].shift(1)

# Multiply the daily performance by the Open Exposure
df['Contribution'] = df['daily_perf'] * df['Open Exposure']

# Sum the weighted daily performance by date
weighted_performance_sum = df.groupby('refdate')['Contribution'].sum().reset_index()

# Calculate the compounded return
weighted_performance_sum['compounded_return'] = (weighted_performance_sum['Contribution'] + 1).cumprod() - 1

# Plotting the compounded return
plt.figure(figsize=(12, 8))
plt.plot(weighted_performance_sum['refdate'], weighted_performance_sum['compounded_return'], marker='o')
plt.title('Compounded Gross Performance')
plt.xlabel('Ref Date')
plt.ylabel('Compounded Return')
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# ITD Tables
final_day = df['refdate'].max()
final_day_data = df[df['refdate'] == final_day]

# Sum the 'Weight (%)' and 'Active Weight (%)' by GICS_sector on the final day
final_day_summary = final_day_data.groupby('GICS_sector')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()

# Rename columns in the final day summary
final_day_summary.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

# Merge the summary back into the contributions table
contribution_by_sector = df.groupby('GICS_sector')['Contribution'].sum().reset_index()
contribution_by_sector = contribution_by_sector.merge(final_day_summary, on='GICS_sector', how='left')

# Calculate the total contributions, exposures, and active exposures for GICS_sector before formatting
total_contribution_sector_value = contribution_by_sector['Contribution'].sum()
total_exposure_sector_value = contribution_by_sector['Exposure'].sum()
total_active_exposure_sector_value = contribution_by_sector['Active Exposure'].sum()

# Format the Contribution, Exposure, and Active Exposure as percentages
contribution_by_sector['Contribution'] = contribution_by_sector['Contribution'].apply(lambda x: f'{x:.2%}')
contribution_by_sector['Exposure'] = contribution_by_sector['Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
contribution_by_sector['Active Exposure'] = contribution_by_sector['Active Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')

# Add total row with formatted total values
total_contribution_sector = pd.DataFrame([['Total', f'{total_exposure_sector_value:.2%}', f'{total_active_exposure_sector_value:.2%}', f'{total_contribution_sector_value:.2%}']], columns=['GICS_sector', 'Exposure', 'Active Exposure', 'Contribution'])
contribution_by_sector = pd.concat([contribution_by_sector, total_contribution_sector], ignore_index=True)

# Reorder columns to move 'Contribution' to the end
contribution_by_sector = contribution_by_sector[['GICS_sector', 'Exposure', 'Active Exposure', 'Contribution']]

# Displaying the table of contributions per sector
print("Contributions by GICS Sector ITD:")
print(contribution_by_sector.to_string(index=False))

# Sum the contribution by Country
contribution_by_country = df.groupby('Country')['Contribution'].sum().reset_index()
final_day_summary_country = final_day_data.groupby('Country')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()
final_day_summary_country.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)
contribution_by_country = contribution_by_country.merge(final_day_summary_country, on='Country', how='left')
total_contribution_country_value = contribution_by_country['Contribution'].sum()
total_exposure_country_value = contribution_by_country['Exposure'].sum()
total_active_exposure_country_value = contribution_by_country['Active Exposure'].sum()
contribution_by_country['Contribution'] = contribution_by_country['Contribution'].apply(lambda x: f'{x:.2%}')
contribution_by_country['Exposure'] = contribution_by_country['Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
contribution_by_country['Active Exposure'] = contribution_by_country['Active Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
total_contribution_country = pd.DataFrame([['Total', f'{total_exposure_country_value:.2%}', f'{total_active_exposure_country_value:.2%}', f'{total_contribution_country_value:.2%}']], columns=['Country', 'Exposure', 'Active Exposure', 'Contribution'])
contribution_by_country = pd.concat([contribution_by_country, total_contribution_country], ignore_index=True)

# Reorder columns to move 'Contribution' to the end
contribution_by_country = contribution_by_country[['Country', 'Exposure', 'Active Exposure', 'Contribution']]

print("\nContributions by Country ITD:")
print(contribution_by_country.to_string(index=False))

# Sum the contribution by Market Cap Bucket
contribution_by_market_cap = df.groupby('Market Cap Bucket')['Contribution'].sum().reset_index()
final_day_summary_market_cap = final_day_data.groupby('Market Cap Bucket')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()

# Rename columns in the final day summary
final_day_summary_market_cap.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

contribution_by_market_cap = contribution_by_market_cap.merge(final_day_summary_market_cap, on='Market Cap Bucket', how='left')
total_contribution_market_cap_value = contribution_by_market_cap['Contribution'].sum()
total_exposure_market_cap_value = contribution_by_market_cap['Exposure'].sum()
total_active_exposure_market_cap_value = contribution_by_market_cap['Active Exposure'].sum()
contribution_by_market_cap['Contribution'] = contribution_by_market_cap['Contribution'].apply(lambda x: f'{x:.2%}')
contribution_by_market_cap['Exposure'] = contribution_by_market_cap['Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
contribution_by_market_cap['Active Exposure'] = contribution_by_market_cap['Active Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
total_contribution_market_cap = pd.DataFrame([['Total', f'{total_exposure_market_cap_value:.2%}', f'{total_active_exposure_market_cap_value:.2%}', f'{total_contribution_market_cap_value:.2%}']], columns=['Market Cap Bucket', 'Exposure', 'Active Exposure', 'Contribution'])
contribution_by_market_cap = pd.concat([contribution_by_market_cap, total_contribution_market_cap], ignore_index=True)

# Reorder columns to move 'Contribution' to the end
contribution_by_market_cap = contribution_by_market_cap[['Market Cap Bucket', 'Exposure', 'Active Exposure', 'Contribution']]

print("\nContributions by Market Capitalization Bucket ITD:")
print(contribution_by_market_cap.to_string(index=False))

# YTD Tables
start_date = '2023-01-01'
df_ytd = df[df['refdate'] >= start_date]

# Get the final day
final_day_ytd = df_ytd['refdate'].max()
final_day_data_ytd = df_ytd[df_ytd['refdate'] == final_day_ytd]

# Sum the 'Weight (%)' and 'Active Weight (%)' by GICS_sector on the final day
final_day_summary_ytd = final_day_data_ytd.groupby('GICS_sector')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()

# Rename columns in the final day summary
final_day_summary_ytd.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

# Merge the summary back into the contributions table
contribution_by_sector_ytd = df_ytd.groupby('GICS_sector')['Contribution'].sum().reset_index

# YTD Tables

# Sum the 'Weight (%)' and 'Active Weight (%)' by GICS_sector on the final day
final_day_summary_ytd = final_day_data_ytd.groupby('GICS_sector')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()

# Rename columns in the final day summary
final_day_summary_ytd.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

# Merge the summary back into the contributions table
contribution_by_sector_ytd = df_ytd.groupby('GICS_sector')['Contribution'].sum().reset_index()
contribution_by_sector_ytd = contribution_by_sector_ytd.merge(final_day_summary_ytd, on='GICS_sector', how='left')

# Calculate the total contributions, exposures, and active exposures for GICS_sector before formatting
total_contribution_sector_value_ytd = contribution_by_sector_ytd['Contribution'].sum()
total_exposure_sector_value_ytd = contribution_by_sector_ytd['Exposure'].sum()
total_active_exposure_sector_value_ytd = contribution_by_sector_ytd['Active Exposure'].sum()

# Format the Contribution, Exposure, and Active Exposure as percentages
contribution_by_sector_ytd['Contribution'] = contribution_by_sector_ytd['Contribution'].apply(lambda x: f'{x:.2%}')
contribution_by_sector_ytd['Exposure'] = contribution_by_sector_ytd['Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
contribution_by_sector_ytd['Active Exposure'] = contribution_by_sector_ytd['Active Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')

# Add total row with formatted total values
total_contribution_sector_ytd = pd.DataFrame([['Total', f'{total_exposure_sector_value_ytd:.2%}', f'{total_active_exposure_sector_value_ytd:.2%}', f'{total_contribution_sector_value_ytd:.2%}']], columns=['GICS_sector', 'Exposure', 'Active Exposure', 'Contribution'])
contribution_by_sector_ytd = pd.concat([contribution_by_sector_ytd, total_contribution_sector_ytd], ignore_index=True)

# Reorder columns to move 'Contribution' to the end
contribution_by_sector_ytd = contribution_by_sector_ytd[['GICS_sector', 'Exposure', 'Active Exposure', 'Contribution']]

# Displaying the table of contributions per sector
print("\nContributions by GICS Sector YTD:")
print(contribution_by_sector_ytd.to_string(index=False))

# Sum the contribution by Country for YTD
contribution_by_country_ytd = df_ytd.groupby('Country')['Contribution'].sum().reset_index()
final_day_summary_country_ytd = final_day_data_ytd.groupby('Country')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()
final_day_summary_country_ytd.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)
contribution_by_country_ytd = contribution_by_country_ytd.merge(final_day_summary_country_ytd, on='Country', how='left')
total_contribution_country_value_ytd = contribution_by_country_ytd['Contribution'].sum()
total_exposure_country_value_ytd = contribution_by_country_ytd['Exposure'].sum()
total_active_exposure_country_value_ytd = contribution_by_country_ytd['Active Exposure'].sum()
contribution_by_country_ytd['Contribution'] = contribution_by_country_ytd['Contribution'].apply(lambda x: f'{x:.2%}')
contribution_by_country_ytd['Exposure'] = contribution_by_country_ytd['Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
contribution_by_country_ytd['Active Exposure'] = contribution_by_country_ytd['Active Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
total_contribution_country_ytd = pd.DataFrame([['Total', f'{total_exposure_country_value_ytd:.2%}', f'{total_active_exposure_country_value_ytd:.2%}', f'{total_contribution_country_value_ytd:.2%}']], columns=['Country', 'Exposure', 'Active Exposure', 'Contribution'])
contribution_by_country_ytd = pd.concat([contribution_by_country_ytd, total_contribution_country_ytd], ignore_index=True)

# Reorder columns to move 'Contribution' to the end
contribution_by_country_ytd = contribution_by_country_ytd[['Country', 'Exposure', 'Active Exposure', 'Contribution']]

print("\nContributions by Country YTD:")
print(contribution_by_country_ytd.to_string(index=False))

# Sum the contribution by Market Cap Bucket for YTD
contribution_by_market_cap_ytd = df_ytd.groupby('Market Cap Bucket')['Contribution'].sum().reset_index()
final_day_summary_market_cap_ytd = final_day_data_ytd.groupby('Market Cap Bucket')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()
final_day_summary_market_cap_ytd.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)
contribution_by_market_cap_ytd = contribution_by_market_cap_ytd.merge(final_day_summary_market_cap_ytd, on='Market Cap Bucket', how='left')
total_contribution_market_cap_value_ytd = contribution_by_market_cap_ytd['Contribution'].sum()
total_exposure_market_cap_value_ytd = contribution_by_market_cap_ytd['Exposure'].sum()
total_active_exposure_market_cap_value_ytd = contribution_by_market_cap_ytd['Active Exposure'].sum()
contribution_by_market_cap_ytd['Contribution'] = contribution_by_market_cap_ytd['Contribution'].apply(lambda x: f'{x:.2%}')
contribution_by_market_cap_ytd['Exposure'] = contribution_by_market_cap_ytd['Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
contribution_by_market_cap_ytd['Active Exposure'] = contribution_by_market_cap_ytd['Active Exposure'].apply(lambda x: f'{x:.2%}' if pd.notnull(x) else '')
total_contribution_market_cap_ytd = pd.DataFrame([['Total', f'{total_exposure_market_cap_value_ytd:.2%}', f'{total_active_exposure_market_cap_value_ytd:.2%}', f'{total_contribution_market_cap_value_ytd:.2%}']], columns=['Market Cap Bucket', 'Exposure', 'Active Exposure', 'Contribution'])
contribution_by_market_cap_ytd = pd.concat([contribution_by_market_cap_ytd, total_contribution_market_cap_ytd], ignore_index=True)

# Reorder columns to move 'Contribution' to the end
contribution_by_market_cap_ytd = contribution_by_market_cap_ytd[['Market Cap Bucket', 'Exposure', 'Active Exposure', 'Contribution']]

print("\nContributions by Market Capitalization Bucket YTD:")
print(contribution_by_market_cap_ytd.to_string(index=False))

# Exporting tables to an XLSX file
with pd.ExcelWriter(r'C:\Users\London\Documents\Portfolio_Performance_Analysis.xlsx') as writer:
    contribution_by_sector.to_excel(writer, sheet_name='ITD_GICS_Sector', index=False)
    contribution_by_country.to_excel(writer, sheet_name='ITD_Country', index=False)
    contribution_by_market_cap.to_excel(writer, sheet_name='ITD_Market_Cap', index=False)
    contribution_by_sector_ytd.to_excel(writer, sheet_name='YTD_GICS_Sector', index=False)
    contribution_by_country_ytd.to_excel(writer, sheet_name='YTD_Country', index=False)
    contribution_by_market_cap_ytd.to_excel(writer, sheet_name='YTD_Market_Cap', index=False)
