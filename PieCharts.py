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

# ITD Tables
final_day = df['refdate'].max()
final_day_data = df[df['refdate'] == final_day]

# Sum the 'Weight (%)' and 'Active Weight (%)' by GICS_sector on the final day
final_day_summary = final_day_data.groupby('GICS_sector')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()

# Rename columns in the final day summary
final_day_summary.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

# Sum the 'Weight (%)' and 'Active Weight (%)' by Market Cap Bucket on the final day
final_day_summary_market_cap = final_day_data.groupby('Market Cap Bucket')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()
final_day_summary_market_cap.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

# Sum the 'Weight (%)' and 'Active Weight (%)' by Country on the final day
final_day_summary_country = final_day_data.groupby('Country')[['Weight (%)', 'Active Weight (%)']].sum().reset_index()
final_day_summary_country.rename(columns={'Weight (%)': 'Exposure', 'Active Weight (%)': 'Active Exposure'}, inplace=True)

# Create Pie Charts for Exposure and Active Exposure by Market Cap Bucket
fig, axes = plt.subplots(2, 1, figsize=(10, 12))

# Exposure by Market Cap Bucket
axes[0].pie(final_day_summary_market_cap['Exposure'], labels=final_day_summary_market_cap['Market Cap Bucket'], autopct='%1.1f%%', startangle=140)
axes[0].set_title(f'Exposure by Market Cap Bucket as of {final_day.strftime("%Y-%m-%d")}')

# Active Exposure by Market Cap Bucket
axes[1].pie(final_day_summary_market_cap['Active Exposure'], labels=final_day_summary_market_cap['Market Cap Bucket'], autopct='%1.1f%%', startangle=140)
axes[1].set_title(f'Active Exposure by Market Cap Bucket as of {final_day.strftime("%Y-%m-%d")}')

plt.tight_layout()
plt.show()

# Create Pie Charts for Exposure and Active Exposure by GICS Sector
fig, axes = plt.subplots(2, 1, figsize=(10, 12))

# Exposure by GICS Sector
axes[0].pie(final_day_summary['Exposure'], labels=final_day_summary['GICS_sector'], autopct='%1.1f%%', startangle=140)
axes[0].set_title(f'Exposure by Sector as of {final_day.strftime("%Y-%m-%d")}')

# Active Exposure by GICS Sector
axes[1].pie(final_day_summary['Active Exposure'], labels=final_day_summary['GICS_sector'], autopct='%1.1f%%', startangle=140)
axes[1].set_title(f'Active Exposure by Sector as of {final_day.strftime("%Y-%m-%d")}')

plt.tight_layout()
plt.show()

# Create Pie Charts for Exposure and Active Exposure by Country
fig, axes = plt.subplots(2, 1, figsize=(10, 12))

# Exposure by Country
axes[0].pie(final_day_summary_country['Exposure'], labels=final_day_summary_country['Country'], autopct='%1.1f%%', startangle=140)
axes[0].set_title(f'Exposure by Country as of {final_day.strftime("%Y-%m-%d")}')

# Active Exposure by Country
axes[1].pie(final_day_summary_country['Active Exposure'], labels=final_day_summary_country['Country'], autopct='%1.1f%%', startangle=140)
axes[1].set_title(f'Active Exposure by Country as of {final_day.strftime("%Y-%m-%d")}')

plt.tight_layout()
plt.show()
