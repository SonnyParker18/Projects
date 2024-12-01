import pandas as pd
import matplotlib.pyplot as plt

# Configure Pandas to display the full table in the console
pd.options.display.max_rows = None
pd.options.display.max_columns = None
pd.options.display.width = 1000

# Load the data into a DataFrame
file_path = r'C:\Users\London\Downloads\ASSESSMENT DATA Portfolio Analyst.xlsx'
df = pd.read_excel(file_path)

# Ensure refdate is in datetime format
df['refdate'] = pd.to_datetime(df['refdate'], format='%d/%m/%Y')

# Remove weekends from the dataset
df = df[~df['refdate'].dt.weekday.isin([5, 6])]

# Normalize the weights so that they sum to 100% for each day
df['Normalized Weight'] = df.groupby('refdate')['Weight (%)'].transform(lambda x: x / x.sum())

# Calculate the normalized weighted ESG scores
df['Normalized Weighted Overall ESG Score'] = df['Normalized Weight'] * df['Overall ESG Score']

# Define function for ESG category assignment
def categorize_esg_score(score):
    if 8.571 <= score <= 10.000:
        return 'AAA'
    elif 7.143 <= score < 8.571:
        return 'AA'
    elif 5.714 <= score < 7.143:
        return 'A'
    elif 4.286 <= score < 5.714:
        return 'BBB'
    elif 2.857 <= score < 4.286:
        return 'BB'
    elif 1.429 <= score < 2.857:
        return 'B'
    else:
        return 'CCC'

df['Overall ESG Category'] = df['Overall ESG Score'].apply(categorize_esg_score)

# Create a column to track the previous ESG category
df['Previous ESG Category'] = df.groupby('Asset Name')['Overall ESG Category'].shift(1)

# Define category ranking to compare movements
category_order = {'AAA': 6, 'AA': 5, 'A': 4, 'BBB': 3, 'BB': 2, 'B': 1, 'CCC': 0}

# Map categories to numerical ranks for comparison
df['Category Rank'] = df['Overall ESG Category'].map(category_order)
df['Previous Category Rank'] = df.groupby('Asset Name')['Category Rank'].shift(1)

# Identify category changes
def categorize_change(row):
    if pd.isna(row['Previous Category Rank']):
        return 'No Previous Data'
    elif row['Category Rank'] > row['Previous Category Rank']:
        return 'Up'
    elif row['Category Rank'] < row['Previous Category Rank']:
        return 'Down'
    else:
        return 'No Change'

df['Category Change'] = df.apply(categorize_change, axis=1)

# Filter only rows where there is a category change
category_changes = df[df['Category Change'].isin(['Up', 'Down'])]

# Summarize the changes for each security
category_changes_summary = category_changes[['Asset Name', 'refdate', 'Overall ESG Category', 'Previous ESG Category', 'Category Change', 'Normalized Weight']]
category_changes_summary = category_changes_summary.sort_values(['refdate', 'Asset Name'])

# Save the category changes summary to an Excel file
category_changes_summary.to_excel(r'C:\Users\London\Documents\Category_Changes_With_Dates.xlsx', index=False)

# Display the full table of category changes in the console
print("\nFull Table of Category Changes:")
print(category_changes_summary)

# Save last day's changes for reference
last_day_changes = category_changes[category_changes['refdate'] == df['refdate'].max()]
print("\nSecurities with Category Changes on the Last Observed Date:")
print(last_day_changes[['Asset Name', 'refdate', 'Overall ESG Category', 'Previous ESG Category', 'Category Change', 'Normalized Weight']])

# Save the last day's changes to an Excel file
last_day_changes.to_excel(r'C:\Users\London\Documents\Last_Day_Changes_With_Dates.xlsx', index=False)
