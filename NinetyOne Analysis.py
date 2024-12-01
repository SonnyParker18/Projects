import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


#Importing the data we need
# excel file path
path = 'C:\\Users\\London\\Downloads\\ASSESSMENT DATA Portfolio Analyst.xlsx'

# Load the Excel file into a DataFrame df =
df = pd.read_excel(path)

# Filter out rows where 'GICS_sector' is empty to remove currencies and other non-equities
df = df[df['GICS_sector'].notna()]

# Remove weekends from the dataset
df = df[~df['refdate'].dt.weekday.isin([5, 6])]

# Create market cap buckets
bins = [-float('inf'), 250e6, 2e9, 10e9, 200e9, float('inf')]
labels = ['Micro-cap', 'Small-cap', 'Mid-cap', 'Large-cap', 'Mega-cap']
df['Market Cap Bucket'] = pd.cut(df['Market Capitalization (USD)'], bins=bins, labels=labels)

def plot_exposures(df, group_by_column, value_column, plot_title, legend_title):
    """
    Plots the sum of value_column by group_by_column over time.

    :param df: pandas DataFrame, input data
    :param group_by_column: str, column name to group by (e.g., 'GICS_sector', 'country')
    :param value_column: str, column name of the values to sum (e.g., 'weight (%)')
    :param plot_title: str, title of the plot
    :param legend_title: str, title of the legend
    """
    # Ensure refdate is treated as a datetime object
    df['refdate'] = pd.to_datetime(df['refdate'])

    # Create a pivot table to group by refdate and the specified group_by_column
    pivot_table = pd.pivot_table(df, index=['refdate', group_by_column], values=value_column, aggfunc='sum')

    # Reset the index to make 'refdate' a column again
    pivot_table = pivot_table.reset_index()

    # Plotting
    fig, ax = plt.subplots(figsize=(12, 8))

    # Loop through each group and plot the data
    for group in pivot_table[group_by_column].unique():
        group_data = pivot_table[pivot_table[group_by_column] == group]
        ax.plot(group_data['refdate'], group_data[value_column], label=group)

    # Adding title and labels
    ax.set_title(plot_title)
    ax.set_xlabel('Ref Date')
    ax.set_ylabel(f'Sum of {value_column}')

    # Customize legend
    ax.legend(title=legend_title, bbox_to_anchor=(1.05, 1), loc='upper left')

    # Format the x-axis for better readability
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()

    # Display the plot
    plt.show()

plot_exposures(
    df=df,
    group_by_column="GICS_sector",
    value_column="Weight (%)",
    plot_title="Exposure by Sector Over Time",
    legend_title="GICS Sector"
)
plot_exposures(
    df=df,
    group_by_column="GICS_sector",
    value_column="Active Weight (%)",
    plot_title="Active Exposure by Sector Over Time",
    legend_title="GICS Sector"
)
plot_exposures(
    df=df,
    group_by_column="Country",
    value_column="Weight (%)",
    plot_title="Exposure by Country Over Time",
    legend_title="Country"
)
plot_exposures(
    df=df,
    group_by_column="Country",
    value_column="Active Weight (%)",
    plot_title="Active Exposure by Country Over Time",
    legend_title="Country"
)
plot_exposures(
    df=df,
    group_by_column="GICS_sector",
    value_column="%Contribution to Total Risk",
    plot_title="Sector Risk Contribution",
    legend_title="Sector"
)
plot_exposures(
    df=df,
    group_by_column="GICS_sector",
    value_column="%Contribution to Tracking Error",
    plot_title="Sector TE Contribution",
    legend_title="Sector"
)
plot_exposures(
    df=df,
    group_by_column="Country",
    value_column="%Contribution to Total Risk",
    plot_title="Country Risk Contribution",
    legend_title="Country"
)
plot_exposures(
    df=df,
    group_by_column="Market Cap Bucket",
    value_column="Weight (%)",
    plot_title="Market Cap Exposures",
    legend_title="Market Cap"
)
plot_exposures(
    df=df,
    group_by_column="Market Cap Bucket",
    value_column="%Contribution to Total Risk",
    plot_title="Market Cap Risk Contribution",
    legend_title="Market Cap"
)
