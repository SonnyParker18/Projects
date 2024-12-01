import pandas as pd
import matplotlib.pyplot as plt

# Load the data into a DataFrame
file_path = r'C:\Users\London\Downloads\ASSESSMENT DATA Portfolio Analyst.xlsx'
df = pd.read_excel(file_path)

# Ensure refdate is in datetime format
df['refdate'] = pd.to_datetime(df['refdate'], format='%d/%m/%Y')

# Remove weekends from the dataset
df = df[~df['refdate'].dt.weekday.isin([5, 6])]

# Normalize the weights so that they sum to 100% for each day
df['Normalized Weight'] = df.groupby('refdate')['Weight (%)'].transform(lambda x: x / x.sum())

# Calculate the Total Overall ESG Score for each day using the normalized weighted scores
df['Normalized Weighted Overall ESG Score'] = df['Normalized Weight'] * df['Overall ESG Score']
df['Normalized Weighted ESG Environmental Score'] = df['Normalized Weight'] * df['Overall ESG Environmental Score']
df['Normalized Weighted ESG Social Score'] = df['Normalized Weight'] * df['Overall ESG Social Score']
df['Normalized Weighted ESG Governance Score'] = df['Normalized Weight'] * df['Overall ESG Governance Score']

daily_esg_scores = df.groupby('refdate')[['Normalized Weighted Overall ESG Score', 'Normalized Weighted ESG Environmental Score', 'Normalized Weighted ESG Social Score', 'Normalized Weighted ESG Governance Score']].sum().reset_index()

# Plotting the trends of ESG Scores through time
plt.figure(figsize=(12, 8))
plt.plot(daily_esg_scores['refdate'], daily_esg_scores['Normalized Weighted Overall ESG Score'], label='Overall ESG Score')
plt.plot(daily_esg_scores['refdate'], daily_esg_scores['Normalized Weighted ESG Environmental Score'], label='Environmental ESG Score')
plt.plot(daily_esg_scores['refdate'], daily_esg_scores['Normalized Weighted ESG Social Score'], label='Social ESG Score')
plt.plot(daily_esg_scores['refdate'], daily_esg_scores['Normalized Weighted ESG Governance Score'], label='Governance ESG Score')
plt.title('Normalized Weighted ESG Scores Through Time')
plt.xlabel('Date')
plt.ylabel('Normalized Weighted ESG Score')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\London\Documents\Normalized_ESG_Scores_Trend.png')
plt.show()

# Group all of the scores based on the given categories
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
df['Environmental ESG Category'] = df['Overall ESG Environmental Score'].apply(categorize_esg_score)
df['Social ESG Category'] = df['Overall ESG Social Score'].apply(categorize_esg_score)
df['Governance ESG Category'] = df['Overall ESG Governance Score'].apply(categorize_esg_score)

# Calculate the weighted scores for each category over time
category_scores = df.groupby(['refdate', 'Overall ESG Category'])['Normalized Weight'].sum().unstack().fillna(0)
category_scores_env = df.groupby(['refdate', 'Environmental ESG Category'])['Normalized Weight'].sum().unstack().fillna(0)
category_scores_social = df.groupby(['refdate', 'Social ESG Category'])['Normalized Weight'].sum().unstack().fillna(0)
category_scores_gov = df.groupby(['refdate', 'Governance ESG Category'])['Normalized Weight'].sum().unstack().fillna(0)

# Plotting the ESG scores based on categories over time
plt.figure(figsize=(12, 8))
for category in category_scores.columns:
    plt.plot(category_scores.index, category_scores[category], label=category)
plt.title('Overall ESG Scores by MSCI Rating')
plt.xlabel('Date')
plt.ylabel('Normalized Weight')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\London\Documents\Normalized_Overall_ESG_Categories_Trend.png')
plt.show()

plt.figure(figsize=(12, 8))
for category in category_scores_env.columns:
    plt.plot(category_scores_env.index, category_scores_env[category], label=category)
plt.title('Environmental ESG Scores by MSCI Rating')
plt.xlabel('Date')
plt.ylabel('Normalized Weight')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\London\Documents\Normalized_Environmental_ESG_Categories_Trend.png')
plt.show()

plt.figure(figsize=(12, 8))
for category in category_scores_social.columns:
    plt.plot(category_scores_social.index, category_scores_social[category], label=category)
plt.title('Social ESG Scores by MSCI Rating')
plt.xlabel('Date')
plt.ylabel('Normalized Weight')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\London\Documents\Normalized_Social_ESG_Categories_Trend.png')
plt.show()

plt.figure(figsize=(12, 8))
for category in category_scores_gov.columns:
    plt.plot(category_scores_gov.index, category_scores_gov[category], label=category)
plt.title('Governance ESG Scores by MSCI Rating')
plt.xlabel('Date')
plt.ylabel('Normalized Weight')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\London\Documents\Normalized_Governance_ESG_Categories_Trend.png')
plt.show()

# Calculate the percentage weight in each category as of the last day of the period
last_day = df['refdate'].max()
last_day_data = df[df['refdate'] == last_day]

category_weights = last_day_data.groupby('Overall ESG Category')['Normalized Weight'].sum().reset_index()
category_weights['Percentage Weight'] = category_weights['Normalized Weight'] * 100

# Plotting the pie chart for Overall ESG Category weights as of the last day of the period
plt.figure(figsize=(10, 7))
plt.pie(category_weights['Percentage Weight'], labels=category_weights['Overall ESG Category'], autopct='%1.1f%%', startangle=140)
plt.title('Overall ESG Category Weights as of Last Day')
plt.savefig(r'C:\Users\London\Documents\Normalized_Overall_ESG_Category_Weights.png')
plt.show()

# Sum up the normalized weighted overall ESG score for the last day of the period
total_weighted_esg_last_day = last_day_data['Normalized Weighted Overall ESG Score'].sum()

# Assign the summed-up score into a category
total_weighted_esg_category = categorize_esg_score(total_weighted_esg_last_day)
print(f'Total Normalized Weighted Overall ESG Score on Last Day: {total_weighted_esg_last_day:.2f}, Category: {total_weighted_esg_category}')


