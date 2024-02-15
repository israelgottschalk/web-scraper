import pandas as pd

def calculate_firm_relevance(row):
    big_tech_firms = ["Apple", "Google", "Amazon", "Facebook", "Meta", "iPhone", "iPay", "IOS", "Alphabet", "Instagram", "Microsoft", "Windows", "tech giants", "digital platforms", " X ", "Twitter", "Elon Musk", "Big Tech"]
    big_payment_firms = ["Visa", "Mastercard", "Master ", "Alibaba", "PayPal", "American Express", "Amex", "JPMorgan", "Square", "Tencent", "Bank of America", "Adyen"]
    
    firm_relevance_score = 0
    
    # Check for Big Tech firms
    for firm in big_tech_firms:
        if firm.lower() in row['Companies'].lower():
            firm_relevance_score += 1
    
    # Check for Big Payment firms
    for firm in big_payment_firms:
        if firm.lower() in row['Companies'].lower():
            firm_relevance_score += 1
    
    return firm_relevance_score

def calculate_subject_relevance(row):
    # Exclude 'Partnerships' from categories
    categories_list = [category.strip() for category in row['Categories'].split(',')]
    subject_relevance_score = len(set(categories_list) - set(["Partnerships"]))
    return subject_relevance_score

def calculate_country_score(row):
    country_scores = {"United Kingdom": 3, "Europe": 2, "United States": 2, "Canada": 2, "Australia": 2}
    
    country_score = 0
    
    # Check for each country and assign scores
    for country, score in country_scores.items():
        if country.lower() in row['Countries'].lower():
            country_score = score
    
    return country_score

def calculate_overall_score(row):
    return row['Firm Relevance Score'] + row['Subject Relevance Score'] + row['Country Score']

def main():
    # Load the combined data
    combined_file = 'News_combined_2024-01-19.xlsx'  # Update with the correct file name
    combined_data = pd.read_excel(combined_file)

    # Apply the relevance score calculation functions
    combined_data['Firm Relevance Score'] = combined_data.apply(calculate_firm_relevance, axis=1)
    combined_data['Subject Relevance Score'] = combined_data.apply(calculate_subject_relevance, axis=1)
    combined_data['Country Score'] = combined_data.apply(calculate_country_score, axis=1)
    
    # Calculate the overall score
    combined_data['Overall Score'] = combined_data.apply(calculate_overall_score, axis=1)

    # Save the updated data with relevance scores
    combined_data.to_excel(combined_file, index=False)
    print(f'Updated news data with relevance scores saved to {combined_file}')

if __name__ == "__main__":
    main()
