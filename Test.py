import pandas as pd

# Load the Excel file
df = pd.read_excel("testDataark.xlsx")

# Print the dataframe
#print(df)

# Get the email of a specific customer by their name
def get_email(name):
    email = df.loc[df['Customer name'] == name, 'VAT  nr. NO'].iloc[0]
    return email

# Get the customer number of a specific customer by their name
def get_customer_nr(name):
    customer_nr = df.loc[df['Customer name'] == name, 'Path'].iloc[0]
    return customer_nr

# Example usage
email = get_email("7 Days") 
#customer_nr = get_customer_nr("7 Days")
print("Email:", email)
#print("Customer Number:", customer_nr)



