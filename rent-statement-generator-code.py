import pandas as pd

def generate_rent_statement(mpesa_statement_file):
    # Rent amount due and due date
    rent_due = 115000
    due_date = pd.Timestamp.now().replace(day=5)

    # Read the MPESA statement Excel file
    mpesa_df = pd.read_excel(mpesa_statement_file)

    # Sort the transactions by date
    mpesa_df.sort_values(by='Date', inplace=True)

    # Initialize variables for balances
    balance_forward = 0
    balances_due = []

    # Iterate through each transaction in the MPESA statement
    for index, row in mpesa_df.iterrows():
        payment_date = row['Date']
        amount_paid = row['Amount']
        transaction_code = row['Transaction Code']

        # Handle missing amount values
        if pd.isna(amount_paid):
            amount_paid = 0  # Default missing values to 0

        # Calculate balance brought forward for the month
        balance_forward += amount_paid

        # Check if the payment date is within the current month
        if payment_date.month == due_date.month and payment_date.year == due_date.year:
            # Subtract the rent due from the balance forward
            balance_forward -= rent_due

        # Add the balance due at the end of the month to the list
        if payment_date.month != due_date.month or payment_date.year != due_date.year:
            balances_due.append({
                'Month': due_date.strftime('%B'),
                'Year': due_date.year,
                'Balance Due': balance_forward - rent_due
            })
    
    # Create DataFrame for balances due
    balances_due_df = pd.DataFrame(balances_due)

    # Write the rent statement to an Excel file
    rent_statement_file = 'rent_statement.xlsx'
    with pd.ExcelWriter(rent_statement_file) as writer:
        balances_due_df.to_excel(writer, sheet_name='Rent Statement', index=False)

    print(f'Rent statement generated successfully. Saved as {rent_statement_file}')

# Example usage
#mpesa_statement_file = input("mpesa_statement_file.xlsx")

generate_rent_statement('/Users/user/Desktop/Rent-Statement-Generator/mpesa_statement_file.xlsx')