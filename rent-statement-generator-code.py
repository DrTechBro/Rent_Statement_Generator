import pandas as pd

def generate_rent_statement(mpesa_statement_file, rent_due=115000, due_date=None):
    if due_date is None:
        due_date = pd.Timestamp.now().replace(day=5)

    try:
        # Read the MPESA statement Excel file
        mpesa_df = pd.read_excel(mpesa_statement_file)

        # Sort the transactions by date
        mpesa_df.sort_values(by='Date', inplace=True)

        # Determine the range of months to include
        start_date = mpesa_df['Date'].min().replace(day=1)
        end_date = mpesa_df['Date'].max().replace(day=1)
        current_date = start_date

        monthly_balances = {}

        # Pre-fill the monthly balances dictionary for each month in the range
        while current_date <= end_date:
            monthly_balances[(current_date.year, current_date.month)] = {
                'Amount Paid': 0,
                'Rent Due': rent_due,
                'Balance Due': 0,
                'Transaction Codes': []  # List to store transaction codes
            }
            current_date += pd.DateOffset(months=1)

        # Iterate through each transaction in the MPESA statement
        for index, row in mpesa_df.iterrows():
            payment_date = row['Date'].replace(day=1)
            amount_paid = row['Amount']
            transaction_code = row['Transaction Code']  # Capture the transaction code

            # Handle missing amount values
            if pd.isna(amount_paid):
                continue  # Skip transactions with undefined amounts

            # Update the total paid for the month and add the transaction code
            if (payment_date.year, payment_date.month) in monthly_balances:
                monthly_balances[(payment_date.year, payment_date.month)]['Amount Paid'] += amount_paid
                monthly_balances[(payment_date.year, payment_date.month)]['Transaction Codes'].append(transaction_code)

        # Calculate balances due for each month
        cumulative_balance = 0
        for year_month, details in sorted(monthly_balances.items()):
            cumulative_balance += details['Amount Paid'] - details['Rent Due']
            monthly_balances[year_month]['Balance Due'] = cumulative_balance

        # Create DataFrame for balances due
        balances_due = [{'Year': year,
                         'Month': pd.Timestamp(year=year, month=month, day=1).strftime('%B'),
                         'Rent Due': details['Rent Due'],
                         'Amount Paid': details['Amount Paid'],
                         'Balance Due': details['Balance Due'],
                         'Transaction Codes': ", ".join(details['Transaction Codes'])}  # Join all transaction codes into a single string
                        for (year, month), details in monthly_balances.items()]
        balances_due_df = pd.DataFrame(balances_due)

        # Write the rent statement to an Excel file
        rent_statement_file = 'rent_statement.xlsx'
        with pd.ExcelWriter(rent_statement_file) as writer:
            balances_due_df.to_excel(writer, sheet_name='Rent Statement', index=False)

        print(f'Rent statement generated successfully. Saved as {rent_statement_file}')

    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
mpesa_statement_file = '/Users/user/Desktop/Rent-Statement-Generator/mpesa_statement_file.xlsx'
generate_rent_statement(mpesa_statement_file)
