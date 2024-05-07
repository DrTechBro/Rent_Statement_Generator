# Rent_Statement_Generator
## Overview
The Rent Statement Generator is a Python tool that processes MPESA statements in Excel format to generate monthly rent statements. It tracks monthly payments, calculates balances due, and outputs a comprehensive rent statement in an Excel file. This tool is ideal for landlords and property managers who need to manage rent payments efficiently.

## Installation

### Prerequisites

1. Python 3.6 or higher
2. Pandas library
3. openpyxl library

### Setup

To set up the Rent Statement Generator on your local machine, follow these steps:

### Clone the repository:

Copy this into your terminal
`git clone https://github.com/yourusername/rent-statement-generator.git`

Navigate to the repository directory:
`cd rent-statement-generator`

Install the required Python packages:
`pip install pandas openpyxl`

### Usage

#### Preparing Your Data

Ensure your MPESA statement Excel file contains at least the following columns:

- Date: The date of the transaction
- Amount: The amount paid in the transaction
- Transaction Code: A unique code identifying the transaction
- Running the Script

Execute the script by running:

`python rent_statement_generator.py`

You will need to specify the path to your MPESA statement Excel file. Optionally, you can specify the monthly rent due and the rent due date.

### Output

The script will generate an Excel file named rent_statement.xlsx that includes the following columns:

- Year: The year of the transaction
- Month: The month of the transaction
- Rent Due: The amount of rent due for the month
- Amount Paid: The total amount paid during the month
- Balance Due: The cumulative balance due after payments
- Transaction Codes: A comma-separated list of transaction codes for payments made during the month

## Contributing

Contributions to the Rent Statement Generator are welcome. Here's how you can contribute:

1. Fork the repository.
2. Create a new branch for your feature (git checkout -b feature/AmazingFeature).
3. Commit your changes (git commit -am 'Add some AmazingFeature').
4. Push to the branch (git push origin feature/AmazingFeature).
5. Open a new Pull Request.

## License
Distributed under the MIT License. See LICENSE for more information.

