import mysql.connector
import pdfplumber
import csv
from openpyxl import Workbook

# Path to the PDF file
pdf_path = 'pdf/13-09.pdf'

# MySQL Database configuration
db_config = {
    'host': 'localhost',  # Update with your MySQL host
    'user': 'root',  # Update with your MySQL username
    'password': 'xxxxx',  # Update with your MySQL password
    'database': 'new_sao_ke_data'  # Update with your MySQL database name
}


# Function to connect to MySQL
def connect_to_mysql():
    return mysql.connector.connect(**db_config)


# Function to check if a table exists in the database
def check_table_exists(cursor, table_name):
    cursor.execute(f"SHOW TABLES LIKE '{table_name}';")
    result = cursor.fetchone()
    return result is not None


# Function to create table if it doesn't exist
def create_table_if_not_exists():
    conn = connect_to_mysql()
    cursor = conn.cursor()
    table_name = 'saoke_13_09'

    if not check_table_exists(cursor, table_name):
        create_table_query = '''
        CREATE TABLE IF NOT EXISTS `saoke_13_09` (
            `id` int NOT NULL AUTO_INCREMENT,
            `doc_no` varchar(255) NULL DEFAULT NULL,
            `amount` varchar(255) NULL DEFAULT NULL,
            `details` varchar(255) NULL DEFAULT NULL,
            `date` varchar(255) NULL DEFAULT NULL,
            `name` varchar(255) NULL DEFAULT NULL,
            `balance` varchar(255) NULL DEFAULT NULL,
            PRIMARY KEY (`id`)
        );
        '''
        cursor.execute(create_table_query)
        conn.commit()
        print(f"Table '{table_name}' created.")
    else:
        print(f"Table '{table_name}' already exists.")

    cursor.close()
    conn.close()


# Function to insert transactions into MySQL
def insert_transactions_to_mysql(transactions):
    conn = connect_to_mysql()
    cursor = conn.cursor()

    # Insert each transaction into MySQL
    insert_query = '''
    INSERT INTO saoke_13_09 (date, amount, details)
    VALUES (%s, %s, %s)
    '''
    for transaction in transactions:
        cursor.execute(insert_query, (
            transaction['date'],
            transaction['amount'],
            transaction['details']
        ))
    conn.commit()

    # Close the cursor and connection
    cursor.close()
    conn.close()
    print("Data exported to MySQL.")


# Function to export data to CSV
def export_to_csv(transactions, csv_file_path):
    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=["date", "amount", "details"])
        writer.writeheader()
        for transaction in transactions:
            writer.writerow(transaction)


# Function to export data to Excel using openpyxl
def export_to_excel(transactions, excel_file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"

    # Write the header
    ws.append(["Date", "Amount", "Details"])

    # Write the data
    for transaction in transactions:
        ws.append([transaction['date'], transaction['amount'], transaction['details']])

    # Save the workbook
    wb.save(excel_file_path)


# Function to read PDF and extract transactions
def extract_transactions_from_pdf(pdf_path, max_pages=-1):
    transactions = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if max_pages == -1:
                max_pages = len(pdf.pages)
            if i >= max_pages:
                break  # Stop extraction after reaching the maximum pages
            lists = page.extract_table()
            if i == 1:
                data = lists[1:]
            else:
                data = lists
            for row in data:
                if len(row) == 4:
                    try:
                        amount = int(row[2].replace('.', '').replace(',', ''))
                        transaction_data = {
                            "date": row[1],
                            "amount": amount,
                            "details": row[3].replace('"', ''),
                        }
                        if amount is not None and amount != 0:
                            transactions.append(transaction_data)
                    except ValueError:
                        print(f"Skipping row: {row}")
        print(f"Extracted data from {min(len(pdf.pages), max_pages)} pages.")
    return transactions


# User menu to choose the export format
print("\nSelect export format:")
print("1. MySQL")
print("2. CSV")
print("3. Excel")

choice = input("Enter your choice (1/2/3): ")

# User input for max pages to extract
try:
    max_pages = int(input("Enter the maximum number of pages to extract (leave blank for all): ") or -1)
except ValueError:
    print("Invalid input. Extracting all pages.")
    max_pages = -1

# Extract transactions only after choosing an export option
transactions = extract_transactions_from_pdf(pdf_path, max_pages=max_pages)

if choice == "1":
    create_table_if_not_exists()  # Check and create table if it doesn't exist
    insert_transactions_to_mysql(transactions)

elif choice == "2":
    csv_file_path = 'data/exported_transactions.csv'
    export_to_csv(transactions, csv_file_path)
    print(f"Data exported to CSV at {csv_file_path}.")

elif choice == "3":
    excel_file_path = 'data/exported_transactions.xlsx'
    export_to_excel(transactions, excel_file_path)
    print(f"Data exported to Excel at {excel_file_path}.")

else:
    print("Invalid choice. Exiting.")
