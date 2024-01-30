import sys
import os
from datetime import date

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)



# Get path of sales data CSV file from the command line

def get_sales_csv():
    # Check whether command line parameter provided
    if len(sys.argv) < 2:
        print("Error : Missing CSV file path")
        sys.exit(1)


    # Check whether provide parameter is valid path of file
    if not os.path.isfile(sys.argv[1]):
        print("Error : Invalid CSV file path")
        sys.exit(2)

    return sys.argv[1]

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    salescsv_path = os.path.abspath(sales_csv)
    salescsv_fol = os.path.dirname(salescsv_path)
     
    # Determine the name and path of the directory to hold the order data files
    current_date = date.today().isoformat()
    ordersfolder = f"Orders_{current_date}"
    orders_dir = os.path.join(salescsv_fol, ordersfolder)

    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)

    return orders_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    # Insert a new "TOTAL PRICE" column into the DataFrame
    # Remove columns from the DataFrame that are not needed
    # Group the rows in the DataFrame by order ID
    # For each order ID:
        # Remove the "ORDER ID" column
        # Sort the items by item number
        # Append a "GRAND TOTAL" row
        # Determine the file name and full path of the Excel sheet
        # Export the data to an Excel sheet
        # TODO: Format the Excel sheet
    pass

if __name__ == '__main__':
    main()