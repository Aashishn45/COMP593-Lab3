import sys
import os
from datetime import date
import pandas as pd 

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
    sales_dframe = pd.read_csv(sales_csv)

    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_dframe.insert(7, "TOTAL PRICE", sales_dframe["ITEM QUANTITY"] * sales_dframe["ITEM PRICE"])


    # Remove columns from the DataFrame that are not needed
    sales_dframe.drop(columns=["ADDRESS", "CITY", "STATE", "POSTAL CODE", "COUNTRY"], inplace=True)
    

    # Group the rows in the DataFrame by order ID
    for orderid, order_dframe in sales_dframe.groupby("ORDER ID"):


    # For each order ID:
        # Remove the "ORDER ID" column
        order_dframe.drop(columns=["ORDER ID"], inplace= True)

        # Sort the items by item number
        order_dframe.sort_values(by='ITEM NUMBER', inplace= True)
        
        # Append a "GRAND TOTAL" row
        grandtot = order_dframe["TOTAL PRICE"].sum()
        grandtot_df = pd.DataFrame({"ITEM PRICE": ['GRAND TOTAL:'], "TOTAL PRICE": [grandtot]})
        order_dframe = pd.concat([order_dframe, grandtot_df])

        # Determine the file name and full path of the Excel sheet
        ordernam = f"ORDER_{orderid}.xlsx"
        orders_df_path = os.path.join(orders_dir, ordernam)
    

        # Export the data to an Excel sheet
        name_of_sheet = f"{orderid}" 
        order_dframe.to_excel(orders_df_path, index= False, name_of_sheet= name_of_sheet)

       # TODO: Format the Excel sheet
        ex_write = pd.ExcelWriter(orders_df_path, engine= "xlsxwriter")
        order_dframe.to_excel(ex_write, index= False, name_of_sheet=name_of_sheet)
        work_book = ex_write.book
        work_sheet = ex_write.sheets[name_of_sheet]
        

        
    pass

if __name__ == '__main__':
    main()