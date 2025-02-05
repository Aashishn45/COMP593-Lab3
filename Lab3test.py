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
    if len(sys.argv) < 2:
        print("Error: Missing CSV file path")
        sys.exit(1)

    if not os.path.isfile(sys.argv[1]):
        print("Error: Invalid CSV file path")
        sys.exit(2)

    return sys.argv[1]

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    salescsv_path = os.path.abspath(sales_csv)
    salescsv_fol = os.path.dirname(salescsv_path)

    current_date = date.today().isoformat()
    ordersfolder = f"Orders_{current_date}"
    orders_dir = os.path.join(salescsv_fol, ordersfolder)

    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)

    return orders_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    sales_dframe = pd.read_csv(sales_csv)

    # Insert "TOTAL PRICE" column
    sales_dframe.insert(7, "TOTAL PRICE", sales_dframe["ITEM QUANTITY"] * sales_dframe["ITEM PRICE"])

    # Remove unnecessary columns
    sales_dframe.drop(columns=["ADDRESS", "CITY", "STATE", "POSTAL CODE", "COUNTRY"], inplace=True)

    # Group by ORDER ID
    for orderid, order_dframe in sales_dframe.groupby("ORDER ID"):
        order_dframe = order_dframe.copy()  # Prevent SettingWithCopyWarning
        order_dframe.drop(columns=["ORDER ID"], inplace=True)
        order_dframe.sort_values(by='ITEM NUMBER', inplace=True)

        # Append "GRAND TOTAL" row
        grandtot = order_dframe["TOTAL PRICE"].sum()
        grandtot_df = pd.DataFrame({"ITEM PRICE": ['GRAND TOTAL:'], "TOTAL PRICE": [grandtot]})
        order_dframe = pd.concat([order_dframe, grandtot_df], ignore_index=True)

        # Determine file path
        ordernam = f"ORDER_{orderid}.xlsx"
        orders_df_path = os.path.join(orders_dir, ordernam)

        # Save to Excel with formatting
        with pd.ExcelWriter(orders_df_path, engine="xlsxwriter") as ex_write:
            order_dframe.to_excel(ex_write, index=False, sheet_name=str(orderid))

            # Formatting
            work_book = ex_write.book
            work_sheet = ex_write.sheets[str(orderid)]

            # Set column width
            for col_num, value in enumerate(order_dframe.columns.values):
                work_sheet.set_column(col_num, col_num, max(12, len(value) + 2))

            # Apply bold format to header
            header_format = work_book.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            work_sheet.set_row(0, None, header_format)

if __name__ == '__main__':
    main()
