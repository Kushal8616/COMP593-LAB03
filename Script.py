import os
import sys
import pandas as pd
from datetime import datetime
import re


def validate_args(args):
    if len(args) != 2:
        print("C:\Users\KUSHAL\OneDrive\other computer\COMP593-LAB03\salescsv.py")
        sys.exit(1)
    csv_path = args[1]
    if not os.path.isfile(csv_path):
        print(f"C:\Users\KUSHAL\OneDrive\other computer\COMP593-LAB03\salescsv.py: The file '{csv_path}' does not exist.")
        sys.exit(1)
    return csv_path

def create_orders_directory(csv_path):
    today = datetime.today().strftime('%Y-%m-%d')
    orders_dir = os.path.join(os.path.dirname(csv_path), f"Orders_{today}")
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    return orders_dir

def process_csv(csv_path, orders_dir):
    df = pd.read_csv(csv_path)
    order_ids = df['ORDER ID'].unique()
    
    for order_id in order_ids:
        order_df = df[df['ORDER ID'] == order_id].sort_values(by='ITEM NUMBER')
        order_df['TOTAL PRICE'] = order_df['ITEM QUANTITY'] * order_df['ITEM PRICE']
        total_price_sum = order_df['TOTAL PRICE'].sum()
        
        # Formatting
        order_df['ITEM PRICE'] = order_df['ITEM PRICE'].map('${:,.2f}'.format)
        order_df['TOTAL PRICE'] = order_df['TOTAL PRICE'].map('${:,.2f}'.format)
        
        # Write to Excel
        order_file = os.path.join(orders_dir, f'Order_{order_id}.xlsx')
        with pd.ExcelWriter(order_file, engine='xlsxwriter') as writer:
            order_df.to_excel(writer, sheet_name='Order', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Order']
            money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
            
            for col_num, value in enumerate(order_df.columns.values):
                column_len = max(order_df[value].astype(str).map(len).max(), len(value))
                worksheet.set_column(col_num, col_num, column_len + 2, money_fmt if 'PRICE' in value else None)
                
            worksheet.write(len(order_df) + 1, 0, 'Grand Total')
            worksheet.write(len(order_df) + 1, 4, '${:,.2f}'.format(total_price_sum), money_fmt)

if __name__ == '__main__':
    csv_path = validate_args(sys.argv)
    orders_dir = create_orders_directory(csv_path)
    process_csv(csv_path, orders_dir)
    print(f"Order files have been successfully created in '{orders_dir}'")
