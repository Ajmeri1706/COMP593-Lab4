import pandas as pd
import os
import sys
from datetime import date


def create_orders_folder(csv_file):
    csv_folder = os.path.dirname(csv_file)
    today = date.today().isoformat()
    orders_folder = os.path.join(csv_folder, f"Orders_{today}")

    os.makedirs(orders_folder, exist_ok=True)

    return orders_folder


def create_order_files(csv_file, orders_folder):
    sales_data = pd.read_csv(csv_file)

    for order_id in sales_data["ORDER ID"].unique():
        order_data = sales_data[sales_data["ORDER ID"] == order_id].copy()

        order_data = order_data.sort_values("ITEM NUMBER")

        order_data["TOTAL PRICE"] = (
            order_data["ITEM QUANTITY"] * order_data["ITEM PRICE"]
        )

        grand_total = order_data["TOTAL PRICE"].sum()

        grand_total_row = {column: "" for column in order_data.columns}
        grand_total_row["ITEM PRICE"] = "GRAND TOTAL:"
        grand_total_row["TOTAL PRICE"] = grand_total

        order_data = pd.concat(
            [order_data, pd.DataFrame([grand_total_row])],
            ignore_index=True
        )

        output_file = os.path.join(
            orders_folder,
            f"Order_{order_id}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            order_data.to_excel(writer, sheet_name="Order", index=False)

            workbook = writer.book
            worksheet = writer.sheets["Order"]

            money_format = workbook.add_format({
                "num_format": "$#,##0.00"
            })

            worksheet.set_column("A:A", 11)
            worksheet.set_column("B:B", 13)
            worksheet.set_column("C:C", 15)
            worksheet.set_column("D:D", 15)
            worksheet.set_column("E:E", 15, money_format)
            worksheet.set_column("F:F", 13)
            worksheet.set_column("G:G", 13)
            worksheet.set_column("H:H", 10)
            worksheet.set_column("I:I", 30, money_format)


def main():
    csv_file = "sales_data.csv"
    orders_folder = create_orders_folder(csv_file)
    create_order_files(csv_file, orders_folder)

    print("Order Excel files created successfully.")


main()