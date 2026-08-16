import tkinter as tk
from pathlib import Path
from textwrap import dedent
from tkinter import filedialog, messagebox

import pandas as pd


# * FUNCTIONS
def choose_file():
    file_path = filedialog.askopenfilename(
        title="Select an Excel or CSV file",
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
    )

    if file_path:
        selected_file.set(file_path)


def choose_save_location():
    output_file = filedialog.asksaveasfilename(
        title="Save MLS Report As",
        defaultextension=".xlsx",
        filetypes=[("Text File", "*.txt"), ("Word Document", "*.docx")],
    )

    if output_file:
        save_location.set(output_file)


def generate_trend_card(df_row):
    """Formats the latest period's trend metrics into clean text."""
    # 1. Median Price Trend
    price_chg = df_row["Median_Price_MoM_%"]
    price_arrow = "↑" if price_chg > 0 else ("↓" if price_chg < 0 else "→")
    price_line = (
        f"Median price {price_arrow} {abs(price_chg):.1f}% compared to previous month"
    )

    # 2. Sales Volume Trend
    vol_chg = df_row["Sales_Volume_MoM_%"]
    vol_arrow = "↑" if vol_chg > 0 else ("↓" if vol_chg < 0 else "→")
    vol_line = (
        f"Sales volume {vol_arrow} {abs(vol_chg):.1f}% compared to previous month"
    )

    # 3. DOM Difference
    dom_diff = df_row["DOM_MoM_Diff"]
    dom_arrow = "↑" if dom_diff > 0 else ("↓" if dom_diff < 0 else "→")
    dom_line = f"DOM {dom_arrow} {abs(dom_diff):.0f} days compared to previous month"

    return (
        f"Market Trend ({df_row['CloseMonth']})\n{price_line}\n{vol_line}\n{dom_line}"
    )


def process_file():
    input_path = selected_file.get()

    if not input_path:
        messagebox.showwarning(
            "No File Selected", "Please choose an Excel or CSV file first."
        )
        return

    input_path = Path(input_path)
    output_path = save_location.get()

    # * MAIN LOGIC
    try:
        # 1. Load data
        if input_path.suffix == ".xlsx":
            df = pd.read_excel(input_path)
        else:
            df = pd.read_csv(input_path)

        # 2. Deduplicate
        df = df.drop_duplicates(subset=["ListingKey"], keep="first")

        # 3. Clean string columns (strip whitespace & standardize casing)
        str_cols = ["City", "StandardStatus", "PropertyType", "ListOfficeName"]
        for col in str_cols:
            df[col] = df[col].astype(str).str.strip().str.title()

        # 4. Standardize phone numbers (extract digits only)
        df["ListAgentDirectPhone_Clean"] = (
            df["ListAgentDirectPhone"].astype(str).str.replace(r"\D", "", regex=True)
        )

        # 5. Parse mixed date formats safely
        date_cols = ["ListingContractDate", "PendingDate", "CloseDate"]
        for col in date_cols:
            df[col] = pd.to_datetime(df[col], format="mixed", errors="coerce")

        # * COLUMNS
        close_price_col = df["ClosePrice"]
        days_on_market = df["DaysOnMarket"]

        # * PRICE & MARKET RESULTS
        closing_price_sum = int(close_price_col.sum())
        avg_sale_price = int(close_price_col.mean())
        median_sale_price = int(close_price_col.median())
        avg_days_on_market = days_on_market.mean()
        median_days_on_market = int(days_on_market.median())

        # Standardize status column to handle casing and whitespace edge cases
        df["StandardStatus_Clean"] = (
            df["StandardStatus"].astype(str).str.strip().str.title()
        )

        # * INVENTORY METRICS
        # 1. Count Active Inventory
        active_inventory = df[
            df["StandardStatus_Clean"].isin(["Active", "Contingent"])
        ].shape[0]

        # 2. Filter for closed listings (DataFrame)
        closed_df = df[df["StandardStatus_Clean"] == "Closed"].copy()

        # 3. Count Closed Sales (Integer count)
        closed_sales_count = closed_df.shape[0]

        # 4. Safely parse CloseDate to datetime objects (handles mixed formats & blanks)
        closed_df["CloseDate"] = pd.to_datetime(
            closed_df["CloseDate"], format="mixed", errors="coerce"
        )

        # 5. Calculate span in days
        min_date = closed_df["CloseDate"].min()
        max_date = closed_df["CloseDate"].max()

        # Check if we have valid dates before calculating
        if pd.notnull(min_date) and pd.notnull(max_date):
            days_span = (max_date - min_date).days
        else:
            days_span = 0

        # Convert days to months (using standard 30.44 days/month average)
        timeframe_months = days_span / 30.44 if days_span > 0 else 1.0

        # Calculate Absorption Rate (Sales per Month)
        absorption_rate = (
            closed_sales_count / timeframe_months if timeframe_months > 0 else 0
        )

        # Calculate Months of Supply
        if absorption_rate > 0:
            months_of_supply = active_inventory / absorption_rate
        else:
            months_of_supply = float(
                "inf"
            )  # Avoid division by zero if 0 sales occurred

        # * AVG $/SQ FT & SALE-TO-LIST RATIOS
        # 1. Convert columns to numeric, coercing bad data to NaN
        closed_df["ClosePrice"] = pd.to_numeric(
            closed_df["ClosePrice"], errors="coerce"
        )

        closed_df["LivingArea"] = pd.to_numeric(
            closed_df["LivingArea"], errors="coerce"
        )

        closed_df["ListPrice"] = pd.to_numeric(closed_df["ListPrice"], errors="coerce")

        closed_df["OriginalListPrice"] = pd.to_numeric(
            closed_df["OriginalListPrice"], errors="coerce"
        )

        # 2. Calculate $/sq ft per row (ignoring zero or missing sq ft to avoid inf)
        valid_sqft_mask = (closed_df["LivingArea"] > 0) & (
            closed_df["ClosePrice"].notnull()
        )
        closed_df.loc[valid_sqft_mask, "Price_Per_SqFt"] = (
            closed_df["ClosePrice"] / closed_df["LivingArea"]
        )

        # 3. Take the mean of individual $/sq ft values
        avg_price_per_sqft = closed_df["Price_Per_SqFt"].mean()

        # 4. Calculate row-level ratios (multiplied by 100 for percentage format)
        # Filter out zero/null list prices to avoid division errors
        valid_mask = (closed_df["ListPrice"] > 0) & (closed_df["ClosePrice"].notnull())

        closed_df.loc[valid_mask, "Sale_To_List_Ratio"] = (
            closed_df["ClosePrice"] / closed_df["ListPrice"]
        ) * 100

        closed_df.loc[valid_mask, "Sale_To_Orig_List_Ratio"] = (
            closed_df["ClosePrice"] / closed_df["OriginalListPrice"]
        ) * 100

        # 6. Aggregate metrics (Average and Median)
        avg_sale_to_list = closed_df["Sale_To_List_Ratio"].mean()
        median_sale_to_list = closed_df["Sale_To_List_Ratio"].median()

        avg_sale_to_orig = closed_df["Sale_To_Orig_List_Ratio"].mean()
        median_sale_to_orig = closed_df["Sale_To_Orig_List_Ratio"].median()

        # * SALES COMPARED TO ASKING PRICE
        # Drop rows with missing price data
        valid_closed = closed_df.dropna(subset=["ClosePrice", "ListPrice"]).copy()
        total_closed_sales = len(valid_closed)

        if total_closed_sales > 0:
            # 4. Categorize sales
            above_asking_count = (
                valid_closed["ClosePrice"] > valid_closed["ListPrice"]
            ).sum()
            at_asking_count = (
                valid_closed["ClosePrice"] == valid_closed["ListPrice"]
            ).sum()
            below_asking_count = (
                valid_closed["ClosePrice"] < valid_closed["ListPrice"]
            ).sum()

            # 5. Calculate percentages
            pct_above = (above_asking_count / total_closed_sales) * 100
            pct_at = (at_asking_count / total_closed_sales) * 100
            pct_below = (below_asking_count / total_closed_sales) * 100

            sales_compared_to_list_price = f"""
            Total Closed Analyzed: {total_closed_sales}
            Above Asking Price: {pct_above:.2f}% ({above_asking_count} homes)
            At Asking Price:    {pct_at:.2f}% ({at_asking_count} homes)
            Below Asking Price: {pct_below:.2f}% ({below_asking_count} homes)
            """

            sales_compared_to_list_price = dedent(sales_compared_to_list_price).strip()

        else:
            sales_compared_to_list_price = "No valid closed sales available to analyze."

            # print("---")
            # print("No valid closed sales available to analyze.")

        # * PRICE DISTRIBUTION
        # Define standard price brackets
        bins = [0, 300000, 500000, 750000, 1000000, float("inf")]
        labels = [
            "Under $300k",
            "$300k - $500k",
            "$500k - $750k",
            "$750k - $1M",
            "$1M+",
        ]

        # Categorize listings
        df["Price_Bracket"] = pd.cut(df["ListPrice"], bins=bins, labels=labels)

        # Generate distribution summary table
        distribution_table = df["Price_Bracket"].value_counts(sort=False).reset_index()
        distribution_table.columns = ["Price Bracket", "Property Count"]
        distribution_table["Percentage"] = (
            distribution_table["Property Count"] / len(df)
        ) * 100

        # print("---")
        # print(distribution_table.to_string(index=False))

        # * MONTHLY TRENDS
        # 1. Create Month and Quarter Period columns
        closed_df["CloseMonth"] = closed_df["CloseDate"].dt.to_period(
            "M"
        )  # e.g., '2025-01', '2025-02'
        closed_df["CloseQuarter"] = closed_df["CloseDate"].dt.to_period(
            "Q"
        )  # e.g., '2025Q1', '2025Q2'

        # 2. Group by Month or Quarter to calculate trends
        monthly_trends = (
            closed_df.groupby("CloseMonth")
            .agg(
                Closed_Sales=("ListingKey", "count"),
                Avg_Close_Price=("ClosePrice", "mean"),
                Median_Close_Price=("ClosePrice", "median"),
                Avg_DOM=("DaysOnMarket", "mean"),
            )
            .reset_index()
        )

        # 3. Ensure it is sorted chronologically
        monthly_trends = monthly_trends.sort_values("CloseMonth").reset_index(drop=True)

        # 4. Calculate percentage change for Price and Volume
        monthly_trends["Median_Price_MoM_%"] = (
            monthly_trends["Median_Close_Price"].pct_change() * 100
        ).round(2)

        monthly_trends["Sales_Volume_MoM_%"] = (
            monthly_trends["Closed_Sales"].pct_change() * 100
        ).round(2)

        # 5. Calculate absolute change for DOM
        monthly_trends["DOM_MoM_Diff"] = monthly_trends["Avg_DOM"].diff().round(2)

        # * Format specific columns directly
        formatted_trends = monthly_trends.copy()

        # Currency columns ($ and commas, 2 decimals)
        currency_cols = ["Avg_Close_Price", "Median_Close_Price"]
        for col in currency_cols:
            formatted_trends[col] = formatted_trends[col].apply(lambda x: f"${x:,.0f}")

        # Numeric/Metric columns (commas, 2 decimals)
        formatted_trends["Avg_DOM"] = formatted_trends["Avg_DOM"].apply(
            lambda x: f"{x:,.1f}"
        )

        # Integer counts (commas only)
        formatted_trends["Closed_Sales"] = formatted_trends["Closed_Sales"].apply(
            lambda x: f"{x:,}"
        )

        formatted_monthly_trends_string = formatted_trends.to_string(index=False)
        # print("---")
        # print(formatted_trends.to_string(index=False))

        # Run for the latest available month
        # Replace NaN with 'N/A' or '-' for reporting
        market_trend_summary = generate_trend_card(monthly_trends.iloc[-1])
        # print("---")
        # print(market_trend_summary)

        # * FORMATTED FINAL REPORT
        raw_report = f"""
        ==================================================
                        MLS MARKET REPORT
        ==================================================
        *** OVERVIEW ***

        Avg Sale Price: ${avg_sale_price:,}
        Median Sale Price: ${median_sale_price:,}
        Sum Of All Sales: ${closing_price_sum:,}
        Avg Days On Market: {avg_days_on_market:.1f}
        Median Days On Market: {median_days_on_market}

        Active Inventory: {active_inventory} listings
        Closed Sales: {closed_sales_count} listings
        Absorption Rate: {absorption_rate:.2f} sales/month
        Months of Supply: {months_of_supply:.2f} months

        *** Sale Price Compared To List Price Data ***

        Average Sale-To-List-Price Ratio: {avg_sale_to_list:.2f}%
        Median Sale-To-List-Price Ratio:  {median_sale_to_list:.2f}%
        Average Sale-To-Original-List-Price Ratio: {avg_sale_to_orig:.2f}%
        Median Sale-To-Original-List-Price Ratio: {median_sale_to_orig:.2f}%
        Average $/SqFt: ${avg_price_per_sqft:.2f}

        {sales_compared_to_list_price}

        *** Monthly Trends ***

        {formatted_monthly_trends_string}

        *** Market Trend Summary ***

        {market_trend_summary}
        """

        # 2. Strip leading whitespace from every line and join
        final_report_string = "\n".join(
            line.lstrip() for line in raw_report.strip().splitlines()
        )

        # * CREATE REPORT
        if not output_path:
            messagebox.showwarning(
                "No Save Location", "Please choose where to save the report."
            )
            return

        with open(output_path, "w") as file:
            file.write(final_report_string)

        messagebox.showinfo(
            "Complete", f"Processing complete!\n\nSaved to:\n{output_path}"
        )

    except Exception as error:
        messagebox.showerror("Error", f"Something went wrong:\n\n{error}")


# * GUI SECTION
window = tk.Tk()
window.title("MLS Report App")
w, h = 600, 300
# Center the window
x = (window.winfo_screenwidth() - w) // 2
y = (window.winfo_screenheight() - h) // 2
window.geometry(f"{w}x{h}+{x}+{y}")

selected_file = tk.StringVar()
save_location = tk.StringVar()

title_label = tk.Label(window, text="Select the file to process:", font=("Arial", 14))
title_label.pack(pady=(25, 10))

choose_button = tk.Button(window, text="Choose File...", command=choose_file)
choose_button.pack()

file_label = tk.Label(window, textvariable=selected_file, wraplength=550)
file_label.pack(pady=10)

save_button = tk.Button(window, text="Save Report As...", command=choose_save_location)
save_button.pack()

save_label = tk.Label(window, textvariable=save_location, wraplength=550)
save_label.pack(pady=10)

process_button = tk.Button(window, text="Process", command=process_file)
process_button.pack(pady=10)

window.mainloop()
