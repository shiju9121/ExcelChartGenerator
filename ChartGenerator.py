import pandas as pd
import os

# Step 0: Brief introduction
print("""
===========================================
 Excel Chart Generator Program
-------------------------------------------
This program lets you:
1. Choose an input Excel file.
2. Select X-axis and Y-axis columns (by number, name, or ranges like 3-5).
3. Pick chart type (bar, pie, line).
4. Save charts into a new Excel file.
5. Prevent overwriting existing files unless you confirm.
6. Generate multiple charts in multiple output files.

Tips:
- Paste the full file path (without quotes, but quotes will be stripped if present).
- You can mix inputs: e.g. '2, Revenue, 4-6'.
===========================================
""")

# Step 1: Ask user for input file path
input_path = input("Enter full path of input Excel file: ").strip()

# Strip quotes if user pasted with them
if input_path.startswith('"') and input_path.endswith('"'):
    input_path = input_path[1:-1]

# Step 2: Check if file exists
if not os.path.exists(input_path):
    raise FileNotFoundError(f"Input file not found: {input_path}")

# Step 3: Read Excel file
df = pd.read_excel(input_path)

# Step 4: Display available columns
print("Available columns:")
for i, col in enumerate(df.columns):
    print(f"{i+1}. {col}")

# Helper function to validate and convert input
def get_col_name(inp):
    inp = inp.strip()
    if inp.isdigit():
        idx = int(inp) - 1
        return df.columns[idx] if 0 <= idx < len(df.columns) else None
    return inp if inp in df.columns else None

while True:
    # Step 5: Ask user for output file name
    output_name = input("Enter output Excel file name (without extension): ").strip()
    output_path = os.path.join(os.path.dirname(input_path), f"{output_name}.xlsx")

    # Step 6: Check if file exists
    if os.path.exists(output_path):
        choice = input(f"File '{output_name}.xlsx' already exists. Overwrite? (yes/no): ").strip().lower()
        if choice != "yes":
            print("Please enter a new file name.")
            continue

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Data', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Data']

            # Step 7: Get X-axis
            x_axis = None
            while x_axis is None:
                x_input = input("Enter column number or name for X-axis: ").strip()
                x_axis = get_col_name(x_input)
                if x_axis is None:
                    print("Invalid X-axis input. Try again.")

            # Step 8: Get Y-axis (multiple allowed, supports ranges like 3-5)
            y_axis_cols = []
            while not y_axis_cols:
                y_input = input("Enter column numbers or names for Y-axis (comma separated, ranges allowed like 3-5): ").split(",")
                for item in y_input:
                    item = item.strip()
                    if "-" in item and item.replace("-", "").isdigit():
                        # Handle ranges like 3-5
                        start, end = item.split("-")
                        if start.isdigit() and end.isdigit():
                            for idx in range(int(start), int(end) + 1):
                                col_name = get_col_name(str(idx))
                                if col_name:
                                    y_axis_cols.append(col_name)
                                else:
                                    print(f"Invalid input in range: {idx}")
                    else:
                        col_name = get_col_name(item)
                        if col_name:
                            y_axis_cols.append(col_name)
                        else:
                            print(f"Invalid input: {item}")
                if not y_axis_cols:
                    print("No valid Y-axis columns entered. Try again.")

            # Step 9: Ask chart type
            chart_type = input("Enter chart type (bar/pie/line): ").strip().lower()
            if chart_type not in ["bar", "pie", "line"]:
                print("Invalid chart type. Defaulting to bar chart.")
                chart_type = "bar"

            # Step 10: Create chart
            chart = workbook.add_chart({'type': 'column' if chart_type == 'bar' else chart_type})

            # Add multiple Y-axis series in same chart
            for col in y_axis_cols:
                chart.add_series({
                    'name':       col,
                    'categories': ['Data', 1, df.columns.get_loc(x_axis), len(df), df.columns.get_loc(x_axis)],
                    'values':     ['Data', 1, df.columns.get_loc(col), len(df), df.columns.get_loc(col)],
                })

            chart.set_title({'name': f"{', '.join(y_axis_cols)} vs {x_axis}"})
            chart.set_x_axis({'name': x_axis})
            chart.set_y_axis({'name': ', '.join(y_axis_cols)})

            worksheet.insert_chart(2, len(df.columns) + 2, chart)

        print(f"Excel file saved as {output_name}.xlsx with embedded chart(s).")

    except PermissionError:
        print(f"Permission denied: '{output_name}.xlsx' is open. Please close it and try again.")
        continue

    # Step 11: Ask if user wants another chart in a new file
    again = input("Do you want to create another chart in a new Excel file? (yes/no): ").strip().lower()
    if again != "yes":
        break
