import os
import pandas as pd

# Read the first Excel file
data = pd.read_excel(r"path-to-inventory file")

# Count occurrences of each unique item
item_counts = data["Item"].value_counts().reset_index()
item_counts.columns = ["Item Code", "Count"]

# Ensure the 'Item Code' column is a string
item_counts["Item Code"] = item_counts["Item Code"].astype(str)

# Read the second Excel file for item descriptions
descriptions = pd.read_excel(r"path-to-item no-mapping-to-item-descriptions")

# Ensure the 'Item Code' column in descriptions is a string
descriptions["Item Code"] = descriptions["Item Code"].astype(str)

# Merge the counts with descriptions
merged_data = pd.merge(item_counts, descriptions, on="Item Code", how="left")

# Drop unnecessary columns and sort by count (greatest to least)
final_output = merged_data[["Item Code", "Count", "Item Description"]]
final_output = final_output.sort_values(by="Count", ascending=False)

# Save the result to the specified location on Desktop
output_path = r"path-where-output-file-will-be"
final_output.to_excel(output_path, index=False)

print(f"Output saved to '{output_path}'.")

# Prompt user for the number of items
num_items = int(input("Enter the number of top items to create text files and final sheet for: "))

# Prompt user for the subinventory value
subinventory = input("Enter the subinventory value to include in the final template: ")

# Extract the top N items
top_items = final_output.head(num_items)

# Create a folder named "serials" on the Desktop
serials_folder = r"path-to-output-folder-for-serials"
os.makedirs(serials_folder, exist_ok=True)

# Prepare a list for the final Excel sheet
final_serials_list = []

# Loop through the selected top items to create .txt files
for _, row in top_items.iterrows():
    item_code = row["Item Code"]
    item_description = row["Item Description"]
    
    # Replace invalid characters in the filename
    sanitized_description = "".join(c for c in item_description if c.isalnum() or c in " -_").strip()
    file_name = f"{item_code}_{sanitized_description}.txt"
    file_path = os.path.join(serials_folder, file_name)
    
    # Extract serial numbers for the current item
    serials = data[data["Item"] == int(item_code)]["Serial"].dropna().astype(str).tolist()
    
    # Write the serial numbers to the .txt file (only as-is, no variations)
    with open(file_path, "w") as f:
        for serial in serials:
            f.write(serial + "\n")
    
    # Add serial numbers to the final list with the required format
    for serial in serials:
        if serial.startswith("s"):
            final_serials_list.append({"SERIAL_NUMBER": serial[1:], "TO_SUBINVENTORY": subinventory})
            final_serials_list.append({"SERIAL_NUMBER": serial, "TO_SUBINVENTORY": subinventory})
            final_serials_list.append({"SERIAL_NUMBER": f"S{serial[1:]}", "TO_SUBINVENTORY": subinventory})
        elif serial.startswith("S"):
            final_serials_list.append({"SERIAL_NUMBER": serial[1:], "TO_SUBINVENTORY": subinventory})
            final_serials_list.append({"SERIAL_NUMBER": f"s{serial[1:]}", "TO_SUBINVENTORY": subinventory})
            final_serials_list.append({"SERIAL_NUMBER": serial, "TO_SUBINVENTORY": subinventory})
        else:
            final_serials_list.append({"SERIAL_NUMBER": serial, "TO_SUBINVENTORY": subinventory})
            final_serials_list.append({"SERIAL_NUMBER": f"s{serial}", "TO_SUBINVENTORY": subinventory})
            final_serials_list.append({"SERIAL_NUMBER": f"S{serial}", "TO_SUBINVENTORY": subinventory})

print(f"Serial files created in '{serials_folder}'.")

# Convert the final list to a DataFrame and add empty columns for the template
final_serials_df = pd.DataFrame(final_serials_list)
columns = ["SERIAL_NUMBER", "ASSET_TAG", "REFERENCE", "TO_SUBINVENTORY", "ASSIGNED_TO_USER",
           "SUBLOCATION1", "SUBLOCATION2", "SUBLOCATION3", "SUBLOCATION4", "SUBLOCATION5",
           "SUBLOCATION6", "SUBLOCATION7", "END_DATE", "REQUESTER"]
for col in columns:
    if col not in final_serials_df.columns:
        final_serials_df[col] = ""

# Reorder the columns to match the template
final_serials_df = final_serials_df[columns]

# Save the final sheet to an Excel file on the desktop
final_excel_path = r"this-is-the-output-template-path"
final_serials_df.to_excel(final_excel_path, index=False)

print(f"Final serials template saved to '{final_excel_path}'.")
