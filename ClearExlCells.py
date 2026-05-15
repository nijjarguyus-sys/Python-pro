import os
import openpyxl

# --- CONFIGURATION ---
# Use "." if the Excel files are in the exact same folder as this script
TARGET_FOLDER = "C:\\Users\\singhha1\\Downloads\\files"

# List the specific cells you want to wipe clean
CELLS_TO_CLEAR = ["B14", "C14"]


def clear_excel_cells(folder_path, target_cells):
    # Verify the folder actually exists
    if not os.path.exists(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

  
    files_processed = 0

    for filename in os.listdir(folder_path):
        # Scan only for valid Excel sheets and ignore open temporary files
        if filename.endswith(".xlsx") and not filename.startswith("~$"):
            file_path = os.path.join(folder_path, filename)
            files_processed += 1
            print(f"Opening: {filename}...")

            try:
                wb = openpyxl.load_workbook(file_path)

                # Loop through EVERY worksheet tab inside the file
                for sheet in wb.worksheets:
                    for cell in target_cells:
                        sheet[cell].value = None

                wb.save(file_path)
                wb.close()
                print(f"Successfully cleared and saved: {filename}")

            except PermissionError:
                print(
                    f"CRITICAL ERROR: Close '{filename}' in Microsoft Excel before running!"
                )
            except Exception as e:
                print(f"Could not process {filename}. Reason: {e}")

    if files_processed == 0:
        print(f"No '.xlsx' files were found inside the folder '{folder_path}'.")


# Run the program
clear_excel_cells(TARGET_FOLDER, CELLS_TO_CLEAR)
