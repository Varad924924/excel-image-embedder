import os
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from tkinter import Tk, filedialog

def embed_screenshots_in_excel(excel_path, screenshot_dir):
    """
    Takes an Excel file and a directory of screenshots and embeds the images
    into the 'Screenshot' column of the Excel file, adjusting column and row size.
    """
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at {excel_path}")
        return

    if not os.path.exists(screenshot_dir):
        print(f"Error: Screenshot directory not found at {screenshot_dir}")
        return
    
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        # Adjust column and row dimensions for better image display
        # Set a larger width for the 'Screenshot' column (Column F)
        ws.column_dimensions['C'].width = 40
        
        # Adjusting the width of other columns as well
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 80
        ws.column_dimensions['D'].width = 5
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 50

        # Iterate through rows starting from the second row (to skip headers)
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            # Set a larger height for each row that contains an image
            ws.row_dimensions[row_idx].height = 120

            part_id = row[0].value # Part ID is in the first column
            if part_id is None:
                continue

            screenshot_filename = f"part_{part_id}.jpg"
            screenshot_path = os.path.join(screenshot_dir, screenshot_filename)

            if os.path.exists(screenshot_path):
                try:
                    img = XLImage(screenshot_path)
                    
                    # Increase the size of the images for clarity
                    img.width = 280
                    img.height = 160
                    
                    # Embed the image in the 'Screenshot' column (column F, index 5)
                    ws.add_image(img, f'C{row_idx}')
                except Exception as e:
                    print(f"Error embedding image for part ID {part_id}: {e}")
        
        # Save the modified workbook
        wb.save(excel_path)
        print(f"Successfully embedded images and saved file to: {excel_path}")
        
    except Exception as e:
        print(f"An error occurred while processing the Excel file: {e}")

if __name__ == '__main__':
    # Initialize tkinter and hide the main window
    root = Tk()
    root.withdraw()

    # Open the file dialog for the Excel file
    excel_path = filedialog.askopenfilename(
        title="Select the Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )
    
    # Check if a file was selected
    if not excel_path:
        print("Excel file selection cancelled. Exiting.")
    else:
        # Open the directory dialog for the screenshot folder
        screenshot_dir = filedialog.askdirectory(
            title="Select the Screenshots Folder"
        )

        # Check if a directory was selected
        if not screenshot_dir:
            print("Screenshots directory selection cancelled. Exiting.")
        else:
            # Now, call the function with the selected paths
            embed_screenshots_in_excel(excel_path, screenshot_dir)