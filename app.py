import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import zipfile
import io
import os
import tempfile


def embed_images_to_excel(excel_bytes, image_zip_bytes):
    """
    Embeds images from a zip file into an Excel workbook using specific layout settings.

    Args:
        excel_bytes (bytes): The content of the .xlsx file.
        image_zip_bytes (bytes): The content of the .zip file containing images.

    Returns:
        tuple: A tuple containing the BytesIO buffer of the modified Excel file
               and an error message string (or None if successful).
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            with zipfile.ZipFile(io.BytesIO(image_zip_bytes), 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            image_folder = temp_dir
            extracted_items = os.listdir(temp_dir)
            if len(extracted_items) == 1 and os.path.isdir(os.path.join(temp_dir, extracted_items[0])):
                image_folder = os.path.join(temp_dir, extracted_items[0])

        except zipfile.BadZipFile:
            return None, "Error: The uploaded file is not a valid ZIP archive."

        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active

        try:
            header = [cell.value for cell in ws[1]]
            id_col_idx = header.index("Property ID") + 1
            ss_col_idx = header.index("Screenshot") + 1
        except (ValueError, IndexError):
            return None, "Error: 'Property ID' or 'Screenshot' column not found in the Excel file's first row."

        # ---  New Sizing Logic Applied Here  ---
        # Set specific column widths as requested
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 80
        ws.column_dimensions[get_column_letter(ss_col_idx)].width = 40  # Dynamically find screenshot column ('C')
        ws.column_dimensions['D'].width = 5
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 50
        # --- End of New Sizing Logic ---

        for row in range(2, ws.max_row + 1):
            # ---  Set Row Height  ---
            ws.row_dimensions[row].height = 120

            prop_id = ws.cell(row=row, column=id_col_idx).value
            if not prop_id:
                continue

            image_name = f"part_{prop_id}.jpg"
            image_path = os.path.join(image_folder, image_name)

            if os.path.exists(image_path):
                try:
                    img = XLImage(image_path)

                    # ---  Set Image Dimensions  ---
                    img.width = 280
                    img.height = 160

                    cell_anchor = f"{get_column_letter(ss_col_idx)}{row}"
                    ws.add_image(img, cell_anchor)

                except Exception as e:
                    st.warning(f"Could not process image for Property ID {prop_id}: {e}")

        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)

        return output_buffer, None


def main():
    """Main function to run the Streamlit app."""
    st.set_page_config(layout="wide", page_title="Excel Image Embedder")
    st.title("Excel Screenshot Embedder")
    st.markdown("This tool automates embedding part screenshots into your 'Properties Info' Excel file.")

    st.header("Step 1: Prepare Your Files")
    st.info(
        "1.  **Excel File**: The `.xlsx` file containing property information.\n"
        "2.  **Screenshots**: Place all your `part_XXXX.jpg` screenshot files into a single folder and **compress it into a ZIP file**."
    )

    st.markdown("---")

    st.header("Step 2: Upload Your Files")
    uploaded_excel = st.file_uploader("Upload the Excel Workbook (.xlsx)", type=["xlsx"])
    uploaded_zip = st.file_uploader("Upload the Screenshots ZIP file (.zip)", type=["zip"])

    st.markdown("---")

    if st.button(" Embed Images into Excel", type="primary"):
        if uploaded_excel is not None and uploaded_zip is not None:
            with st.spinner("Processing... Inserting images into Excel file."):
                excel_bytes = uploaded_excel.getvalue()
                zip_bytes = uploaded_zip.getvalue()

                result_buffer, error_message = embed_images_to_excel(excel_bytes, zip_bytes)

            if error_message:
                st.error(error_message)
            else:
                st.success("Success! Images have been embedded.")

                original_filename = os.path.splitext(uploaded_excel.name)[0]
                new_filename = f"{original_filename}_with_images.xlsx"

                st.download_button(
                    label="Download Modified Excel File",
                    data=result_buffer,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Please upload both the Excel file and the ZIP file.")


if __name__ == "__main__":
    main()