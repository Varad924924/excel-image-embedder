import os
import ansa
from ansa import base, constants, utils
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import shutil

def allPartsSelection():
    all_parts = base.CollectEntities(constants.OPTISTRUCT, None, "__PROPERTIES__")
    print("No. of parts: ", len(all_parts))
    
    all_materials = base.CollectEntities(constants.OPTISTRUCT, None, "__MATERIALS__")
    print("No. of materials: ", len(all_materials))
    
    material_dict = {m._id: m._name for m in all_materials}
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    output_dir = "D:\\Automation\\AI Fare\\Properties_Info"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    output_file = os.path.join(output_dir, f"Properties_Info_{timestamp}.xlsx")
    screenshot_dir = os.path.join(output_dir, f"screenshots_{timestamp}")
    if not os.path.exists(screenshot_dir):
        os.makedirs(screenshot_dir)

    # Use shutil.rmtree to clean up any existing screenshots directory
    if os.path.exists(screenshot_dir):
        shutil.rmtree(screenshot_dir)
    os.makedirs(screenshot_dir)

    wb = Workbook()
    ws = wb.active
    ws.title = "Properties Info"

    headers = ["Property ID", "Property Name", "Thickness", "MID", "Material Name", "Screenshot"]
    ws.append(headers)

    base.Or(all_parts)

    for i, part in enumerate(all_parts):
        base.Not(all_parts)
        base.Or(part)
        base.BestView(part)
        
        part_id = part._id
        part_name = part._name
        screenshot_filename = f"part_{part_id}.jpg"
        screenshot_path = os.path.join(screenshot_dir, screenshot_filename)

        try:
            min_x, min_y, min_z, max_x, max_y, max_z = base.BoundBox(part)
            max_dim = max(abs(max_x - min_x), abs(max_y - min_y), abs(max_z - min_z))
            base.ZoomInEnt(part)
        except Exception as e:
            print(f"Error Zooming to part {part_id}: {e}")
            base.ZoomInEnt(part)

        status = utils.SnapShot(screenshot_path, transparent=True, image_size=(640, 480))

        if status != 0:
            print(f"Failed to capture screenshot for part ID: {part_id}")
            screenshot_path = None

        # Get thickness and material info
        values = base.GetEntityCardValues(constants.OPTISTRUCT, part, ("T", "MID", "MID1"))
        thickness = values.get("T", "N/A")
        mid = values.get("MID") or values.get("MID1") or "N/A"
        mname = material_dict.get(mid, "N/A")

        # Append row with placeholder for screenshot
        ws.append([part_id, part_name, thickness, mid, mname, ""])

        # Reset selection and view
        base.Or(all_parts)
        base.ZoomAll()

    # Save Excel file
    wb.save(output_file)
    print(f"Excel file saved to: {output_file}")
    print(f"Screenshots saved to: {screenshot_dir}")

if __name__ == '__main__':
    allPartsSelection()