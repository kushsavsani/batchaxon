'''
process_overlay: program to process the 40X overlay images
overlay images are manually created in Fiji
the program opens the image in Fiji and exports the overlay data to a temp file, then deletes the file after deleting it
requirement: the program saves the path to the user's Fiji executable in a text file. if the program is not working, check ot make sure that file path is correct
'''

# import necessary dependencies
import os
import subprocess
import roifile
from PIL import Image, ImageDraw
import numpy as np

# set the path for the hard coded macro that creates the overlay roi file and where that roi file will be saved
macro_path = r"batchaxon\overlay_analysis\export_roi.ijm"
roi_path = r"batchaxon\overlay_analysis\temp_outline.roi"

# get the file path to the fiji executable
def get_fiji_path():
    fiji_executable_path = open(r"batchaxon\overlay_analysis\fiji_executable_path.txt", 'r').read()[:-1]
    if fiji_executable_path == '':
        fiji_executable_path = input('Enter the path to your Fiji executable: ')
        with open(r"batchaxon\overlay_analysis\fiji_executable_path.txt", 'w') as f: print(fiji_executable_path, file=f)
    return fiji_executable_path

# run the fiji macro and return True if it successfully runs
def run_fiji_automation(img_path):
    fiji_executable_path = get_fiji_path()
    if os.path.exists(roi_path):
        os.remove(roi_path)
    command_args = f"{img_path},{roi_path}"
    command = [fiji_executable_path, "-batch", macro_path, command_args]
    try:
        subprocess.run(command, check=True, timeout=90)
    except Exception as e:
        print(f'Error during Fiji execution: {e}')
        return False
    if not os.path.exists(roi_path):
        print('Python check failed: Could not find the output ROI file.')
        return False
    return True

# get the area of the overlay from the temp roi file, then delete the file
def get_overlay_area(img_path):
    run_fiji_automation(img_path)    
    
    roi = roifile.ImagejRoi.fromfile(roi_path)
    with Image.open(img_path) as img:
        image_width, image_height = img.size
    mask = Image.new('L', (image_width, image_height), 0)
    draw = ImageDraw.Draw(mask)

    coords = [(x, y) for x, y in roi.coordinates()]
    draw.polygon(coords, outline=1, fill=1)

    mask_array = np.array(mask)
    pixel_count = np.count_nonzero(mask_array)
    
    os.remove(roi_path)
    
    return pixel_count