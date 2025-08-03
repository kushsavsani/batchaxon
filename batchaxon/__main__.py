import os
import analyze_nerve
import pandas as pd
import xlsxwriter

# receive user input
fourx_path = input("Input the path to the 4X spreadsheet: ")
study_dir_path = input("Input the path to the study directory: ")
workbook_path = os.path.join(os.path.join(study_dir_path, os.path.basename(study_dir_path)+' Data.xlsx'))
workbook = xlsxwriter.Workbook(workbook_path)
headers = ['Name', 'CSA', 'total axons', 'g-ratio', 'axon_diam', 'total_est_axons']

# obtain data from the 4x spreadsheet
fourx_df = pd.read_excel(fourx_path)
fourx_df[['Animal', 'Magnification', 'Nerve']] = fourx_df['Name'].str.split('_', expand=True)

# parse through the study directory and obtain all necessary data
for animal_dir_name in os.listdir(study_dir_path):
    print(f'Getting data for {animal_dir_name}...')
    print(f'Making new worksheet for {animal_dir_name}')
    worksheet = workbook.add_worksheet(animal_dir_name)
    current_row = 0
    animal_dir_path = os.path.join(study_dir_path, animal_dir_name)
    for nerve_dir_name in os.listdir(animal_dir_path):
        nerve_dir_path = os.path.join(animal_dir_path, nerve_dir_name)
        if os.path.isdir(nerve_dir_path):
            print(f'   Getting data for {nerve_dir_name}...')
            nerve_data, nerve_agg_xl_data = analyze_nerve.get_nerve_data(nerve_dir_path)
            
            worksheet.write(current_row, 0, f'{animal_dir_name}_{nerve_dir_name}')
            current_row += 1
            
            for item in range(0,len(headers)-1):
                worksheet.write(current_row, item, headers[item])
            current_row += 1
            
            fourx_csa = fourx_df[(fourx_df['Animal'] == animal_dir_name) & (fourx_df['Nerve'] == nerve_dir_name)]['CSA'].iloc[0]
            worksheet.write(current_row, 0, f'{animal_dir_name}_4x_{nerve_dir_name}')
            worksheet.write(current_row, 1, fourx_csa)
            current_row += 1
            
            for entry in nerve_data:
                worksheet.write(current_row, 0, entry['name'])
                worksheet.write(current_row, 1, entry['csa'])
                worksheet.write(current_row, 2, entry['total_axons'])
                worksheet.write(current_row, 3, entry['g-ratio'])
                worksheet.write(current_row, 4, entry['axon_diam'])
                current_row +=1
            
            worksheet.write(current_row, 0, f"total (n={nerve_agg_xl_data['total_imgs']})")
            worksheet.write_number(current_row, 1, nerve_agg_xl_data['csa'])
            worksheet.write_number(current_row, 2, nerve_agg_xl_data['total_axons'])
            worksheet.write_number(current_row, 3, nerve_agg_xl_data['g-ratio'])
            worksheet.write_number(current_row, 4, nerve_agg_xl_data['axon_diam'])
            worksheet.write_number(current_row, 5, nerve_agg_xl_data['total_axons']/nerve_agg_xl_data['csa']*fourx_csa)
            current_row += 4
    worksheet.set_column(0, 0, 24)

workbook.close()