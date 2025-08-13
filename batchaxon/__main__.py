import os
import analyze_nerve
import pandas as pd
import xlsxwriter

# receive user input
fourx_path = input("Input the path to the 4X spreadsheet: ")
study_dir_path = input("Input the path to the study directory: ")
workbook_path = os.path.join(os.path.join(study_dir_path, os.path.basename(study_dir_path)+' Data.xlsx'))
# workbook = xlsxwriter.Workbook(workbook_path)
# headers = ['Slide Number - Nerve', 'Sample', 'CSA (um2)', 'Total Axons', 'G ratio', 'Axon Diameter', 'Full Axon Count', 'Axons/um2']

# format all the possible cell types for the spreadsheet
# header_format = workbook.add_format({'border':2, 'bold':True})
# bold_format = workbook.add_format({'bold':True})
# grey_format = workbook.add_format({'bg_color':"#8a8a8a", 'pattern':1})

# obtain data from the 4x spreadsheet
fourx_df = pd.read_excel(fourx_path)
fourx_df[['Animal', 'Magnification', 'Nerve']] = fourx_df['Name'].str.split('_', expand=True)

# parse through the study directory and obtain all necessary data
for animal_dir_name in os.listdir(study_dir_path):
    print(f'Getting data for {animal_dir_name}...')
    print(f'Making new worksheet for {animal_dir_name}')
    sheet_data = []
    # worksheet = workbook.add_worksheet(animal_dir_name)
    # current_row = 0
    animal_dir_path = os.path.join(study_dir_path, animal_dir_name)
    for nerve_dir_name in os.listdir(animal_dir_path):
        nerve_dir_path = os.path.join(animal_dir_path, nerve_dir_name)
        if os.path.isdir(nerve_dir_path):
            print(f'   Getting data for {nerve_dir_name}...')
            nerve_data, nerve_agg_xl_data = analyze_nerve.get_nerve_data(nerve_dir_path)
            
            fourx_csa = fourx_df[(fourx_df['Animal'] == animal_dir_name) & (fourx_df['Nerve'] == nerve_dir_name)]['CSA'].iloc[0]*2.173913043
            nerve_label = ''
            if nerve_dir_name[-4:] == 'Prox':
                nerve_label = f'{animal_dir_name} - Proximal'
            elif nerve_dir_name[-4:] == 'Dist':
                nerve_label = f'{animal_dir_name} - Distal'
            else:
                nerve_label = f'{animal_dir_name} - {nerve_dir_name}'
                
            if not sheet_data:
                sheet_data.append({})
            
            sheet_data.append({'Slide Number - Nerve': nerve_label, 'Sample': f'{animal_dir_name}_4x_{nerve_dir_name}', 'CSA (um2)': fourx_csa})
            for entry in nerve_data:
                sheet_data.append({'Sample': entry['name'], 'CSA (um2)': entry['csa'], 'Total Axons': entry['total_axons'], 'G ratio': entry['g-ratio'], 'Axon Diameter': entry['axon_diam'], 'Axons/um2': entry['total_axons'] / entry['csa']})
            sheet_data.append({'Sample': f"Totals (n={nerve_agg_xl_data['total_imgs']})", 'CSA (um2)': nerve_agg_xl_data['csa'], 'Total Axons': nerve_agg_xl_data['total_axons'], 'G ratio': nerve_agg_xl_data['g-ratio'], 'Axon Diameter': nerve_agg_xl_data['axon_diam'], 'Full Axon Count': nerve_agg_xl_data['total_axons'] / nerve_agg_xl_data['csa'] * fourx_csa, 'Axons/um2': nerve_agg_xl_data['total_axons'] / nerve_agg_xl_data['csa']})
            sheet_data.append({})
            
            # place headers
            # for item in range(0,len(headers)):
            #     worksheet.write(current_row, item, headers[item], header_format)
            # current_row += 1
            
            # place 4x image data
            # fourx_csa = fourx_df[(fourx_df['Animal'] == animal_dir_name) & (fourx_df['Nerve'] == nerve_dir_name)]['CSA'].iloc[0]*2.173913043
            # if nerve_dir_name[-4:] == 'Prox':
            #     worksheet.write(current_row, 0, f'{animal_dir_name} - Proximal', bold_format)
            # elif nerve_dir_name[-4:] == 'Dist':
            #     worksheet.write(current_row, 0, f'{animal_dir_name} - Distal', bold_format)
            # else:
            #     worksheet.write(current_row, 0, f'{animal_dir_name} - {nerve_dir_name}', bold_format)
            # worksheet.write(current_row, 1, f'{animal_dir_name}_4x_{nerve_dir_name}')
            # worksheet.write(current_row, 2, fourx_csa)
            # worksheet.write(current_row, 3, '', grey_format)
            # worksheet.write(current_row, 4, '', grey_format)
            # worksheet.write(current_row, 5, '', grey_format)
            # worksheet.write(current_row, 6, '', grey_format)
            # worksheet.write(current_row, 7, '', grey_format)
            # current_row += 1
            
            # first_row = current_row
            # for entry in nerve_data:
            #     worksheet.write(current_row, 1, entry['name'])
            #     worksheet.write(current_row, 2, entry['csa'])
            #     worksheet.write(current_row, 3, entry['total_axons'])
            #     worksheet.write(current_row, 4, entry['g-ratio'])
            #     worksheet.write(current_row, 5, entry['axon_diam'])
            #     worksheet.write(current_row, 6, '', grey_format)
            #     worksheet.write(current_row, 7, entry['total_axons'] / entry['csa'])
            #     current_row += 1
            # last_row = current_row-1
            # worksheet.conditional_format(first_row, 3, last_row, 3, {'type':'2_color_scale', 'min_color':"#ffffff", 'max_color':"#d53c3c"})
            # worksheet.conditional_format(first_row, 4, last_row, 4, {'type':'2_color_scale', 'min_color':"#ffffff", 'max_color':"#2DB133"})
            
            # worksheet.write(current_row, 1, f"Totals (n={nerve_agg_xl_data['total_imgs']})", bold_format)
            # worksheet.write_number(current_row, 2, nerve_agg_xl_data['csa'], bold_format)
            # worksheet.write_number(current_row, 3, nerve_agg_xl_data['total_axons'], bold_format)
            # worksheet.write_number(current_row, 4, nerve_agg_xl_data['g-ratio'], bold_format)
            # worksheet.write_number(current_row, 5, nerve_agg_xl_data['axon_diam'], bold_format)
            # worksheet.write_number(current_row, 6, nerve_agg_xl_data['total_axons']/nerve_agg_xl_data['csa']*fourx_csa, bold_format)
            # worksheet.write_number(current_row, 7, nerve_agg_xl_data['total_axons']/nerve_agg_xl_data['csa'], bold_format)
            # current_row += 4
    final_df = pd.DataFrame(sheet_data)
    headers = ['Slide Number - Nerve', 'Sample', 'CSA (um2)', 'Total Axons', 'G ratio', 'Axon Diameter', 'Full Axon Count', 'Axons/um2']
    final_df = final_df.reindex(columns=headers)
    
    mode = 'a' if os.path.exists(workbook_path) else 'w'
    
    with pd.ExcelWriter(workbook_path, engine='xlsxwriter', mode=mode) as writer:
        final_df.to_excel(writer, sheet_name=animal_dir_name, index=False, startrow=1, header=False)

        workbook = writer.book
        worksheet = writer.sheets[animal_dir_name]

        header_format = workbook.add_format({'border': 2, 'bold': True})
        bold_format = workbook.add_format({'bold': True})
        grey_format = workbook.add_format({'bg_color': "#8a8a8a", 'pattern': 1})
        
        for col_num, value in enumerate(final_df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        worksheet.set_column('A:B', 24)
        worksheet.set_column('C:H', 15)
        
        worksheet.write_blank('D2', None, grey_format)
        worksheet.write_blank('E2', None, grey_format)
        worksheet.write_blank('F2', None, grey_format)
        worksheet.write_blank('G2', None, grey_format)
        worksheet.write_blank('H2', None, grey_format)

        totals_row_index = final_df[final_df['Sample'].str.contains("Totals", na=False)].index[0] + 1 # +1 for header row
        worksheet.set_row(totals_row_index, None, bold_format)
        
        data_rows = len(final_df) - 2
        worksheet.conditional_format(f'D3:D{data_rows}', {'type':'2_color_scale', 'min_color':"#ffffff", 'max_color':"#d53c3c"})
        worksheet.conditional_format(f'E3:E{data_rows}', {'type':'2_color_scale', 'min_color':"#ffffff", 'max_color':"#2DB133"})
    
    # worksheet.set_column(0, 1, 24)
    # worksheet.set_column(2, 7, 15)

# workbook.close()