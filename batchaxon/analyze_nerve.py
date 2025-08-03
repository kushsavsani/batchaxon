'''
analyze_nerve.py

The purpose of this chunk of code is to obtain the data from a single nerve folder
input: file path to a nerve
output: aggregate morphometrics sheet for the nerve
        get csa, total axons, g-ratio, axon_diam, est axons
        
logic: morpometrics from each image can be obtained from the respective file in ./{mag}_morphometrics
       csa needs to be extracted from the respective image in ./{mag}_overlay
       total/avg measurements can be obtained from the aggregate file
       
order: first make the aggregate file
       next obtain data from individual files
       place totals after that
'''

import os
import pandas as pd
from overlay_analysis import process_overlay

def make_aggregate_xl(morph_dir_path):
    print(f'      Making aggregate spreadsheet for {os.path.basename(os.path.dirname(morph_dir_path))}')
    agg_df = pd.DataFrame()
    
    for morph_file_name in os.listdir(morph_dir_path):
       img_df = pd.read_excel(os.path.join(morph_dir_path, morph_file_name))
       agg_df = pd.concat([agg_df, img_df])
       
    nerve_dir_path = os.path.dirname(morph_dir_path)
    nerve_dir_name = os.path.basename(nerve_dir_path)
    animal_dir_path = os.path.dirname(nerve_dir_path)
    animal_dir_name = os.path.basename(animal_dir_path)
    
    agg_df_path = os.path.join(animal_dir_path, animal_dir_name+'_'+nerve_dir_name+'.xlsx')
    agg_df.to_excel(agg_df_path)
    return agg_df_path
    
def get_img_data(morph_file_path):
    morph_file_df = pd.read_excel(morph_file_path)
    
    img_data = {
        'total_axons': len(morph_file_df),
        'g-ratio': morph_file_df['gratio'].mean(),
        'axon_diam': morph_file_df['axon_diam'].mean()
    }

    return img_data
    
def get_agg_data(agg_xl_path):
    agg_xl_df = pd.read_excel(agg_xl_path)
    agg_xl_data = {
        'csa':0,
        'total_axons':len(agg_xl_df),
        'g-ratio':agg_xl_df['gratio'].mean(),
        'axon_diam':agg_xl_df['axon_diam'].mean()
    }
    
    return agg_xl_data

def find_mag(subdirs):
    # finds the magnification in the nerve folder based on the naming convention
    mags = [item.split('_')[0] for item in subdirs]
    mags.remove('4x')
    return int(mags[0][:mags[0].find('x')])

def get_nerve_data(nerve_dir_path):
    # this is the function that will do all the processing
    nerve_dir_subdirs = os.listdir(nerve_dir_path)
    nerve_dir_subdirs = [i.lower() for i in nerve_dir_subdirs]
    mag = find_mag(nerve_dir_subdirs)
    
    if mag == 40:
        overlay_dir_path = os.path.join(nerve_dir_path, f'{mag}x_overlay')
    morph_dir_path = os.path.join(nerve_dir_path, f'{mag}x_morphometrics')
    agg_xl_path = make_aggregate_xl(morph_dir_path)
    
    # this chunk of code collects all the data from individual images and puts them in a list of dicts with the following format
    '''
    [
        {'name':IMAGE NAME, 'csa':CROSS SECTIONAL AREA, 'total_axons':AXON_COUNT, 'g-ratio':AVERAGE GRATIO FOR IMAGE, 'axon_diam':AVG AXON DIAM FOR IMAGE}
        {'name':IMAGE NAME, 'csa':CROSS SECTIONAL AREA, 'total_axons':AXON_COUNT, 'g-ratio':AVERAGE GRATIO FOR IMAGE, 'axon_diam':AVG AXON DIAM FOR IMAGE}
    ]
    '''
    nerve_data = []
    img_count = 0
    
    for morph_file_name in os.listdir(morph_dir_path):
        img_name = morph_file_name[:morph_file_name.find('.xlsx')]
        morph_file_path = os.path.join(morph_dir_path, morph_file_name)
        morph_data = get_img_data(morph_file_path)
       
        if mag == 40: 
            overlay_file_path = os.path.join(overlay_dir_path, img_name+'_FO.tif')
            print(f'      Getting overlay data for {morph_file_name}...')
            overlay_data = process_overlay.get_overlay_area(overlay_file_path)
                   
            img_data = {
                'name':img_name,
                'csa':overlay_data * 0.217391**2,
                'total_axons':morph_data['total_axons'],
                'g-ratio':morph_data['g-ratio'],
                'axon_diam':morph_data['axon_diam']*0.217391
            }
        else:
            img_data = {
                'name':img_name,
                'csa':1474560, # NEED A CONVERSION FACTOR FOR 100X MAGNIFICATION
                'total_axons':morph_data['total_axons'],
                'g-ratio':morph_data['g-ratio'],
                'axon_diam':morph_data['axon_diam']*0.217391
            }
        
        nerve_data.append(img_data)
        img_count += 1
    
    # this chunk of data collects the aggregate data
    agg_xl_data = get_agg_data(agg_xl_path)
    agg_xl_data['csa'] = sum(e['csa'] for e in nerve_data)
    agg_xl_data['total_imgs'] = img_count
    
    return nerve_data, agg_xl_data