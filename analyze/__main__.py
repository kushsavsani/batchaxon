import os
import analyze
import pandas as pd

morph_path = input("Input the path to the morphometrics folder: ")
fourty_path = input("Input the path to the 40x images: ")
four_path = input("Input the path to the 4x images: ")

fourty_df = pd.DataFrame(columns=['img', 'project', 'mag', 'image', 'img_area', 'axon_count', 'gratio_avg', 'diam_avg'])
four_df = pd.DataFrame(columns=['img', 'project', 'mag', 'image', 'img_area'])

for file_name in os.listdir(morph_path):
    file_path = os.path.join(morph_path, file_name)
    project, mag, image = analyze.get_img_chars(file_name)
    axon_count, gratio_avg, diam_avg = analyze.get_avgs(file_path)
    area = analyze.get_40Xarea(os.path.join(fourty_path,file_name.split('.')[0]+'.ome.tif'))
    fourty_df = fourty_df.concat(fourty_df, pd.DataFrame({file_name.split('.')[0], project, mag, image, area, axon_count, gratio_avg, diam_avg}))
    fourty_df.to_excel(os.path.join(morph_path, '40X_morphometrics.xlsx'))

for file_name in os.listdir(four_path):
    file_path = os.path.join(four_path, file_name)
    project, mag, image = analyze.get_img_chars(file_name)
    area = analyze.get_4Xarea(file_path)
    four_df = four_df.concat(four_df, pd.DataFrame({file_name.split('.')[0], project, mag, image, area}))
    four_df.to_excel(os.path.join(morph_path, '4X_morphometrics.xlsx'))