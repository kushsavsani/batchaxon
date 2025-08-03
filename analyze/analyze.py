import os
import pandas as pd
from scipy.ndimage import gaussian_filter, label, find_objects
import cv2
import numpy as np

def get_40Xarea(img_path):
    img = cv2.imread(img_path, 0)
    blurred_img = gaussian_filter(img, sigma=25)
    mask = blurred_img > np.quantile(blurred_img, 0.75)
    return np.sum(~mask)

def get_4Xarea(img_path):
    img = cv2.imread(img_path, 0)
    blurred_img = gaussian_filter(img, sigma=50)
    mask = blurred_img > np.quantile(blurred_img, 0.25)
    labeled_array, num_features = label(~mask)
    regions = find_objects(labeled_array)
    center_x, center_y = np.array(mask.shape) // 2
    area = 0
    for i, region in enumerate(regions):
        if region is not None:
            x_slice, y_slice = region
            if (x_slice.start <= center_x < x_slice.stop) and (y_slice.start <= center_y < y_slice.stop):
                area = np.sum(labeled_array[x_slice, y_slice] == (i + 1))
                break
    return area

def get_avgs(morph_path):
    morph = pd.read_excel(morph_path)
    axon_count = len(morph)
    gratio_avg = morph['gratio'].mean()
    diam_avg = morph['axon_diam'].mean()
    return axon_count, gratio_avg, diam_avg

def get_img_chars(img_name):
    project = img_name.split('.')[0].split('_')[0]
    mag = img_name.split('.')[0].split('_')[1]
    image = img_name.split('.')[0].split('_')[2]
    return project, mag, image