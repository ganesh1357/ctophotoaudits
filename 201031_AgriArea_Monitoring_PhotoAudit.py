#!/usr/bin/env python
# coding: utf-8

# In[65]:


# ####################################################### 
# Photo Audit Program
# Purpose: Generates excel file with all photos to review from a survey
# Date of creation: 26/10/2020 (GR)
# Author: Ganesh Rao (GR)
# Modified: 31/10/2020 (GR)
# #######################################################

# Import relevant libraries (install if necessary)
import os
import sys
import numpy
import xlsxwriter
import pandas as pd
import io
from PIL import Image
from datetime import datetime
from collections import defaultdict

### UPDATE each time
taskdate = "201031"

### UPDATE paths each time

# Choose the SurveyCTO .csv
#csv_path = 'G:/My Drive/CEGIS Drive/4000 Team Folders/4100 OM/410003 ADI/0235 Agriculture/Pilot/Data/Raw/Final/Agriculture Area Survey - Final_WIDE.csv'
dta_path = 'G:/My Drive/CEGIS Drive/4000 Team Folders/4100 OM/410003 ADI/0235 Agriculture/Pilot/Data/Clean/output'
dta_name = "_201029_AE_Clean_final.dta"
input_dta = os.path.join(dta_path, dta_name)

# Choose the corresponding media folder
image_folder_path = 'G:/My Drive/CEGIS Drive/4000 Team Folders/4100 OM/410003 ADI/0235 Agriculture/Pilot/Data/Raw/Final/'

# Set an output file
output_path = 'G:/My Drive/CEGIS Drive/4000 Team Folders/4100 OM/410003 ADI/0235 Agriculture/Pilot/Data/Monitoring/output/9-photoaudits/'
output_name = taskdate + "_PhotoAudit_Report.xlsx"
output_excel = os.path.join(output_path, output_name)

# Read .csv
df = pd.read_stata(input_dta)

df.head()

# Set variables of interest
# Data Collector ID
dc_id = 'dc_id' # Update each time
# Household ID
hh_id = 'uid_check' # Update each time

# Picture Variables (enter relevant)
var_list = defaultdict(list)
var_list_basic = defaultdict(list)

def varlistpure(x):
    subdivs = 11
    crops = 5
    var_values = defaultdict(list)
    for i in range(1,subdivs,1):
        for j in range(1,crops,1):
            # var name, var label for excel
            var_values[f"{x}_{j}_{i}"].append(f"PureCrop_{j}_NSubdiv_{i}")
            # var name, relevant vars to include in the export
            var_values[f"{x}_{j}_{i}"].append(f"{x}_comment_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"purecrop_name_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"c_4_2_cstatus_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"1")
    return var_values

def varlisthorti(x):
    subdivs = 11
    crops = 5
    var_values = defaultdict(list)
    for i in range(1,subdivs,1):
        for j in range(1,crops,1):
            # var name, var label for excel
            var_values[f"{x}_{j}_{i}"].append(f"HortiCrop_{j}_NSubdiv_{i}")
            # var name, relevant vars to include in the export
            var_values[f"{x}_{j}_{i}"].append(f"{x}_comment_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"horticrop_name_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"c_5_3_cstatus_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"2")
    return var_values

def varlistmixedcase1(x, y):
    subdivs = 11
    crops = 5
    var_values = defaultdict(list)
    for i in range(1,subdivs,1):
        for j in range(1,crops,1):
            # var name, var label for excel
            var_values[f"{x}_{j}_{i}"].append(f"Comp_{y}_InterCrop{j}_NSubdiv_{i}")
            # var name, relevant vars to include in the export
            var_values[f"{x}_{j}_{i}"].append(f"{x}_comment_{j}_{i}")
            if y == "1":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3a_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3a_cstatus_{j}_{i}")
            elif y == "2":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3b_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3b_cstatus_{j}_{i}")
            elif y == "3":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3c_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3c_cstatus_{j}_{i}")
            elif y == "4":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3d_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_3d_cstatus_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"3")
    return var_values

def varlistmixedcase2(x, y):
    subdivs = 11
    crops = 5
    var_values = defaultdict(list)
    for i in range(1,subdivs,1):
        for j in range(1,crops,1):
            # var name, var label for excel
            var_values[f"{x}_{j}_{i}"].append(f"Comp_{y}_MixedCrop{j}_NSubdiv_{i}")
            # var name, relevant vars to include in the export
            var_values[f"{x}_{j}_{i}"].append(f"{x}_comment_{j}_{i}")
            if y == "1":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4a_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4a_cstatus_{j}_{i}")
            elif y == "2":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4b_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4b_cstatus_{j}_{i}")
            elif y == "3":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4c_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4c_cstatus_{j}_{i}")
            elif y == "4":
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4d_name_{j}_{i}")
                var_values[f"{x}_{j}_{i}"].append(f"c_6_4d_cstatus_{j}_{i}")
            var_values[f"{x}_{j}_{i}"].append(f"4")
    return var_values

var_list.update(varlistpure("c_4_8"))
var_list.update(varlisthorti("c_5_11"))
var_list.update(varlistmixedcase1("c_6_3a_p", "1"))
var_list.update(varlistmixedcase1("c_6_3b_p", "2"))
var_list.update(varlistmixedcase1("c_6_3c_p", "3"))
var_list.update(varlistmixedcase1("c_6_3d_p", "4"))
var_list.update(varlistmixedcase2("c_6_4ap", "1"))
var_list.update(varlistmixedcase2("c_6_4bp", "2"))
var_list.update(varlistmixedcase2("c_6_4cp", "3"))
var_list.update(varlistmixedcase2("c_6_4dp", "4"))

var_list_basic["map_image"].append("MapPart_1")
var_list_basic["map_image2"].append("MapPart_2")
var_list_basic["map_image3"].append("MapPart_3")
var_list_basic["map_image"].append(f"0")
var_list_basic["map_image2"].append(f"0")
var_list_basic["map_image3"].append(f"0")

def varlistmisc(x):
    subdivs = 11
    var_values = defaultdict(list)
    for i in range(1,subdivs,1):
        # var name, var label for excel
        var_values[f"{x}_{i}"].append(f"GeotraceObstacle_NSubdiv_{i}")
        # var name, relevant vars to include in the export
        var_values[f"{x}_{i}"].append(f"1")
    return var_values

var_list_basic.update(varlistmisc("obstruct_photo"))
# var_list_basic
# for i, v in var_list.items():
#    print(v[1])

# Keep signature vars out - d11, d13

# Make the date readable (all CTO surveys have a "today" var)

## Uncomment next 4 lines if using csv as data input
# df['today']
# df['today2'] = pd.Series(df['today'].apply(lambda x: x.replace('/',' ')))
# df['today2'] = pd.Series(df['today2']).apply(lambda x: datetime.strptime(x, '%d %m %Y'))
# date = 'today2'
date = 'today'

# OPTIONAL: Set a minimum date of interest in the '%b (Oct) %d (1) %Y'(2018) format
# date_min = datetime.strptime("Oct 22 2018",'%b %d %Y') # Update each time

# OPTIONAL: Keep if row date > than min date
# df = df[df[date] >= date_min]

# OPTIONAL: Keep if the DC == X
# df = df[df.dc_id == 101]

# Set Excel workbook
workbook = xlsxwriter.Workbook(output_excel, {'constant_memory': True, 'nan_inf_to_errors': True})
# Add a format to use for the first row (bold, wrap text, thin borders)
first_format = workbook.add_format({'font_name': 'Arial','border': True, 'text_wrap': True, 'align': 'left', 'valign': 'top', 'bold': True})
# Add a format for subsequent rows (thin borders)
pre_format   = workbook.add_format({'font_name': 'Arial','border': True, 'text_wrap': True, 'align': 'left', 'valign': 'top', 'font_color': 'red'})
# Add a format for subsequent rows (thin borders)
all_format   = workbook.add_format({'font_name': 'Arial','border': True, 'text_wrap': True, 'align': 'left', 'valign': 'top'})
# Add a date format (thin borders, date)
date_format  = workbook.add_format({'font_name': 'Arial','border': True, 'text_wrap': True, 'align': 'left', 'valign': 'top', 'font_color': 'red', 'num_format' : 'yyyy-mm-dd'})

# Set prefill dropdown lists to use in the Excel
yesno    = ['1 - Yes', '0 - No']
clarity  = ['1 - Not clear at all', '2 - Somewhat clear', '3 - Clear']
yesnopar = ['1 - Yes', '0 - No', '2 - Only few']
obskind  = ['None observed', '1 - Trees', '2 - Rock -- Boulders', '3 - Bushes -- Thorns', '4 - Other crops', '5 - Canals', '6 - Pit Hole', '7 - Fence -- Wall', '8 - Construction', '87 - Others', '0 - Unable to judge']
obssize  = ['0 - Unable to judge', '1 - Small', '2 - Medium', '3 - Large']
crop     = ['1 - Pure Agri', '2 - Pure Horti', '3 - Intercropping with rows -- crops in separate rows', '4 - Mixed cropping -- mult crops in same row', '5 - Mixed cropping -- no rows', '0 - Unable to judge']

# Create program to resize pictures when outputting

# Option 1: Reduce file size internally
def get_resized_image_data(image_path, bound_width_height):
    # get the image and resize it
    im = Image.open(image_path)
        # Rotate picture to the usual portrait mode
        # im = im.rotate(90)
    im.thumbnail(bound_width_height, Image.ANTIALIAS)  # ANTIALIAS is important if shrinking
    # stuff the image data into a bytestream that excel can read
    im_bytes = io.BytesIO()
    im.save(im_bytes, format='JPEG')
    return im_bytes

# Option 2: Reduce picture scale (not image size)
#def calculate_scale(image_folder_path, bound_size):
    # check the image size without loading it into memory
    #im = Image.open(image_folder_path)
    #original_width, original_height = im.size

    # calculate the resize factor, keeping original aspect and staying within boundary
    #bound_width, bound_height = bound_size
    #ratios = (float(bound_width) / original_width, float(bound_height) / original_height)
    #return min(ratios)

# Loop through var list
for i, v in var_list.items():
    
    # Print the name for reference
    print(i, v[0])
    
    if i in df:
       
        # keep only if the pictures are not missing
        dfi = df[df[i]!='']

        # start only if pictures are present
        if dfi[i].empty==False:

            # OPTIONAL: Randomly keep a fraction of the sample (between 0 and 1, e.g 0.5 = 50%)
            dfi = dfi.sample(frac=1) #100% of photos for the first time

            # Sort by date and DC, this got affected when drawing a sub-sample
            dfi = dfi.sort_values([date, dc_id])

            # Re-Index dataframe to avoid blank lines on Excel, delete old index
            dfi = dfi.reset_index()
            del dfi['index']

            # Create worksheet
            print(i)
            worksheet = workbook.add_worksheet(v[0])

            # Write first row of Excel, with format
            worksheet.write(0,0, 'Photo', first_format)
            worksheet.write(0,1, 'Photo Name', first_format)
            worksheet.write(0,2, 'Date', first_format)
            worksheet.write(0,3, 'DC ID', first_format)
            worksheet.write(0,4, 'UID', first_format)
            worksheet.write(0,5, 'Crop Name', first_format)
            worksheet.write(0,6, 'Crop Status', first_format)
            worksheet.write(0,7, 'Photo Comment', first_format)
            # Add standard questions 
            worksheet.write(0,8, 'Auditor Name', first_format)
            worksheet.write(0,9, 'Audit Date', first_format)
            worksheet.write(0,10, 'Rate clarity of the photo', first_format)
            worksheet.write(0,11, 'What type of crop is this?', first_format)
            worksheet.write(0,12, 'Does the photo match the crop name entered?', first_format)
            worksheet.write(0,13, 'Does the photo show the crop status entered?', first_format)
            worksheet.write(0,14, 'Does the comment explain mismatch or status issues?', first_format)
            worksheet.write(0,15, 'Any other comments?', first_format)                

            # Freeze first row of each worksheet
            worksheet.freeze_panes(1,1)

            # Widen the columns to make the text clearer
            worksheet.set_column('A:A', 60)
            worksheet.set_column('B:Z', 25)

            # Set image size for xlsxwriter
            bound_width_height = (400, 400)

            # Loop over all the pictures of a given BP and insert in Excel
            for x in dfi.index:
                # Set row height
                worksheet.set_row(x+1, 240)

                # Define a line-by-line image path
                image_path = os.path.join(image_folder_path, str(dfi.loc[x,i]))

                # Use image re-sizing program specified above
                image_data = get_resized_image_data(image_path, bound_width_height)

                # resize_scale = calculate_scale(image_path, bound_width_height) 'x_scale': resize_scale, 'y_scale': resize_scale

                # Set format (thin border) to picture cells
                worksheet.write_blank(x+1, 0, None, all_format)

                # Insert relevant image in each row, with correct size
                worksheet.insert_image(x+1, 0, image_folder_path + str(dfi.loc[x,i]), {'image_data': image_data})

                # Write the name of the picture in each row (e.g. media/12345.jpg)
                worksheet.write(x+1, 1, dfi.loc[x,i], all_format)
                # Write the date for each picture (good for reference)
                worksheet.write(x+1, 2, dfi.loc[x,date], date_format)
                # Write the DC ID for each picture (good for reference)
                worksheet.write(x+1, 3, dfi.loc[x,dc_id], all_format)
                # Write the UID associated with each picture
                worksheet.write(x+1, 4, dfi.loc[x,hh_id], all_format)
                # Write the name of the crop
                worksheet.write(x+1, 5, dfi.loc[x,v[2]], all_format)
                # Write the crop status
                worksheet.write(x+1, 6, dfi.loc[x,v[3]], all_format)
                # Write the comment on the photo
                worksheet.write(x+1, 7, dfi.loc[x,v[1]], all_format)

                worksheet.write_blank(x+1, 8, None, all_format)
                
                worksheet.write_blank(x+1, 9, None, all_format)
                worksheet.data_validation(x+1, 9, x+1, 9, {'validate': 'date',
                                                          'criteria': 'greater than or equal to',
                                                          'value': '=TODAY()',
                                                          'error_type':'warning'})
                
                worksheet.write_blank(x+1, 10, None, all_format)
                worksheet.data_validation(x+1, 10, x+1, 10,  {'validate': 'list',
                                                              'source': clarity,
                                                              'error_type':'warning'})
                 
                worksheet.write_blank(x+1, 11, None, all_format)
                worksheet.data_validation(x+1, 11, x+1, 11,  {'validate': 'list',
                                                              'source': crop,
                                                              'error_type':'warning'})
                
                worksheet.write_blank(x+1, 12, None, all_format)
                worksheet.data_validation(x+1, 12, x+1, 12,  {'validate': 'list',
                                                              'source': yesno,
                                                              'error_type':'warning'})
                
                worksheet.write_blank(x+1, 13, None, all_format)
                worksheet.data_validation(x+1, 13, x+1, 13,  {'validate': 'list',
                                                              'source': yesno,
                                                              'error_type':'warning'})
                
                worksheet.write_blank(x+1, 14, None, all_format)
                worksheet.data_validation(x+1, 14, x+1, 14,  {'validate': 'list',
                                                              'source': yesno,
                                                              'error_type':'warning'})
                
                worksheet.write_blank(x+1, 15, None, all_format)

# Loop through var list
for i, v in var_list_basic.items():

    # Print the name for reference
    print(i, v[0])

    if i in df:

        # keep only if the pictures are not missing
        dfi = df[df[i]!='']

        # start only if pictures are present
        if dfi[i].empty==False:

            # OPTIONAL: Randomly keep a fraction of the sample (between 0 and 1, e.g 0.5 = 50%)
            dfi = dfi.sample(frac=1) #100% of photos for the first time

            # Sort by date and DC, this got affected when drawing a sub-sample
            dfi = dfi.sort_values([date, dc_id])

            # Re-Index dataframe to avoid blank lines on Excel, delete old index
            dfi = dfi.reset_index()
            del dfi['index']

            # Create worksheet
            print(i)
            worksheet = workbook.add_worksheet(v[0])

            # Write first row of Excel, with format
            worksheet.write(0,0, 'Photo', first_format)
            worksheet.write(0,1, 'Photo Name', first_format)
            worksheet.write(0,2, 'Date', first_format)
            worksheet.write(0,3, 'DC ID', first_format)
            worksheet.write(0,4, 'UID', first_format)
            # Add standard questions 
            worksheet.write(0, 5, 'Auditor Name', first_format)
            worksheet.write(0, 6, 'Audit Date', first_format)
            worksheet.write(0, 7, 'Rate clarity of the photo', first_format)
            if v[1] == "0":
                worksheet.write(0, 8, 'Have land marks been added?', first_format)
                worksheet.write(0, 9, 'Have the original subdivisions been marked on the map?', first_format)
                worksheet.write(0, 10, 'Have the new subdivision names been marked on the map?', first_format)
                worksheet.write(0, 11, 'Is the mapping from the original to the new subdivision clear?', first_format)
                worksheet.write(0, 12, 'Any other comments?', first_format)
            elif v[1] == "1":
                worksheet.write(0, 8, 'Describe the kind of obstructions seen', first_format)
                worksheet.write(0, 9, 'Describe the size of the obstructions seen', first_format)
                worksheet.write(0, 10, 'Any other comments?', first_format)

            # Freeze first row of each worksheet
            worksheet.freeze_panes(1,1)

            # Widen the columns to make the text clearer
            worksheet.set_column('A:A', 60)
            worksheet.set_column('B:Z', 25)

            # Set image size for xlsxwriter
            bound_width_height = (400, 400)

            # Loop over all the pictures of a given BP and insert in Excel
            for x in dfi.index:
                # Set row height
                worksheet.set_row(x+1, 240)

                # Define a line-by-line image path
                # image_path = image_folder_path + str(dfi.loc[x,i])
                image_path = os.path.join(image_folder_path, str(dfi.loc[x,i]))

                # Use image re-sizing program specified above
                image_data = get_resized_image_data(image_path, bound_width_height)

                # resize_scale = calculate_scale(image_path, bound_width_height) 'x_scale': resize_scale, 'y_scale': resize_scale

                # Set format (thin border) to picture cells
                worksheet.write_blank(x+1, 0, None, all_format)

                # Insert relevant image in each row, with correct size
                worksheet.insert_image(x+1, 0, image_folder_path + str(dfi.loc[x,i]), {'image_data': image_data})

                # Write the name of the picture in each row (e.g. media/12345.jpg)
                worksheet.write(x+1, 1, dfi.loc[x,i], all_format)
                # Write the date for each picture (good for reference)
                worksheet.write(x+1, 2, dfi.loc[x,date], date_format)
                # Write the DC ID for each picture (good for reference)
                worksheet.write(x+1, 3, dfi.loc[x,dc_id], all_format)
                # Write the UID associated with each picture
                worksheet.write(x+1, 4, dfi.loc[x,hh_id], all_format)

                worksheet.write_blank(x+1, 5, None, all_format)

                worksheet.write_blank(x+1, 6, None, all_format)
                worksheet.data_validation(x+1, 6, x+1, 6, {'validate': 'date',
                                                          'criteria': 'greater than or equal to',
                                                          'value': '=TODAY()',
                                                          'error_type':'warning'})

                worksheet.write_blank(x+1, 7, None, all_format)
                worksheet.data_validation(x+1, 7, x+1, 7,  {'validate': 'list',
                                                          'source': clarity,
                                                          'error_type':'warning'})

                if v[1] == "0":
                    worksheet.write_blank(x+1, 8, None, all_format)
                    worksheet.data_validation(x+1, 8, x+1, 8, {'validate': 'list',
                                                              'source': yesnopar,
                                                              'error_type':'warning'})
                    worksheet.write_blank(x+1, 9, None, all_format)
                    worksheet.data_validation(x+1, 9, x+1, 9, {'validate': 'list',
                                                              'source': yesnopar,
                                                              'error_type':'warning'})
                    worksheet.write_blank(x+1, 10, None, all_format)
                    worksheet.data_validation(x+1, 10, x+1, 10, {'validate': 'list',
                                                                  'source': yesnopar,
                                                                  'error_type':'warning'})
                    worksheet.write_blank(x+1, 11, None, all_format)
                    worksheet.data_validation(x+1, 11, x+1, 11, {'validate': 'list',
                                                                  'source': clarity,
                                                                  'error_type':'warning'})
                    worksheet.write_blank(x+1, 12, None, all_format)
                elif v[1] == "1":
                    worksheet.write_blank(x+1, 8, None, all_format)
                    worksheet.data_validation(x+1, 8, x+1, 8, {'validate': 'list',
                                                              'source': obskind,
                                                              'error_type':'warning'})
                    worksheet.write_blank(x+1, 9, None, all_format)
                    worksheet.data_validation(x+1, 9, x+1, 9, {'validate': 'list',
                                                              'source': obssize,
                                                              'error_type':'warning'})
                    worksheet.write_blank(x+1, 10, None, all_format)

# close workbook
workbook.close()


# In[ ]:




