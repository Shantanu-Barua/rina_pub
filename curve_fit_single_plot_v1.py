### Curve fitting for single plot in a single raw file

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import re
from io import BytesIO
import win32com.client

### Folders
input_folder = 'source/single'
output_folder = 'final/single'
tags_folder = 'tags'

### Make required folders
os.makedirs(output_folder, exist_ok=True)
os.makedirs(input_folder, exist_ok=True)
os.makedirs(tags_folder, exist_ok=True)

### Create output folders if they don't exist
output_file = os.path.join(output_folder, 'single_results.xlsx')

### CSV file for R
"""-----------------------------------------------------------------------"""
csv_output_file = os.path.join(output_folder, 'single_results.csv')

if os.path.exists(csv_output_file): # If CSV exists already, remove it
    os.remove(csv_output_file)

r_data = []
"""-----------------------------------------------------------------------"""

### Variations in estimate limiting
def calculate_estimate_limit1(ci_value): # Standard
    if 0 <= ci_value <= 200:
        return 1
    elif 201 <= ci_value <= 300:
        return 2
    elif 301 <= ci_value <= 400:
        return 2
    elif ci_value > 400:
        return 3
    else:
        return None

def calculate_estimate_limit2(ci_value): # Variation 1, suggested by Stephen
    if 0 <= ci_value <= 100:
        return 1
    elif 101 <= ci_value <= 300:
        return 2
    elif 301 <= ci_value <= 400:
        return 2
    elif ci_value > 400:
        return 3
    else:
        return None


### Refresh excel files in source folder
### Start Excel application
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Set True if you want to see Excel opening

### Loop through each file in the folder
for file in os.listdir(input_folder):
    if file.endswith(".xlsx"):
        file_path = os.path.abspath(os.path.join(input_folder, file))
        # print(f"Processing: {file}")

        try:
            wb = excel.Workbooks.Open(file_path)
            wb.RefreshAll()

            excel.CalculateUntilAsyncQueriesDone()
            excel.CalculateFullRebuild()  # Equivalent to Ctrl+Alt+Shift+F9

            wb.Save()
            wb.Close()
            # print(f"Recalculated and saved: {file}")

        except Exception as e:
            print(f"Error processing {file}: {e}")

### Quit Excel application
excel.Quit()

### Use context manager so file always closes properly
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:

    ### Create a loop to open files in source folder
    for file in os.listdir(input_folder):
        if file.endswith('.xlsx') and '_' in file:
            file_path = os.path.join(input_folder, file)

            try:
                ### Load the Excel file and the 'Measurements' sheet
                df_raw = pd.read_excel(file_path, sheet_name='Measurements', header=None)

                # print(df_raw.iloc[12: , 8:12])

                #### Extract identifier (e.g., 117 from "2025-06-09-0857_117T")
                filename_parts = file.split('_')
                if len(filename_parts) > 1:
                    id_and_rest = filename_parts[1]
                    id_candidate = id_and_rest.split(' ')[0].split('.')[0]
                    match = re.search(r'^(\d+[TMB]?|\d+-\d+)', id_candidate)
                    if match:
                        raw_id = match.group(0)
                    else:
                        print(f"Skipping file {file}: Could not extract numeric identifier. (e.g., 101M, 101-110).")
                        continue
                else:
                    print(f"Skipping file {file}: Filename format unexpected (no underscore found).")
                    continue

                ### Extract data
                r_o2 = pd.to_numeric(df_raw.iat[1, 2], errors='coerce') # Extract O2 from C3 (row=3, col=3)
                ### Extract A, Ci, tleaf, patm from data starting row 16
                a_values = pd.to_numeric(df_raw.iloc[15:, 9], errors='coerce').dropna().values
                ci_values = pd.to_numeric(df_raw.iloc[15:, 11], errors='coerce').dropna().values
                tleaf_values = pd.to_numeric(df_raw.iloc[15:, 19], errors='coerce').dropna().values
                patm_values = pd.to_numeric(df_raw.iloc[15:, 64], errors='coerce').dropna().values

                ### Prepare to write CSV file for R
                # Ensure same length across arrays
                n = min(len(a_values), len(ci_values), len(tleaf_values), len(patm_values))

                df_r_temp = pd.DataFrame({
                    "Obs": np.arange(1, n+1),
                    "Plant": raw_id,
                    "Rep": 1,
                    "Photo": a_values[:n],
                    "Ci": ci_values[:n],
                    "Tleaf": tleaf_values[:n],
                    "Press": patm_values[:n]
                    
                })

                r_data.append(df_r_temp)

                ### Use average Tleaf, patm from measurement section, in case of null defualt value is set
                r_t_leaf = np.mean(tleaf_values) if not np.isnan(np.mean(tleaf_values)) else 28.0
                r_patm = np.mean(patm_values) if not np.isnan(np.mean(patm_values)) else 101
                r_o2 = r_o2 if not np.isnan(r_o2) else 21.0
                # print(r_t_leaf, r_patm , r_o2 )

                raw_df = pd.DataFrame({'r_a': a_values, 'r_ci': ci_values}) # Assemble input data into a dataframe
                # print(raw_df)

                # Convert to numeric safely
                raw_df['r_ci'] = pd.to_numeric(raw_df['r_ci'], errors='coerce')
                raw_df['r_a'] = pd.to_numeric(raw_df['r_a'], errors='coerce')

                # Drop rows with NaNs
                raw_df = raw_df.dropna(subset=['r_ci', 'r_a'])

                raw_df = raw_df.sort_values(by='r_ci').reset_index(drop=True) # Sort a an ci accending for better curve
                # print(raw_df)

                ### Raw plot data Constants
                t_leaf = r_t_leaf # °, collected from raw
                patm = r_patm # kPa, collected from raw
                o2 = r_o2 # kPa, collected from raw

                ### Check later for meaning, For fitting/constants section
                ### **Constant section**
                ## Constants
                c_kc = 35.9774
                ha_kc = 80.99
                
                c_ko = 12.3772
                ha_ko = 23.72
                
                c_i = 11.187
                ha_i = 24.46

                ## @25
                kc25 = 27.238
                ko25 = 16.582
                i25 = 3.743*o2/21

                ### Calculated part in constant section
                ## Caluculations from Constants section @T leaf
                kc_pa_t = np.exp(c_kc - (ha_kc / (0.008314 * (273.15 + t_leaf))))
                ko_pa_t = np.exp(c_ko - (ha_ko / (0.008314 * (273.15 + t_leaf))))
                i_pa_t = np.exp(c_i - (ha_i / (0.008314 * (273.15 + t_leaf)))) * o2 / 21

                ### For adjusting to 25°C
                ## Constants
                c_vcmax = 26.355
                c_j = 17.71
                c_tpu = 21.46
                c_rd = 18.715
                c_gm =20.01

                ha_vcmax = 65.33
                ha_j = 43.9
                ha_tpu = 53.1
                ha_rd = 46.39
                ha_gm =49.6

                hd_tpu = 201.8
                s_tpu = 0.65

                hd_gm = 437.4
                s_gm = 1.4

                ## @25, calculated (constant section)
                vcmax25 = np.exp(c_vcmax - (ha_vcmax / (0.008314 * (273.15 + 25))))
                j25 = np.exp(c_j - (ha_j / (0.008314 * (273.15 + 25))))
                tpu25 = np.exp(c_tpu - ha_tpu / (0.008314 * (273.15 + 25)))/(1+np.exp((s_tpu * (25 + 273.15) - hd_tpu)/(0.008314*(25 + 273.15))))
                rd25 = np.exp(c_rd-(ha_rd / (0.008314 * (273.15 + 25))))
                gm25 = np.exp(c_gm-ha_gm / (0.008314 * (273.15 + 25))) / (1+np.exp((s_gm * (25 + 273.15) - hd_gm)/(0.008314 * (25 + 273.15))))

                ## @T leaf, calculated (constant section)
                vcmax_tleaf = np.exp(c_vcmax - (ha_vcmax/(0.008314*(273.15 + t_leaf))))
                j_tleaf = np.exp(c_j - (ha_j / (0.008314 * (273.15 + t_leaf))))
                tpu_tleaf = np.exp(c_tpu - ha_tpu / (0.008314 * (273.15 + t_leaf)))/(1+np.exp((s_tpu * (t_leaf + 273.15) - hd_tpu)/(0.008314*(t_leaf + 273.15))))
                rd_tleaf = np.exp(c_rd-(ha_rd / (0.008314 * (273.15 + t_leaf))))
                gm_tleaf = np.exp(c_gm-ha_gm / (0.008314 * (273.15 + t_leaf))) / (1+np.exp((s_gm * (t_leaf + 273.15) - hd_gm)/(0.008314 * (t_leaf + 273.15))))

                ### **From outputs section of the excel**
                ## Output constants
                ## @25
                vcmax_o25 = 100
                j_o25  = 140
                tpu_o25  = 9.7
                rd_o25  = 0.77
                gm_o25  = 8.64

                ## @ T leaf, reverse calculated
                vcmax = vcmax_tleaf * vcmax_o25 if not np.isnan(vcmax_tleaf * vcmax_o25 ) else 130
                j = j_tleaf * j_o25 if not np.isnan(j_tleaf * j_o25 ) else 167
                tpu = tpu_tleaf * tpu_o25 if not np.isnan(tpu_tleaf * tpu_o25 ) else 1.14
                rd = rd_tleaf * rd_o25 if not np.isnan(rd_tleaf * rd_o25 ) else 0.93
                gm = gm_tleaf * gm_o25 if not np.isnan(gm_tleaf * gm_o25 ) else 10.53

                ### Data for lines (fixed)
                cur_cc = [1, 2, 3, 4, 5, 7.5, 10, 12.5, 15, 17.5, 20, 22.5, 25, 27.5, 30, 
                          32.5, 35, 37.5, 40, 42.5, 45, 47.5, 50, 52.5, 55, 57.5, 60, 62.5, 
                          65, 67.5, 70, 72.5, 75, 77.5, 80, 82.5, 85, 87.5, 90, 92.5, 95]
                
                cur_cc_np = np.array(cur_cc)
                cur_ac = vcmax * (cur_cc_np - i_pa_t) / (cur_cc_np + kc_pa_t * (1 + o2 / ko_pa_t)) - rd
                cur_aj = j * (cur_cc_np - i_pa_t) / (4 * cur_cc_np + 8 * i_pa_t) - rd
                cur_at = 3 * tpu - rd
                cur_amin = np.minimum(np.minimum(cur_ac, cur_aj), cur_at)
                
                cur_df = pd.DataFrame({
                    'cc': cur_cc,
                    'ac': cur_ac,
                    'aj': cur_aj,
                    'at': cur_at,
                    'amin': cur_amin
                    })


                # print(cur_df)

                ### Prepare dataframe for analysing
                data = {'a': raw_df['r_a'], 'ci': raw_df['r_ci']}
                input_df = pd.DataFrame(data)
                # print(input_df)


                """  ---  Estimate limiting variations: From here --- """


                ### Standard variation
                input_df1 = input_df.copy()

                ### Create the new 'estimate_limit' column and calculate
                input_df1['est_lim'] = input_df1['ci'].apply(calculate_estimate_limit1)
                ### Calculate Ci pa, Cc, Ac, Aj, At
                input_df1['ci pa'] = input_df1['ci'] * patm * 0.001
                input_df1['cc'] = input_df1['ci pa'] - input_df1['a'] / gm
                input_df1['ac'] = vcmax * ((input_df1['cc'] - i_pa_t) / (input_df1['cc'] + kc_pa_t * (1 + o2 / ko_pa_t))) - rd
                input_df1['aj'] = j * ((input_df1['cc'] - i_pa_t) / (4 * input_df1['cc'] + 8 * i_pa_t)) - rd
                input_df1['at'] = 3 * tpu - rd
                input_df1['calc_a'] = input_df1[['ac', 'aj', 'at']].min(axis=1)

                ### Define column names for the error terms
                input_df1['err_ac'] = np.where(input_df1['est_lim'] == 1, (input_df1['ac'] - input_df1['a'])**2, np.nan)
                input_df1['err_aj'] = np.where(input_df1['est_lim'] == 2, (input_df1['aj'] - input_df1['a'])**2, np.nan)
                input_df1['err_at'] = np.where(input_df1['est_lim'] == 3, (input_df1['at'] - input_df1['a'])**2, np.nan)

                ### Variantion: 1
                input_df2 = input_df.copy()

                ### Create the new 'estimate_limit' column and calculate
                input_df2['est_lim'] = input_df2['ci'].apply(calculate_estimate_limit2)
                ### Calculate Ci pa, Cc, Ac, Aj, At
                input_df2['ci pa'] = input_df2['ci'] * patm * 0.001
                input_df2['cc'] = input_df2['ci pa'] - input_df2['a'] / gm
                input_df2['ac'] = vcmax * ((input_df2['cc'] - i_pa_t) / (input_df2['cc'] + kc_pa_t * (1 + o2 / ko_pa_t))) - rd
                input_df2['aj'] = j * ((input_df2['cc'] - i_pa_t) / (4 * input_df2['cc'] + 8 * i_pa_t)) - rd
                input_df2['at'] = 3 * tpu - rd
                input_df2['calc_a'] = input_df2[['ac', 'aj', 'at']].min(axis=1)

                ### Define column names for the error terms
                input_df2['err_ac'] = np.where(input_df2['est_lim'] == 1, (input_df2['ac'] - input_df2['a'])**2, np.nan)
                input_df2['err_aj'] = np.where(input_df2['est_lim'] == 2, (input_df2['aj'] - input_df2['a'])**2, np.nan)
                input_df2['err_at'] = np.where(input_df2['est_lim'] == 3, (input_df2['at'] - input_df2['a'])**2, np.nan)


                ### Plot / curve fitting
                ### Variation standard
                plot_df1 = input_df1[input_df1['est_lim'] != 0].copy()
                fig1, ax1 = plt.subplots(figsize=(8, 5))
                fig1.patch.set_facecolor("#ded2ce")  # Outer area
                ax1.set_facecolor("#f3eeee")  # Plot area

                ax1.plot(plot_df1['cc'], plot_df1['a'], 'o-', label='Aobs (Measured)', color='royalblue', markersize=8, markerfacecolor='white', markeredgewidth=2)
                
                ax1.plot(cur_df['cc'], cur_df['ac'], '--', label='Ac (Rubisco-limited)', color='crimson', linewidth=2) # making smoother
                ax1.plot(cur_df['cc'], cur_df['aj'], '--', label='Aj (RuBP-limited)', color='seagreen', linewidth=2)
                ax1.plot(cur_df['cc'], cur_df['at'], '-.', label='At (TPU-limited)',  color='darkorange', linewidth=2)
                ax1.plot(cur_df['cc'], cur_df['amin'], '-', label='min(Ac, Aj, At)', color='black',  linewidth=1)
                ax1.set_xlabel('Cc (ppm)')
                ax1.set_ylabel(r'A ($\mu$mol m$^{-2}$ s$^{-1}$)')
                ax1.set_title(f'A/Cc Curve: {file}, , Standard')
                ax1.grid(True)
                ax1.legend()

                img_buf1 = BytesIO()
                plt.savefig(img_buf1, format='png')
                img_buf1.seek(0)
                plt.close(fig1)

                # print(input_df1)

                ### Variation 1
                plot_df2 = input_df2[input_df2['est_lim'] != 0].copy()
                fig2, ax2 = plt.subplots(figsize=(8, 5))
                fig2.patch.set_facecolor("#f1cec2")  # Outer area
                ax2.set_facecolor("#f3eeee")  # Plot area

                ax2.plot(plot_df2['cc'], plot_df2['a'], 'o-', label='Aobs (Measured)', color='royalblue', markersize=8, markerfacecolor='white', markeredgewidth=2)
                

                ax2.plot(cur_df['cc'], cur_df['ac'], '--', label='Ac (Rubisco-limited)', color='crimson', linewidth=2) # making smoother
                ax2.plot(cur_df['cc'], cur_df['aj'], '--', label='Aj (RuBP-limited)', color='seagreen', linewidth=2)
                ax2.plot(cur_df['cc'], cur_df['at'], '-.', label='At (TPU-limited)',  color='darkorange', linewidth=2)
                ax2.plot(cur_df['cc'], cur_df['amin'], '-', label='min(Ac, Aj, At)', color='black',  linewidth=1)
                ax2.set_xlabel('Cc (ppm)')
                ax2.set_ylabel(r'A ($\mu$mol m$^{-2}$ s$^{-1}$)')
                ax2.set_title(f'A/Cc Curve: {file}, Variation 1')
                ax2.grid(True)
                ax2.legend()

                img_buf2 = BytesIO()
                plt.savefig(img_buf2, format='png')
                img_buf2.seek(0)
                plt.close(fig2)

                """ ^^^  ---  Estimate limiting variations: Ends here --- ^^^ """

                ### Plot A/Ci graph at the end, based on standard estimate limiting
                plot_dfci = input_df1[input_df1['est_lim'] != 0].copy()
                figci, axci = plt.subplots(figsize=(8, 5))
                axci.plot(plot_dfci['ci'], plot_dfci['a'], 'o-', label='Aobs (Measured)', color='royalblue', markersize=8, markerfacecolor='white', markeredgewidth=2)

                axci.plot(plot_dfci['ci'], plot_dfci['ac'], '--', label='Ac (Rubisco-limited)',  color='crimson', linewidth=2)
                axci.plot(plot_dfci['ci'], plot_dfci['aj'], '--', label='Aj (RuBP-limited)', color='seagreen', linewidth=2)
                axci.plot(plot_dfci['ci'], plot_dfci['at'], '-.', label='At (TPU-limited)', color='darkorange', linewidth=2)
                axci.plot(plot_dfci['ci'], plot_dfci['calc_a'], '-', label='min(Ac, Aj, At)', color='black',  linewidth=1)
                axci.set_xlabel('Ci (ppm)')
                axci.set_ylabel(r'A ($\mu$mol m$^{-2}$ s$^{-1}$)')
                axci.set_title(f'A/Ci Curve: {file}, , Standard observation (without smoothing)')
                axci.grid(True)
                axci.legend()

                img_bufci = BytesIO()
                plt.savefig(img_bufci, format='png')
                img_bufci.seek(0)
                plt.close(figci)

                """ ^^^  ---  Separate curve for A/Ci --- ^^^ """

                ### Build table for output section
                ## Create table structure
                otpt_data = {
                    'Outputs': ['Vcmax', 'J', 'TPU', 'Rd*', 'gm*'],
                    '@ T leaf': [vcmax, j, tpu, rd, gm],
                    '@ 25 °C': [vcmax_o25, j_o25, tpu_o25, rd_o25, gm_o25],
                    'Units': ['μmol m⁻² s⁻¹', 'μmol m⁻² s⁻¹', 'μmol m⁻² s⁻¹', 'μmol m⁻² s⁻¹', 'μmol m⁻² s⁻¹ Pa⁻¹']
                }

                # Create DataFrame for output
                otpt_df_table = pd.DataFrame(otpt_data)

                ### Build table for Constants section
                ## Create table structure
                cons1_data = {
                    'For Fitting': ['Kc (Pa)', 'Ko (kPa)', 'I* (Pa)'],
                    '@ T leaf': [kc_pa_t, ko_pa_t, i_pa_t],
                    '@ 25 °C': [kc25, ko25, i25],
                    'c': [c_kc, c_ko, c_i],
                    'Ha': [ha_kc, ha_ko, ha_i],
                    'Hd': ["", "", ""],
                    'S': ["", "", ""],
                }
                con1_df_table = pd.DataFrame(cons1_data)

                cons2_data = {
                    'For adjusting to 25°C': ['Vcmax', 'J', 'TPU', 'Rd', 'gm'],
                    '@ T leaf': [vcmax_tleaf, j_tleaf, tpu_tleaf, rd_tleaf, gm_tleaf],
                    '@ 25 °C': [vcmax25, j25, tpu25, rd25, gm25],
                    'c': [c_vcmax, c_j, c_tpu, c_rd, c_gm],
                    'Ha': [ha_vcmax, ha_j, ha_tpu, ha_rd, ha_gm],
                    'Hd': ["", "", hd_tpu, "", hd_gm],
                    'S': ["", "", s_tpu, "", s_gm],
                }

                con2_df_table = pd.DataFrame(cons2_data)

                con_hed = {
                    'Constants': [''],
                }
                ch_df = pd.DataFrame(con_hed)

                ### Build table for O2, Patm and tleaf mean values
                mn_hd = {
                    'Means': ['Tleaf', 'Patm', 'O2'],
                    'Values': [t_leaf, patm, o2]
                }
                mean_tab = pd.DataFrame(mn_hd)


                ### Write DataFrames
                sheet_name = f"{raw_id}"[:31]
                input_df1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False) # Standard
                input_df2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False) # Variation 1
                otpt_df_table.to_excel(writer, sheet_name=sheet_name, startrow=36, startcol=0, index=False) # Output table
                ch_df.to_excel(writer, sheet_name=sheet_name, startrow=43, startcol=0, index=False) # Constants header
                con1_df_table.to_excel(writer, sheet_name=sheet_name, startrow=44, startcol=0, index=False) # Constants 1 table
                con2_df_table.to_excel(writer, sheet_name=sheet_name, startrow=48, startcol=0, index=False) # Constants 2 table
                mean_tab.to_excel(writer, sheet_name=sheet_name, startrow=36, startcol=7, index=False) # Mean values

                ### Insert plots
                worksheet = writer.sheets[sheet_name]
                worksheet.insert_image('A12', f"{file}.png", {'image_data': img_buf1}) # Standard
                worksheet.insert_image('P12', f"{file}.png", {'image_data': img_buf2}) # Variation 1

                worksheet.insert_image('A55', f"{file}.png", {'image_data': img_bufci}) # A/Ci curve

                print(f"✅ Processed and written: {file}")

            except Exception as e:
                print(f"❌ Error processing {file}: {e}")

"""Append df_r_temp to CSV file"""
if r_data:
    final_r_df = pd.concat(r_data, ignore_index=True)
    with open(csv_output_file, "w", newline="") as f:
        f.write("from py\n")  # A1 cell text
    final_r_df.to_csv(csv_output_file, mode="a", index=False)
    print("✅ .csv file has been added successfully!")
else:
    print("⚠️ No valid data found.")

"""END: Append df_r_temp to CSV file"""


""" Combined output sheet """

### Prepare list to collect summary rows
output_list = []

### Loop through each file in the output folder
for file in os.listdir(output_folder):
    if file.endswith(".xlsx"):
        file_path = os.path.abspath(os.path.join(output_folder, file))
        # print(f"Processing: {file}")

        ### Open the final workbook
        output_xls = pd.ExcelFile(file_path)

        ### Loop through all sheets
        for sheet_name in output_xls.sheet_names:
            try:
                ### Read only columns A and B (first two), no header so we can access rows directly
                op_ls_df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="A:B", header=None)

                ### Find the row index where "Vcmax" appears
                vcmax_row = op_ls_df[op_ls_df[0] == "Vcmax"].index

                if len(vcmax_row) > 0:
                    row = vcmax_row[0]
                    vcmax = op_ls_df.loc[row, 1]
                    J = op_ls_df.loc[row + 1, 1]
                    TPU = op_ls_df.loc[row + 2, 1]
                    Rd = op_ls_df.loc[row + 3, 1]
                    gm = op_ls_df.loc[row + 4, 1]

                    ### Append as dict
                    output_list.append({
                        "plot": sheet_name,
                        "Vcmax": vcmax,
                        "J": J,
                        "TPU": TPU,
                        "Rd*": Rd,
                        "gm*": gm
                    })

            except Exception as e:
                print(f"Skipped sheet {sheet_name}: {e}")

        ### Create summary DataFrame
        output_df = pd.DataFrame(output_list)

        ### Save to a new sheet in the same file
        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name="Compiled_Outputs", index=False)

        print("✅ Compiled outputs sheet added successfully!")
