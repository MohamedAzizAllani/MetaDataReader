from pySPM import Bruker
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox as mb
import os
import sys
from os.path import dirname, join,abspath
import logging
import subprocess

# Setup logging
log_dir = os.path.join(os.getenv('LOCALAPPDATA'), 'Meta_Data_Reader')
os.makedirs(log_dir, exist_ok=True)  # Create directory if it doesn't exist
log_file = os.path.join(log_dir, 'script.log')
logging.basicConfig(level=logging.DEBUG, filename=log_file, filemode='a')
logging.debug("Script started")

def create_gui(root, columns):
    logging.debug("Entering create_gui")
    interesting_columns = [
        'Filename', 'Up / Down', 'Setpoint', 'Sample Bias', 'Lines (real)', 
        'Lines (set.)', 'Samples per line', 'Scan x size', 'Scan y size (setting)', 
        'X Offset', 'Y Offset', 'Stagepos X', 'Stagepos Y', 'Capture Type '
    ]
    interleave_cols = ['IL Setpoint', 'IL I Gain', 'IL P Gain', 'IL Sample Bias']
    
    root.title("Select Metadata Columns")
    root.configure(bg='#ffffff')
    
    try:
        # Determine logo path
        if getattr(sys, 'frozen', False):  # PyInstaller executable
            base_path = sys._MEIPASS
        else:  # Running as script
            base_path = os.path.dirname(__file__)
        logo_path = os.path.join(base_path, 'logo.ico')
        root.iconbitmap(logo_path)  # Use .ico with iconbitmap
        logging.debug("Custom logo set successfully")
    except Exception as e:
        logging.debug(f"Error setting logo: {str(e)}")
        print(f"Error setting logo: {str(e)}")
    
    style = ttk.Style()
    if "modern" not in style.theme_names():
        style.theme_create("modern", parent="clam", settings={
            "TCheckbutton": {
                "configure": {
                    "font": ('Helvetica', 10),
                    "padding": 6,
                    "background": '#ffffff',
                    "foreground": '#000000',
                    "indicatorbackground": '#ffffff',
                    "indicatorforeground": '#007bff',
                    "indicatormargin": [2, 2, 2, 2]
                },
                "map": {
                    "background": [("selected", '#ffffff'), ("!selected", '#ffffff')],
                    "foreground": [("selected", '#000000'), ("!selected", '#000000')],
                    "indicatorbackground": [("selected", '#007bff'), ("!selected", '#ffffff')],
                    "indicatorforeground": [("selected", '#ffffff'), ("!selected", '#000000')]
                }
            },
            "Disabled.TCheckbutton": {
                "configure": {
                    "font": ('Helvetica', 10),
                    "padding": 6,
                    "background": '#ffffff',
                    "foreground": '#808080',
                    "indicatorbackground": '#e0e0e0',
                    "indicatorforeground": '#808080'
                },
                "map": {
                    "background": [("disabled", '#ffffff')],
                    "foreground": [("disabled", '#808080')],
                    "indicatorbackground": [("disabled", '#e0e0e0')],
                    "indicatorforeground": [("disabled", '#808080')]
                }
            },
            "TButton": {
                "configure": {
                    "font": ('Helvetica', 10, 'bold'),
                    "padding": 10,
                    "background": '#00b7eb',
                    "foreground": '#000000',
                    "borderwidth": 0,
                    "focusthickness": 0,
                    "relief": 'flat'
                },
                "map": {
                    "background": [("active", '#00a3d6'), ("!active", '#00b7eb')],
                    "foreground": [("active", '#000000'), ("!active", '#000000')]
                }
            },
            "TFrame": {
                "configure": {
                    "background": '#ffffff'
                }
            }
        })
    style.theme_use("modern")
    style.configure("TButton", bordercolor='#00b7eb', lightcolor='#00e6ff', darkcolor='#008bb5', anchor='center')
    
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill='both', expand=True)
    
    vars_dict = {col: tk.BooleanVar(value=col in interesting_columns) for col in columns}
    
    def update_interleave_cols(*args):
        interleave_state = vars_dict['Interleavemode'].get()
        for col in interleave_cols:
            vars_dict[col].set(interleave_state)
            # Update style based on Interleavemode state
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Checkbutton) and widget.cget("text") == col:
                    widget.configure(style='Disabled.TCheckbutton' if not interleave_state else 'TCheckbutton')
    
    select_all_var = tk.BooleanVar(value=False)
    def toggle_select_all():
        if select_all_var.get():
            for col in columns:
                if col == 'Interleavemode' or col in interleave_cols:
                    continue
                if not vars_dict[col].get():
                    vars_dict[col].set(True)
        else:
            for col in columns:
                if col not in interesting_columns and col != 'Interleavemode' and col not in interleave_cols:
                    vars_dict[col].set(False)
    
    ttk.Checkbutton(frame, text="Select All", variable=select_all_var, command=toggle_select_all).grid(row=0, column=0, columnspan=8, sticky=tk.W, pady=5)
    
    for i, col in enumerate(interesting_columns):
        row = (i % 5) + 1
        col_idx = i // 5
        cb = ttk.Checkbutton(frame, text=col, variable=vars_dict[col])
        cb.grid(row=row, column=col_idx, sticky=tk.W, padx=10, pady=3)
    
    interleave_all = ['Interleavemode'] + interleave_cols
    interleave_start_col = (len(interesting_columns) - 1) // 5 + 1
    for i, col in enumerate(interleave_all):
        row = i + 1
        if col == 'Interleavemode':
            cb = ttk.Checkbutton(frame, text=col, variable=vars_dict[col], command=update_interleave_cols)
        else:
            cb = ttk.Checkbutton(frame, text=col, variable=vars_dict[col], style='Disabled.TCheckbutton' if not vars_dict['Interleavemode'].get() else 'TCheckbutton')
        cb.grid(row=row, column=interleave_start_col, sticky=tk.W, padx=10, pady=3)
    
    remaining_columns = [col for col in columns if col not in interesting_columns and col != 'Interleavemode' and col not in interleave_cols]
    remaining_start_col = interleave_start_col + 1
    for i, col in enumerate(remaining_columns):
        row = (i % 5) + 1
        col_idx = (i // 5) + remaining_start_col
        cb = ttk.Checkbutton(frame, text=col, variable=vars_dict[col])
        cb.grid(row=row, column=col_idx, sticky=tk.W, padx=10, pady=3)
    
    selected_columns = []
    submit_var = tk.BooleanVar()
    def submit():
        nonlocal selected_columns
        selected_columns[:] = [col for col, var in vars_dict.items() if var.get()]
        logging.debug(f"Columns selected in submit: {selected_columns}")
        frame.destroy()
        submit_var.set(True)
    
    def on_closing():
        nonlocal selected_columns
        selected_columns[:] = []
        logging.debug("GUI closed without submit")
        frame.destroy()
        submit_var.set(True)
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    ttk.Button(frame, text="Submit", command=submit).grid(row=6, column=0, columnspan=8, pady=15)
    logging.debug("Waiting for GUI interaction")
    root.wait_variable(submit_var)
    logging.debug(f"Returning selected columns: {selected_columns}")
    return selected_columns

def retrieve_data(ScanB, Columns, spm_scanner_keys, spm_layers_keys, max_channels):
    Meta_Data = []
    missing_columns = set()
    channel1_missing_columns = set()
    num_channels = len(ScanB.layers)
    logging.debug(f"Processing file with {num_channels} channels")
    
    for j in range(0, max_channels):
        Parameters = {}
        if j >= num_channels:
            if 'Filename' in Columns:
                Parameters['Filename'] = ''
            for col in Columns:
                if col != 'Filename' and col != 'Channel No.':
                    Parameters[col] = 'missing'
                    missing_columns.add(col)
            Meta_Data.append(Parameters)
            continue
        
        for col in Columns:
            if col == 'Filename':
                continue
            idx = Columns.index(col)
            try:
                if idx < len(spm_scanner_keys):
                    if col == 'Channel No.':
                        Parameters[col] = str(j + 1)
                    else:
                        value = ScanB.scanners[0][spm_scanner_keys[idx]][0]
                        if value is None or value == b'':
                            Parameters[col] = 'missing'
                            missing_columns.add(col)
                            if j == 0:  # Only track for channel 1
                                channel1_missing_columns.add(col)
                            logging.debug(f"Missing data for column {col} in scanner")
                            continue
                        Parameters[col] = value.decode("utf-8")
                else:
                    if spm_layers_keys[idx - len(spm_scanner_keys)] == b'@2:Image Data':
                        try:
                            value = ScanB.layers[j][spm_layers_keys[idx - len(spm_scanner_keys)]][0]
                            if value is None or value == b'':
                                Parameters[col] = 'missing'
                                missing_columns.add(col)
                                if j == 0:
                                    channel1_missing_columns.add(col)
                                logging.debug(f"Missing data for column {col} in layers (@2:Image Data)")
                                continue
                            Parameters[col] = value.decode("utf-8") + ' (normal)'
                        except KeyError:
                            value = ScanB.layers[j][b'@3:Image Data'][0]
                            if value is None or value == b'':
                                Parameters[col] = 'missing'
                                missing_columns.add(col)
                                if j == 0:
                                    channel1_missing_columns.add(col)
                                logging.debug(f"Missing data for column {col} in layers (@3:Image Data)")
                                continue
                            Parameters[col] = value.decode("utf-8") + ' (interleave)'
                    else:
                        value = ScanB.layers[j][spm_layers_keys[idx - len(spm_scanner_keys)]][0]
                        if value is None or value == b'':
                            Parameters[col] = 'missing'
                            missing_columns.add(col)
                            if j == 0:
                                channel1_missing_columns.add(col)
                            logging.debug(f"Missing data for column {col} in layers")
                            continue
                        Parameters[col] = value.decode("utf-8")
            except (KeyError, IndexError, AttributeError) as e:
                logging.debug(f"Missing data for column {col}: {str(e)}")
                Parameters[col] = 'missing'
                missing_columns.add(col)
                if j == 0:
                    channel1_missing_columns.add(col)
                continue
        Meta_Data.append(Parameters)
    return Meta_Data, missing_columns, channel1_missing_columns

def retreive_num_val(value):
    for i in range(1, len(str(value)) + 1):
        if not (str(value)[-i] == '.'):
            try:
                int(str(value)[-i])
            except:
                if str(value)[-i] == ':':
                    i -= 1
                break
            else:
                pass
    return str(float(str(value)[-i:]))

def get_real_aspect_ratio(spl, lines):
    AR = spl / lines
    if AR == 1:
        return 1
    else:
        return AR

def correct_units_and_values(dictt, Columns, NVP, Unit_dict):
    for key, value in dictt.items():
        if key not in NVP:
            if value == 'missing':
                dictt[key] = 'missing'
                continue
            try:
                int(str(value)[-1])
            except:
                for i in range(1, len(str(value)) + 1):
                    x = str(value)[-i]
                    try:
                        int(x)
                    except:
                        pass
                    else:
                        dictt[key] = retreive_num_val(str(value)[:-i + 1])
                        break
            else:
                dictt[key] = retreive_num_val(value)
            dictt[key] = dictt[key] + ' ' + Unit_dict.get(key, '')
        else:
            if key == 'Interleavemode':
                if len(value) > 20:
                    dictt[key] = dictt[key][22:-1]
            if key == 'Channel Name':
                if not (dictt[key].find(']') == -1):
                    dictt[key] = dictt[key][dictt[key].find(']') + 2:]
    if 'Samples per line' in dictt and 'Lines (set.)' in dictt:
        if dictt['Samples per line'] != 'missing' and dictt['Lines (set.)'] != 'missing':
            dictt['Asp. rat. (set.)'] = "{:.2f}".format(float(dictt['Samples per line']) / float(dictt['Lines (set.)']))
        else:
            dictt['Asp. rat. (set.)'] = 'missing'
    if 'Samples per line' in dictt and 'Lines (real)' in dictt:
        if dictt['Samples per line'] != 'missing' and dictt['Lines (real)'] != 'missing':
            dictt['Asp. rat. (real)'] = "{:.2f}".format(float(dictt['Samples per line']) / float(dictt['Lines (real)']))
        else:
            dictt['Asp. rat. (real)'] = 'missing'
    if 'Channel No.' in dictt:
        dictt['Channel No.'] = str(int(dictt['Channel No.'][0]))
    if 'Lines (real)' in dictt and dictt['Lines (real)'] != 'missing':
        dictt['Lines (real)'] = str(int(float(dictt['Lines (real)'])))
    if 'Lines (set.)' in dictt and dictt['Lines (set.)'] != 'missing':
        dictt['Lines (set.)'] = str(int(float(dictt['Lines (set.)'])))
    if 'Samples per line' in dictt and dictt['Samples per line'] != 'missing':
        dictt['Samples per line'] = str(int(float(dictt['Samples per line'])))
    return dictt



# [create_gui, retrieve_data, retreive_num_val, correct_units_and_values unchanged]

def main():
    from datetime import datetime
    import re
    import os
    logging.debug("Entering main")
    Columns = [
        'Channel No.', 'Probe', 'Scan x size', 'Scan y size (setting)', 'X Offset', 'Y Offset',
        'Scan Angle', 'Stagepos X', 'Stagepos Y', 'Samples per line', 'Lines (set.)',
        'Asp. rat. (set.)', 'Scan Rate', 'Tip Velocity', 'Units', 'Setpoint Units',
        'Up / Down', 'Capture Type ', 'Interleavemode', 'Deflection Sens.', 'Force Sens.',
        'Setpoint', 'IL Setpoint', 'I Gain', 'P Gain', 'IL I Gain', 'IL P Gain',
        'Sample Bias', 'IL Sample Bias', 'Lines (real)', 'Asp. rat. (real)', 'Line Direction',
        'Channel Name'
    ]
    NVP = ['Probe', 'Units', 'Setpoint Units', 'Up / Down', 'Capture Type ', 'Interleavemode', 'Line Direction', 'Channel Name']
    
    Units = [
        '', '', 'nm', 'nm', 'nm', 'nm', '°', 'µm', 'µm', '', '', '', 'Hz', 'µm/s', '', '',
        '', '', '', 'nm/V', 'nN/V', 'V', 'V', '', '', '', '', 'V', 'V', '', '', '', ''
    ]
    Unit_dict = dict(zip(Columns, Units))
    
    spm_scanner_keys = [
        b'0', b'Tip Serial Number', b'Scan Size', b'Slow Axis Size', b'X Offset', b'Y Offset',
        b'Rotate Ang.', b'Stage X', b'Stage Y', b'Samps/line', b'Lines', b'Aspect Ratio',
        b'Scan Rate', b'Tip Velocity', b'Units', b'Setpoint Units', b'Capture direction',
        b'Capture mode ', b'@InterleaveList', b'@Sens. DeflSens', b'@Sens. ForceDeflSens',
        b'@2:AFMSetDeflection', b'@3:AFMSetDeflection', b'@2:AFMFbIgain', b'@2:AFMFbPgain',
        b'@3:AFMFbIgain', b'@3:AFMFbPgain', b'@2:SSRMSampleBias', b'@3:SSRMSampleBias'
    ]
    spm_layers_keys = [
        b'Number of lines', b'Aspect Ratio', b'Line Direction', b'@2:Image Data'
    ]
    
    root = tk.Tk()
    while True:
        try:
            selected_columns = create_gui(root, [col for col in Columns if col != 'Channel No.'] + ['Filename'])
        except Exception as e:
            logging.debug(f"Error in create_gui: {str(e)}")
            print(f"Error in create_gui: {str(e)}")
            root.destroy()
            return
        
        logging.debug(f"Selected columns in main: {selected_columns}")
        print(f"Selected columns: {selected_columns}")
        if not selected_columns:
            logging.debug("No columns selected, exiting.")
            print("No columns selected, exiting.")
            root.destroy()
            return
        
        preferred_order = ['Filename', 'Channel Name', 'Probe', 'Up / Down', 'Line Direction', 'Setpoint', 'Sample Bias', 'Scan Rate', 'Tip Velocity', 'Scan Angle', 'Lines (real)', 'Samples per line', 'Asp. rat. (real)', 'Scan x size', 'Scan y size (setting)', 'Interleavemode',
                          'IL Sample Bias', 'IL Setpoint', 'IL I Gain', 'IL P Gain', 'I Gain', 'P Gain', 'X Offset', 'Y Offset', 'Capture Type ', 'Lines (set.)', 'Asp. rat. (set.)', 'Deflection Sens.', 'Force Sens.', 'Stagepos X', 'Stagepos Y', 'Units',
                          'Setpoint Units']
        
        filtered_columns = [col for col in Columns if col in selected_columns]
        if 'Filename' in selected_columns:
            filtered_columns.append('Filename')
        filtered_scanner_keys = [spm_scanner_keys[Columns.index(col)] for col in filtered_columns if col in Columns and Columns.index(col) < len(spm_scanner_keys)]
        filtered_layers_keys = [spm_layers_keys[Columns.index(col) - len(spm_scanner_keys)] for col in filtered_columns if col in Columns and Columns.index(col) >= len(spm_scanner_keys)]
        filtered_NVP = [col for col in NVP if col in filtered_columns]
        filtered_unit_dict = {col: Unit_dict.get(col, '') for col in filtered_columns}
        
        filtered_order = [col for col in preferred_order if col in selected_columns]
        filtered_order.extend([col for col in selected_columns if col not in preferred_order])
        
        logging.debug("Opening file dialog")
        try:
            filez = fd.askopenfilenames(
                parent=root,
                title="Choose SPM or Numeric Files"
            )
            logging.debug(f"Selected files: {filez}")
            print(f"Selected files: {filez}")
            if not filez:
                logging.debug("No files selected, reopening GUI.")
                print("No files selected, reopening GUI.")
                mb.showwarning("No Files Selected", "Please select valid .spm files.")
                continue
        except Exception as e:
            logging.debug(f"Error in file dialog: {str(e)}")
            print(f"Error in file dialog: {str(e)}")
            logging.debug("Falling back to console file input")
            print("File dialog failed. Enter .spm file paths (one per line, press Enter twice to finish):")
            filez = []
            while True:
                path = input()
                if path == "":
                    break
                filez.append(path)
            logging.debug(f"Console selected files: {filez}")
            print(f"Console selected files: {filez}")
            if not filez:
                logging.debug("No files selected, reopening GUI.")
                print("No files selected, reopening GUI.")
                mb.showwarning("No Files Selected", "Please select valid .spm files.")
                continue
        
        # Validate extensions
        invalid_extensions = []
        valid_files = []
        for d in filez:
            ext = os.path.splitext(d)[1].lower()
            if ext == '.spm' or re.match(r'^\.\d{3}$', ext):
                valid_files.append(d)
                logging.debug(f"Valid extension detected: {d}")
            else:
                invalid_extensions.append(d)
                logging.debug(f"Invalid extension detected: {d}")
        
        # Show warning for invalid files and exit if no valid files
        if invalid_extensions:
            invalid_msg = "The following files have unsupported extensions and will be skipped:\n" + "\n".join(invalid_extensions) + "\n\nProcessing valid .spm files."
            mb.showwarning("Invalid Files Skipped", invalid_msg)
            logging.debug(invalid_msg)
            print(invalid_msg)
            if not valid_files:
                continue
        
        max_channels = 0
        channel_counts = {}
        raw_data = []
        all_channel1_missing_columns = set()
        invalid_files = invalid_extensions
        
        # Process valid files
        for d in valid_files:
            try:
                logging.debug(f"Checking channels for file: {d}")
                with open(d, 'r', encoding='utf-8', errors='ignore') as f:
                    first_line = f.readline().strip()
                    if 'ISO/TC 201 SPM data transfer format' in first_line:
                        logging.debug(f"Invalid file format (ISO/TC 201): {d}")
                        invalid_files.append(d)
                        continue
                ScanB = Bruker(d)
                num_channels = len(ScanB.layers)
                channel_counts[d] = num_channels
                max_channels = max(max_channels, num_channels)
                logging.debug(f"Valid file with {num_channels} channels: {d}")
            except Exception as e:
                logging.debug(f"Error checking channels for {d}: {str(e)}")
                print(f"Error checking channels for {d}: {str(e)}")
                invalid_files.append(d)
                continue
        
        if len(set(channel_counts.values())) > 1:
            message = f"Not all files have the same number of data channels. Output includes all channels up to the maximum ({max_channels})."
            mb.showwarning("Channel Count Mismatch", message)
            logging.debug(f"Channel count mismatch: {channel_counts}")
        
        for d in valid_files:
            if d in invalid_files:
                continue
            try:
                logging.debug(f"Processing file: {d}")
                ScanB = Bruker(d)
                Parameters, missing_cols, channel1_missing_cols = retrieve_data(ScanB, filtered_columns, filtered_scanner_keys, filtered_layers_keys, max_channels)
                all_channel1_missing_columns.update(channel1_missing_cols)
                for dictt in Parameters:
                    dictt = correct_units_and_values(dictt, filtered_columns, filtered_NVP, filtered_unit_dict)
                raw_data.append(Parameters)
                logging.debug(f"Successfully processed file: {d}")
            except Exception as e:
                logging.debug(f"Error processing {d}: {str(e)}")
                print(f"Error processing {d}: {str(e)}")
                invalid_files.append(d)
                continue
        
        if invalid_files and not raw_data:
            invalid_msg = "The following files were not processed due to incompatible format or errors:\n" + "\n".join(invalid_files) + "\n\nPlease choose valid .spm files."
            mb.showwarning("No Valid Files Processed", invalid_msg)
            continue
        
        # Show missing columns warning for channel 1 only
        if all_channel1_missing_columns:
            missing_msg = "\n".join(f"Missing data for column {col}" for col in sorted(all_channel1_missing_columns))
            logging.debug(f"Missing columns: {all_channel1_missing_columns}")
        
        outer_list = []
        for channel in range(0, max_channels):
            inner_list = []
            for file in range(0, len(raw_data)):
                inner_list.append(raw_data[file][channel])
                if 'Filename' in filtered_columns:
                    inner_list[-1].update({'Filename': valid_files[file]})
            outer_list.append(inner_list)
        
        target_path = join(dirname(abspath(valid_files[0])), f'Meta_Data_{datetime.now().strftime("%Y.%m.%d_%H.%M.%S")}.xlsx')
        with pd.ExcelWriter(target_path) as writer:
            for i in range(0, len(outer_list)):
                df = pd.DataFrame(outer_list[i])
                df = df.reindex(columns=filtered_order)
                df.to_excel(writer, sheet_name='Channel' + str(i + 1), index=True)
                workbook = writer.book
                worksheet = writer.sheets['Channel' + str(i + 1)]
                format1 = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
                format1.set_font_size(8)
                format2 = workbook.add_format({'num_format': '@', 'align': 'right'})
                format2.set_font_size(14)
                format_index = workbook.add_format({'align': 'center'})
                format_index.set_font_size(8)
                worksheet.set_column(0, 0, 5, format_index)
                for col in range(1, len(filtered_order) + 1):  # Fixed: len(filtered_order) + 1
                    if filtered_order[col - 1] == 'Filename':
                        worksheet.set_column(col, col, 30, format2)
                    else:
                        worksheet.set_column(col, col, 20, format1)
        # Open file browser at the directory containing the Excel file
        directory = dirname(abspath(target_path))
        if os.path.exists(directory):
            logging.debug(f"Opening file browser at: {directory}")
            subprocess.run(["explorer", directory], shell=True)
        else:
            logging.debug(f"Directory not found: {directory}, opening Documents instead")
            subprocess.run(["explorer", os.path.expanduser("~/Documents")], shell=True)
        success_msg = f"Excel file created successfully at {target_path}\n\n{len(raw_data)} files processed successfully."
        mb.showinfo("Success", success_msg)
        logging.debug(success_msg)
        print(success_msg)
    
if __name__ == '__main__':
    main()