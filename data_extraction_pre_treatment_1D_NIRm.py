import os
import glob
import datetime
from statistics import mean, stdev
from tkinter import filedialog
import matplotlib.pyplot as plt
import itertools
import xlrd  # Reverted back to xlrd for legacy .xls support

# The Required Naming Format
# [X-Value]_[GroupLabel]_[AnythingElse].xls

# --- Configuration Constants ---
SHEET_NAME = 'Sheet1'
OUTPUT_FILE_NAME = 'output.txt'
DATA_START_ROW = 10  # 0-indexed for xlrd (matches your old index 10)
TARGET_COLUMN = 2    # Column C (0-indexed for xlrd: A=0, B=1, C=2)

def get_file_data(folder_path, filename):
    """
    Opens an old .xls Excel file, extracts the timestamp from B1, 
    and reads data points dynamically from a specific column until the end of the sheet.
    """
    file_path = os.path.join(folder_path, filename)
    
    # Open workbook using xlrd for legacy .xls files
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheet_by_name(SHEET_NAME)
    
    # B1 corresponds to row 0, column 1 in 0-indexed xlrd
    file_date = ws.cell_value(0, 1)
    
    # Read the data column range dynamically until the last populated row
    data_points = []
    total_rows = ws.nrows  # Dynamically gets the total number of rows in the sheet
    
    for row in range(DATA_START_ROW, total_rows):
        val = ws.cell_value(row, TARGET_COLUMN)
        if val != '' and val is not None:
            try:
                data_points.append(float(val))
            except ValueError:
                # Skips any accidental string artifacts or text at the bottom of the column
                continue
            
    return file_date, data_points

def calculate_delta_times(date_strings):
    """Converts a list of 'YYYY-MM-DD HH:MM:SS' strings into elapsed seconds from the first entry."""
    time_secs = []
    for date_str in date_strings:
        # Expecting string structure: "YYYY-MM-DD HH:MM:SS"
        _, time_part = date_str.split(' ')
        h, m, s = map(int, time_part.split(':'))
        
        total_seconds = int(datetime.timedelta(hours=h, minutes=m, seconds=s).total_seconds())
        time_secs.append(total_seconds)
        
    initial_time = time_secs[0]
    return [t - initial_time for t in time_secs]

def save_final_data(save_folder_path, source_folder_name, data_records):
    """Writes the compiled statistical analysis cleanly into a tab-separated text file in the chosen directory."""
    file_name = f"{source_folder_name}_{OUTPUT_FILE_NAME}"
    full_output_path = os.path.join(save_folder_path, file_name)
    
    header = 'date\tdelta_time\tfilename\tfilename_cut\taverage\tdeviation\n'
    units = ' \ts\t \t \tmV\tmV\n'
    
    with open(full_output_path, "w") as file:
        file.write(header)
        file.write(units)
        for r in data_records:
            file.write(f"{r['date']}\t{r['delta_time']}\t{r['filename']}\t"
                       f"{r['filename_cut']}\t{r['avg']:.4f}\t{r['std']:.4f}\n")
            
    print(f"\nSuccess! Output file saved to:\n--> {full_output_path}")

def choose_x_axis(filename_cuts):
    """Prompts the user to decide how they want to scale their X axis values."""
    type_plot = input('\nWanna plot x as it is? Say y/n: ').strip().lower()
    
    if type_plot == 'y':
        return filename_cuts, "x 'as it is'"
        
    type_plot2 = input("Wanna plot '1-10^-x'? Say y/n: ").strip().lower()
    if type_plot2 == 'y':
        x_transformed = [1 - (10 ** -x) for x in filename_cuts]
        print('Plotting x as an absorption factor.')
        return x_transformed, 'x as abs factor'
        
    print('Defaulting to plotting x as it is in the file name.')
    return filename_cuts, "x 'as it is'"

def plot_colored_data(unique_labels, x_vals, y_vals):
    """Separates data points by their explicit labels and tracks plots dynamically with color."""
    x_selected, x_label = choose_x_axis(x_vals)
    
    label_to_color_idx = {label: idx for idx, label in enumerate(unique_labels)}
    colors = ['k', 'r', 'b', 'g', 'm', 'c', 'y']
    stride_n = len(unique_labels)
    
    fig, ax = plt.subplots()
    
    for i, current_label in enumerate(unique_labels):
        # Isolate y data targeted specifically to this sequenced label group
        y_slice = list(itertools.islice(y_vals, i, None, stride_n))
        
        # Cross-reference indices back to the master list to fetch paired X variables
        x_slice = []
        for y_val in y_slice:
            master_idx = y_vals.index(y_val)
            x_slice.append(x_selected[master_idx])
            
        print(f"\nGroup [{current_label}] selected points:\nx = {x_slice}\ny = {y_slice}")
        
        color_key = colors[label_to_color_idx[current_label] % len(colors)]
        ax.scatter(x_slice, y_slice, c=color_key, s=30, label=current_label, edgecolors='k')
        
    ax.legend()
    plt.xlabel(x_label)
    plt.ylabel('Signal (mV)')
    plt.title('Plots\nRaw Data Summary')
    plt.show()

# ==========================================
#                  MAIN
# ==========================================
if __name__ == "__main__":
    #  SELECT DATA FOLDER
    folder_selected = ''
    while not folder_selected:
        input('\nPress ENTER to select the folder containing your raw data files...')
        folder_selected = filedialog.askdirectory(initialdir="/")
        if not folder_selected:
            print("No folder selected.")

    print(f"\n> Source Folder selected:\n{folder_selected}")
    
    # Gather pathing and read raw legacy spreadsheet lists (.xls)
    search_path = os.path.join(folder_selected, '*.xls')
    all_file_paths = glob.glob(search_path)
    
    records = []
    unique_labels = []

    if not all_file_paths:
        print("\n[ERROR] No matches found. Check your file extensions inside the target folder.")
        print("Skipping processing since no files were detected.")
    else:
        print('\nParsing and acquiring target files:')
        for path in all_file_paths:
            base_filename = os.path.basename(path)
            print(f" - {base_filename}")
            
            # Standard parsing string splitting expectation: "Value_Label_ExtraInfo.xls"
            name_without_ext, _ = os.path.splitext(base_filename)
            name_parts = name_without_ext.split('_')
            
            file_cut = name_parts[0]
            file_label = name_parts[1] if len(name_parts) > 1 else "Unknown"
            
            if file_label not in unique_labels:
                unique_labels.append(file_label)
                
            # Collect matrix data from the spreadsheet rows
            file_date, raw_matrix_data = get_file_data(folder_selected, base_filename)
            
            # Calculate statistics metrics
            avg_val = mean(raw_matrix_data) if raw_matrix_data else 0.0
            std_val = stdev(raw_matrix_data) if len(raw_matrix_data) > 1 else 0.0
            
            # Organize data row elements into clear dictionary frames
            records.append({
                'filename': base_filename,
                'filename_cut': file_cut,
                'label': file_label,
                'date': file_date,
                'avg': avg_val,
                'std': std_val,
                'delta_time': 0 # Placeholder filled in next step
            })

    # Only execute calculations, saves, and plots if we successfully extracted records
    if records:
        # Add relative timeline deltas to dataset
        all_dates = [r['date'] for r in records]
        calculated_deltas = calculate_delta_times(all_dates)
        for record, delta in zip(records, calculated_deltas):
            record['delta_time'] = delta

        # 2. SELECT SAVE DESTINATION FOLDER
        input('\nPress ENTER to select the directory where you want to save the final report...')
        save_folder_selected = filedialog.askdirectory(initialdir=folder_selected) # Defaults to where your data is
        if not save_folder_selected:
            print("No save directory chosen. Defaulting to script directory.")
            save_folder_selected = os.getcwd()

        # Save data structure matrix out to flat txt files
        last_folder_node = os.path.basename(os.path.normpath(folder_selected))
        save_final_data(save_folder_selected, last_folder_node, records)

        # --- Filtering Numeric Plots ---
        valid_x_floats = []
        valid_y_averages = []

        for r in records:
            # Strip out non-numeric strings (like "background") from running plot visuals
            if not r['filename_cut'].isalpha():
                valid_x_floats.append(float(r['filename_cut']))
                valid_y_averages.append(r['avg'])

        # Render data graphics visually
        plot_colored_data(unique_labels, valid_x_floats, valid_y_averages)
    else:
        print("\nExecution finished successfully without generating plots or reports.")
