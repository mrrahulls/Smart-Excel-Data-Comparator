import pandas as pd
import glob
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkfont
from datetime import datetime
import sys

class RedirectText(object):
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)  # Scroll to the end as new lines are added

    def flush(self):  # Needed for Python's handling of stdout
        pass

stop_flag = False  # Initialize stop flag

def stop_search():
    """Set stop_flag to True to stop the process."""
    global stop_flag
    stop_flag = True
    print("Stopping the process...")

def browse_data_folder():
    folder_selected = filedialog.askdirectory()
    data_folder_var.set(folder_selected)

def browse_reference_folder():
    folder_selected = filedialog.askdirectory()
    reference_folder_var.set(folder_selected)

def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder_var.set(folder_selected)

def start_search():
    global stop_flag
    stop_flag = False  
    data_folder = data_folder_var.get()
    reference_folder = reference_folder_var.get()
    output_folder = output_folder_var.get()
    
    if not (data_folder and reference_folder and output_folder):
        messagebox.showerror("Error", "Please select all folders.")
        return

    messagebox.showinfo("Info", "Searching... Please wait.")
    
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    data_files = sorted(glob.glob(os.path.join(data_folder, '*.xlsx')) + glob.glob(os.path.join(data_folder, '*.csv')))
    reference_files = sorted(glob.glob(os.path.join(reference_folder, '*.xlsx')) + glob.glob(os.path.join(reference_folder, '*.csv')))
    
    # Create a DataFrame to store all combined results
    all_combined_results = pd.DataFrame()
    total_overall_matches = 0  # Track overall matches

    # Loop over each data file
    for data_file_index, data_file in enumerate(data_files, start=1):
        if stop_flag:
            print("Process stopped by user.")
            break  # Exit loop if stop flag is set

        if data_file.endswith('.xlsx'):
            data_df = pd.read_excel(data_file)
        elif data_file.endswith('.csv'):
            data_df = pd.read_csv(data_file)

        data_file_name = os.path.basename(data_file)  # Get the name of the current data file
        total_matches = 0  # Initialize match counter for the current data file

        print(f"Processing file {data_file_index}/{len(data_files)}: '{data_file_name}'...")
        root.update()  # Update the GUI to stay responsive

        # Loop over each reference file
        for ref_file_index, ref_file in enumerate(reference_files, start=1):
            if stop_flag:
                print("Process stopped by user.")
                break  # Exit loop if stop flag is set

            if ref_file.endswith('.xlsx'):
                reference_df = pd.read_excel(ref_file)
            elif ref_file.endswith('.csv'):
                reference_df = pd.read_csv(ref_file)

            print(f"  Comparing with reference file {ref_file_index}/{len(reference_files)}: '{os.path.basename(ref_file)}'...")
            root.update()

            # Ensure case-insensitive matching by converting 'Name' and 'Father Name' to lowercase
            data_df['Name'] = data_df['Name'].str.lower()
            data_df['Father Name'] = data_df['Father Name'].str.lower()

            reference_df['Name'] = reference_df['Name'].str.lower()
            reference_df['Father Name'] = reference_df['Father Name'].str.lower()

            # Merge on "Name" and "Father Name"
            merged_df = pd.merge(reference_df, data_df, how='inner', on=['Name', 'Father Name'])

            # Filter out rows where either "Name" or "Father Name" is missing or NaN
            filtered_df = merged_df[merged_df['Name'].notna() & merged_df['Father Name'].notna() & 
                                    (merged_df['Father Name'] != '')]

            # If there are valid matches, add the reference file name to the "findout" column
            if not filtered_df.empty:
                filtered_df['findout'] = os.path.basename(ref_file)  # Add the reference file name
                
                # Add a new column to show both file names
                filtered_df['file_names'] = f"{data_file_name} | {os.path.basename(ref_file)}"  # Combine file names
                
                # Save the matches
                all_combined_results = pd.concat([all_combined_results, filtered_df], ignore_index=True)
                matches_count = filtered_df.shape[0]  # Get number of matches found
                total_matches += matches_count  # Increment total matches for current file
                print(f"    Found {matches_count} valid matches in '{data_file_name}' from reference file '{os.path.basename(ref_file)}'.")
                root.update()

        # Print total matches found for the current data file
        if total_matches > 0:
            total_overall_matches += total_matches  # Add to overall matches
            print(f"Total valid matches found in '{data_file_name}': {total_matches}.")
        else:
            print(f"No valid matches found in '{data_file_name}'.")
        root.update()

        # Save partial results and exit if stop flag is set
        if stop_flag:
            print("Saving progress and exiting.")
            save_partial_results(output_folder, all_combined_results)
            break

    # If the process completed without stopping, save the final results
    if not stop_flag and not all_combined_results.empty:
        save_final_results(output_folder, all_combined_results, total_overall_matches)
    elif not stop_flag:
        print("No valid matches found where both 'Name' and 'Father Name' are present.")
        messagebox.showinfo("Info", "No valid matches found where both 'Name' and 'Father Name' are present.")
    
    print("Search process completed.")
    messagebox.showinfo("Info", "Search process completed.")

def save_partial_results(output_folder, all_combined_results):
    """Save partially matched data when process is stopped."""
    if not all_combined_results.empty:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(output_folder, f'partial_matched_data_{timestamp}.xlsx')
        all_combined_results.to_excel(output_path, index=False)
        print(f"Partial results saved to '{output_path}'.")

def save_final_results(output_folder, all_combined_results, total_overall_matches):
    """Save final matched data after the full process."""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')  # Generate a unique timestamp
    output_path = os.path.join(output_folder, f'matched_data_{timestamp}.xlsx')  # Use timestamp in filename
    all_combined_results.to_excel(output_path, index=False)
    print(f"Matching process completed and saved to '{output_path}'.")
    messagebox.showinfo("Info", f"Matching process completed. {total_overall_matches} total matches found and saved to '{output_path}'.")


root = tk.Tk()
root.title("Excel Data Matcher - Smart Match Tool")
root.geometry("600x600")  
root.configure(bg='#ffffff')  


title_font = tkfont.Font(family="Arial", size=20, weight="bold")
title_label = tk.Label(root, text="Excel Data Matcher", font=title_font, bg='#ffffff', fg='#4CAF50')
title_label.pack(pady=(20, 10))


data_folder_var = tk.StringVar()
reference_folder_var = tk.StringVar()
output_folder_var = tk.StringVar()


frame = tk.Frame(root, bg='#f0f0f0', bd=5, relief=tk.RAISED)
frame.pack(pady=20, padx=10)

# Create and place widgets with enhanced styling
label_font = tkfont.Font(family="Arial", size=12)
entry_font = tkfont.Font(family="Arial", size=12)

tk.Label(frame, text="Select Data Folder:", bg='#f0f0f0', font=label_font).grid(row=0, column=0, padx=5, pady=10, sticky='w')
tk.Entry(frame, textvariable=data_folder_var, width=40, font=entry_font).grid(row=0, column=1, padx=5, pady=10)
tk.Button(frame, text="Browse", command=browse_data_folder, font=('Arial', 10), relief=tk.RAISED).grid(row=0, column=2, padx=5, pady=10)

tk.Label(frame, text="Select Reference Folder:", bg='#f0f0f0', font=label_font).grid(row=1, column=0, padx=5, pady=10, sticky='w')
tk.Entry(frame, textvariable=reference_folder_var, width=40, font=entry_font).grid(row=1, column=1, padx=5, pady=10)
tk.Button(frame, text="Browse", command=browse_reference_folder, font=('Arial', 10), relief=tk.RAISED).grid(row=1, column=2, padx=5, pady=10)

tk.Label(frame, text="Select Output Folder:", bg='#f0f0f0', font=label_font).grid(row=2, column=0, padx=5, pady=10, sticky='w')
tk.Entry(frame, textvariable=output_folder_var, width=40, font=entry_font).grid(row=2, column=1, padx=5, pady=10)
tk.Button(frame, text="Browse", command=browse_output_folder, font=('Arial', 10), relief=tk.RAISED).grid(row=2, column=2, padx=5, pady=10)


button_frame = tk.Frame(root, bg='#ffffff')
button_frame.pack(pady=10)

tk.Button(button_frame, text="Start", command=start_search, font=('Arial', 14), relief=tk.RAISED).pack(side=tk.LEFT, padx=10, pady=10)
tk.Button(button_frame, text="Stop", command=stop_search, font=('Arial', 14), relief=tk.RAISED, bg='#f44336', fg='#ffffff').pack(side=tk.LEFT, padx=10, pady=10)


log_text = tk.Text(root, height=12, wrap=tk.WORD, font=('Arial', 10), bg='#f7f7f7')
log_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)


redirect_text = RedirectText(log_text)
sys.stdout = redirect_text


credit_font = tkfont.Font(family="Arial", size=10, slant="italic")
credit_label = tk.Label(root, text="Designed by Mr. Rahul Singh", font=credit_font, bg='#ffffff', fg='#555555')
credit_label.pack(side=tk.BOTTOM, pady=(10, 20))


root.mainloop()
