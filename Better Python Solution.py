import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys  # Import sys to allow for program exit

def read_and_process_file():
    # Ask user to select a text file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select a text file", filetypes=[("Text files", "*.txt")])
    if not file_path:
        messagebox.showwarning("No file selected", "Please select a text file.")
        return

    # Initialize lists to store data
    acct_pt = []
    name = []
    shares = []
    last_shares_value = ""
    
    # Read the text file
    with open(file_path, 'r') as file:
        lines = file.readlines()

        for line in lines:
            line = line.strip()
            
            # Check for "Acct/Pt."
            if "Acct/Pt." in line:
                # Split the line at "Acct/Pt." and get everything after it
                after_acct = line.split("Acct/Pt.")[1]
                
                # Extract the part immediately following "Acct/Pt."
                after_acct_parts = after_acct.strip().split()
                acct_pt_value = ""
                
                # Find the first part that is a valid number (digits, with optional hyphens or periods)
                for part in after_acct_parts:
                    if part.replace("-", "").replace(".", "").isdigit():  # Handles numbers with hyphens or periods
                        acct_pt_value = part
                        break
                
                acct_pt.append(acct_pt_value)
                
                # Check if "UPAL -" is in the same line
                if "UPAL -" in line:
                    name_value = after_acct.split("UPAL -")[1].strip()
                    name_value = "UPAL -" + name_value  # Include "UPAL -" in the name
                else:
                    # Extract potential name after the account number
                    name_value = " ".join(after_acct_parts[1:]).strip()  # Take all parts after the number

                    # If extracted value contains any digits, consider it invalid
                    if any(char.isdigit() for char in name_value):
                        name_value = ""
                
                name.append(name_value)

            # Check for "Shares"
            elif "Shares" in line:
                # Find the position of "Shares" and extract everything before it
                shares_position = line.find("Shares")
                
                # Extract the part immediately before "Shares"
                before_shares = line[:shares_position].strip().split()
                
                # Check if the last part before "Shares" is a valid number
                if before_shares:
                    potential_shares_value = before_shares[-1]
                    if potential_shares_value.replace(",", "").replace(".", "", 1).isdigit():
                        last_shares_value = potential_shares_value  # Store last numeric value before "Shares"

            # Check for "Confidential"
            elif "Confidential" in line:
                shares.append(last_shares_value)  # Add the last shares value before "Confidential"
                last_shares_value = ""  # Reset for the next block of data

    # Create a DataFrame
    data = {
        "Acct.Pt": acct_pt,
        "Name": name,
        "Shares": shares
    }
    df = pd.DataFrame(data)

    # Display data in a popup window
    display_data(df)

def display_data(df):
    # Create the main window
    window = tk.Tk()
    window.title("Processed Data")

    # Ensure program ends when the popup window is closed
    window.protocol("WM_DELETE_WINDOW", on_closing)

    # Create a frame for the DataFrame
    frame = ttk.Frame(window)
    frame.pack(fill=tk.BOTH, expand=True)

    # Create a treeview to display the DataFrame
    tree = ttk.Treeview(frame, columns=list(df.columns), show='headings')
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add column headers
    for col in df.columns:
        tree.heading(col, text=col)

    # Add data to the treeview
    for index, row in df.iterrows():
        tree.insert("", "end", values=list(row))

    # Add a scrollbar
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscroll=scrollbar.set)

    # Add export button
    export_button = tk.Button(window, text="Export to Text File", command=lambda: export_to_text_file(df))
    export_button.pack(pady=10)

    window.mainloop()

def export_to_text_file(df):
    # Ask the user where to save the text file
    file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                             filetypes=[("Text files", "*.txt")],
                                             title="Save file as")
    if not file_path:
        return  # If no file path is selected, do nothing

    # Export DataFrame to text file
    try:
        df.to_csv(file_path, sep='\t', index=False)
        messagebox.showinfo("Export Successful", f"Data has been exported to {file_path}.")
    except Exception as e:
        messagebox.showerror("Export Failed", f"An error occurred while exporting the file: {e}")

def on_closing():
    # Function to handle closing the program
    sys.exit()  # Exit the program completely

# Run the script
read_and_process_file()
