import tkinter as tk
from tkinter import ttk
import pandas as pd

# Define the DataFrame to store the data
df = pd.DataFrame(columns=["Enrollment No", "Name", "Branch", "Term"])

def submit_form():
    global df

    enrollment_no = entry_enrollment_no.get()
    name = entry_name.get()
    branch = combobox_branch.get()
    term = combobox_term.get()

    if not enrollment_no or not name or not branch or not term:
        print("All fields are required!")
        return

    data = {
        "Enrollment No": [enrollment_no],
        "Name": [name],
        "Branch": [branch],
        "Term": [term]
    }

    # Create a temporary DataFrame from the data
    temp_df = pd.DataFrame(data)

    # Append the temporary DataFrame to the main DataFrame
    df = pd.concat([df, temp_df], ignore_index=True)

    # Save to Excel
    df.to_excel("data.xlsx", index=False)
    
    # Clear the entry fields
    entry_enrollment_no.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    combobox_branch.set("")
    combobox_term.set("")
    
    print("Data saved successfully!")
    print(f"Enrollment No: {enrollment_no}, Name: {name}, Branch: {branch}, Term: {term}")

# Create the main window
root = tk.Tk()
root.title("Student Registration Form")

# Create labels and entry fields
label_enrollment_no = tk.Label(root, text="Enrollment No:")
label_enrollment_no.grid(row=0, column=0)
entry_enrollment_no = tk.Entry(root)
entry_enrollment_no.grid(row=0, column=1)

label_name = tk.Label(root, text="Name:")
label_name.grid(row=1, column=0)
entry_name = tk.Entry(root)
entry_name.grid(row=1, column=1)

label_branch = tk.Label(root, text="Branch:")
label_branch.grid(row=2, column=0)
combobox_branch = ttk.Combobox(root, values=["Arch.", "A&R", "Auto.", "Bio.", "Civ.", "Com.", "IT", "ICT", "Electrical", "Electronic", "EC", "Mech.", "Plastic", "IC"], state="readonly")
combobox_branch.grid(row=2, column=1)

label_term = tk.Label(root, text="Term:")
label_term.grid(row=3, column=0)
combobox_term = ttk.Combobox(root, values=["1", "2", "3", "4", "5", "6"], state="readonly")
combobox_term.grid(row=3, column=1)

# Create a submit button
submit_button = tk.Button(root, text="Submit", command=submit_form)
submit_button.grid(row=4, columnspan=2)

# Start the main loop
root.mainloop()
