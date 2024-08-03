import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class JobManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Job Management System")

        # Menu Bar
        self.menu_bar = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Add New Job", command=self.open_add_job_dialog)
        self.file_menu.add_command(label="Save", command=self.save_job)
        self.file_menu.add_command(label="Print Undone Jobs to PDF", command=self.print_pdf)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.root.quit)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="About")
        self.help_menu.add_command(label="Help")
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
        
        self.root.config(menu=self.menu_bar)

        # Job List Section
        self.job_list_frame = tk.Frame(self.root)
        self.job_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.columns = ("Sign Off Date", "Name", "Phone Number", "Location", "Production Date", "Price", "Notes", "Job Number", "Days in Shop", "Status")
        self.job_tree = ttk.Treeview(self.job_list_frame, columns=self.columns, show="headings")
        for col in self.columns:
            self.job_tree.heading(col, text=col)
            self.job_tree.column(col, width=100, anchor=tk.W)
        self.job_tree.pack(fill=tk.BOTH, expand=True)

        # Load data from Excel file
        self.load_jobs_from_excel("jobs.xlsx")

        # Bind selection event
        self.job_tree.bind("<<TreeviewSelect>>", self.on_job_select)

        # Job Details Section
        self.details_frame = tk.LabelFrame(self.root, text="Job Details")
        self.details_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Creating a frame for each detail
        self.entries = {}
        self.labels = ["Sign Off Date", "Name", "Phone Number", "Location", "Production Date", "Price", "Notes", "Job Number"]
        
        for label in self.labels:
            frame = tk.Frame(self.details_frame)
            frame.pack(fill=tk.X, padx=5, pady=2)
            lbl = tk.Label(frame, text=label, width=15)
            lbl.pack(side=tk.LEFT, padx=5)
            if label == "Notes":
                entry = tk.Text(frame, height=4, state=tk.DISABLED)
            else:
                entry = tk.Entry(frame, state=tk.DISABLED)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            self.entries[label] = entry

        # Buttons
        self.button_frame = tk.Frame(self.details_frame)
        self.button_frame.pack(fill=tk.X, padx=5, pady=5)
        self.edit_button = tk.Button(self.button_frame, text="Edit Job", command=self.enable_editing)
        self.edit_button.pack(side=tk.LEFT, padx=5)
        self.save_button = tk.Button(self.button_frame, text="Save Changes", command=self.save_job, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)
        self.mark_done_button = tk.Button(self.button_frame, text="Mark as Done", command=self.mark_done)
        self.mark_done_button.pack(side=tk.LEFT, padx=5)
        self.mark_not_done_button = tk.Button(self.button_frame, text="Mark as Not Done", command=self.mark_not_done)
        self.mark_not_done_button.pack(side=tk.LEFT, padx=5)
        self.delete_button = tk.Button(self.button_frame, text="Delete Job", command=self.delete_job, state=tk.DISABLED)
        self.delete_button.pack(side=tk.LEFT, padx=5)

        # Footer Section
        self.status_bar = tk.Label(self.root, text="Status: Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.file_path = "jobs.xlsx"  # File path for Excel

    def load_jobs_from_excel(self, file_path):
        self.df = pd.read_excel(file_path)
        self.update_treeview()

    def on_job_select(self, event):
        selected_item = self.job_tree.selection()[0]
        selected_job = self.job_tree.item(selected_item, "values")
        job_number = selected_job[6]  # Assuming Job Number is in the 7th column

        try:
            job_status = self.df[self.df["Job Number"] == job_number]["Status"].values[0]
        except IndexError:
            job_status = "Unknown"

        self.sign_off_date_var.set(selected_job[0])
        self.name_var.set(selected_job[1])
        self.phone_number_var.set(selected_job[2])
        self.location_var.set(selected_job[3])
        self.production_date_var.set(selected_job[4])
        self.price_var.set(selected_job[5])
        self.notes_var.set(selected_job[6])
        self.job_number_var.set(selected_job[7])

        if job_status == "Done":
            self.mark_done_button.config(state=tk.DISABLED)
            self.mark_not_done_button.config(state=tk.NORMAL)
        else:
            self.mark_done_button.config(state=tk.NORMAL)
            self.mark_not_done_button.config(state=tk.DISABLED)


    def update_status_buttons(self, status):
        if status == "Done":
            self.mark_done_button.config(state=tk.DISABLED)
            self.mark_not_done_button.config(state=tk.NORMAL)
        else:
            self.mark_done_button.config(state=tk.NORMAL)
            self.mark_not_done_button.config(state=tk.DISABLED)

    def enable_editing(self):
        for label in self.labels:
            self.entries[label].config(state=tk.NORMAL)
        self.save_button.config(state=tk.NORMAL)
        self.edit_button.config(state=tk.DISABLED)

    def save_job(self):
        selected_items = self.job_tree.selection()  # Get selected item IDs
        if not selected_items:  # Check if no items are selected
            self.status_bar.config(text="Status: No job selected")
            return

        selected_item = selected_items[0]  # Get the first selected item ID
        job_number = self.job_tree.item(selected_item, 'values')[self.labels.index("Job Number")]

        # Update the DataFrame with the new values from the entry widgets
        updated_values = {label: self.entries[label].get(1.0, tk.END).strip() if label == "Notes" else self.entries[label].get() for label in self.labels}
        
        # Update DataFrame
        self.df.loc[self.df["Job Number"] == job_number, updated_values.keys()] = updated_values.values()
        
        # Save to Excel
        self.save_to_excel()
        
        # Update Treeview
        self.update_treeview()

        # Clear job details after saving
        self.clear_job_details()

        # Feedback to the user
        self.status_bar.config(text="Status: Job details updated")

    def mark_done(self):
        selected_items = self.job_tree.selection()  # Get selected item IDs
        if not selected_items:  # Check if no items are selected
            self.status_bar.config(text="Status: No job selected")
            return

        selected_item = selected_items[0]  # Get the first selected item ID
        job_number = self.job_tree.item(selected_item, 'values')[self.labels.index("Job Number")]

        # Check the current status
        current_status = self.df[self.df["Job Number"] == job_number]["Status"].values[0]
        if current_status == "Done":
            self.status_bar.config(text="Status: Job already marked as Done")
            return

        self.df.loc[self.df["Job Number"] == job_number, "Status"] = "Done"
        self.update_status_buttons("Done")
        self.status_bar.config(text="Status: Job marked as Done")

        # Update Excel and Treeview
        self.save_to_excel()
        self.update_treeview()

    def mark_not_done(self):
        selected_items = self.job_tree.selection()  # Get selected item IDs
        if not selected_items:  # Check if no items are selected
            self.status_bar.config(text="Status: No job selected")
            return

        selected_item = selected_items[0]  # Get the first selected item ID
        job_number = self.job_tree.item(selected_item, 'values')[self.labels.index("Job Number")]

        # Check the current status
        current_status = self.df[self.df["Job Number"] == job_number]["Status"].values[0]
        if current_status == "Not Done":
            self.status_bar.config(text="Status: Job already marked as Not Done")
            return

        self.df.loc[self.df["Job Number"] == job_number, "Status"] = "Not Done"
        self.update_status_buttons("Not Done")
        self.status_bar.config(text="Status: Job marked as Not Done")

        # Update Excel and Treeview
        self.save_to_excel()
        self.update_treeview()

    def delete_job(self):
        selected_items = self.job_tree.selection()  # Get selected item IDs
        if not selected_items:  # Check if no items are selected
            self.status_bar.config(text="Status: No job selected")
            return

        selected_item = selected_items[0]  # Get the first selected item ID
        job_number = self.job_tree.item(selected_item, 'values')[self.labels.index("Job Number")]

        # Confirmation dialog
        response = messagebox.askyesno("Delete Job", f"Are you sure you want to delete the job with Job Number {job_number}?")
        if response:  # If user confirms
            self.df = self.df[self.df["Job Number"] != job_number]
            self.save_to_excel()
            self.update_treeview()
            self.clear_job_details()  # Clear job details after deletion
            self.status_bar.config(text=f"Status: Job {job_number} deleted")
        else:  # If user cancels
            self.status_bar.config(text="Status: Deletion canceled")

    def clear_job_details(self):
        for label in self.labels:
            if label == "Notes":
                self.entries[label].config(state=tk.NORMAL)
                self.entries[label].delete(1.0, tk.END)
                self.entries[label].config(state=tk.DISABLED)
            else:
                self.entries[label].config(state=tk.NORMAL)
                self.entries[label].delete(0, tk.END)
                self.entries[label].config(state=tk.DISABLED)

        # Disable Edit and Save buttons and the Delete button
        self.edit_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED)

    def save_to_excel(self):
        self.df.to_excel(self.file_path, index=False)

    def update_treeview(self):
        # Clear the Treeview
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)

        # Insert updated data
        for _, row in self.df.iterrows():
            self.job_tree.insert("", tk.END, values=row.tolist())

    def open_add_job_dialog(self):
        add_job_dialog = AddJobDialog(self)
        self.root.wait_window(add_job_dialog.top)

    def print_pdf(self):
        # Implement your PDF printing functionality here
        pass

class AddJobDialog:
    def __init__(self, parent):
        top = self.top = tk.Toplevel(parent.root)
        self.parent = parent
        self.top.title("Add New Job")

        self.entries = {}
        self.labels = [
            ("Sign Off Date (YYYY-MM-DD)", "YYYY-MM-DD"),
            ("Name", "Enter name here"),
            ("Phone Number", "Enter phone number here"),
            ("Location", "Enter location here"),
            ("Production Date (YYYY-MM-DD)", "YYYY-MM-DD"),
            ("Price", "Enter price here"),
            ("Notes", "Enter notes here"),
            ("Job Number", "Enter job number here")
        ]

        for label_text, placeholder in self.labels:
            frame = tk.Frame(self.top)
            frame.pack(fill=tk.X, padx=5, pady=2)
            lbl = tk.Label(frame, text=label_text.split()[0], width=15)
            lbl.pack(side=tk.LEFT, padx=5)
            if "Notes" in label_text:
                entry = tk.Text(frame, height=4)
                entry.insert(tk.END, placeholder)
                entry.bind("<FocusIn>", self.clear_placeholder_text)
                entry.bind("<FocusOut>", self.add_placeholder_text)
            else:
                entry = tk.Entry(frame)
                entry.insert(0, placeholder)
                entry.bind("<FocusIn>", self.clear_placeholder_text)
                entry.bind("<FocusOut>", self.add_placeholder_text)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            self.entries[label_text.split()[0]] = entry

        button_frame = tk.Frame(self.top)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        add_button = tk.Button(button_frame, text="Add Job", command=self.add_job)
        add_button.pack(side=tk.LEFT, padx=5)
        cancel_button = tk.Button(button_frame, text="Cancel", command=self.top.destroy)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def clear_placeholder_text(self, event):
        widget = event.widget
        if isinstance(widget, tk.Text):
            if widget.get("1.0", tk.END).strip() in [placeholder for _, placeholder in self.labels]:
                widget.delete("1.0", tk.END)
        else:
            if widget.get() in [placeholder for _, placeholder in self.labels]:
                widget.delete(0, tk.END)

    def add_placeholder_text(self, event):
        widget = event.widget
        if isinstance(widget, tk.Text):
            if widget.get("1.0", tk.END).strip() == "":
                for _, placeholder in self.labels:
                    if placeholder in widget.get("1.0", tk.END):
                        widget.insert(tk.END, placeholder)
                        break
        else:
            if widget.get() == "":
                for _, placeholder in self.labels:
                    if placeholder in widget.get():
                        widget.insert(0, placeholder)
                        break

    def add_job(self):
        # Extract values from entries
        new_job = {
            "Sign Off Date": self.entries["Sign"].get().strip(),
            "Name": self.entries["Name"].get().strip(),
            "Phone Number": self.entries["Phone"].get().strip(),
            "Location": self.entries["Location"].get().strip(),
            "Production Date": self.entries["Production"].get().strip(),
            "Price": self.entries["Price"].get().strip(),
            "Notes": self.entries["Notes"].get("1.0", tk.END).strip(),
            "Job Number": self.entries["Job"].get().strip(),
            "Status": "Not Done"
        }

        # Validate the dates
        try:
            sign_off_date = pd.to_datetime(new_job["Sign Off Date"], format="%Y-%m-%d")
            production_date = pd.to_datetime(new_job["Production Date"], format="%Y-%m-%d")
            new_job["Sign Off Date"] = sign_off_date.strftime("%Y-%m-%d")
            new_job["Production Date"] = production_date.strftime("%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Invalid Date Format", "Dates must be in YYYY-MM-DD format.")
            return

        # Calculate Days in Shop
        days_in_shop = (pd.Timestamp.now() - production_date).days
        new_job["Days in Shop"] = days_in_shop

        # Convert new job to DataFrame
        new_job_df = pd.DataFrame([new_job])

        # Concatenate the new job to the DataFrame
        self.parent.df = pd.concat([self.parent.df, new_job_df], ignore_index=True)

        # Save to Excel and update Treeview
        self.parent.save_to_excel()
        self.parent.update_treeview()

        self.top.destroy()

# Running the application
if __name__ == "__main__":
    root = tk.Tk()
    app = JobManagementApp(root)
    root.mainloop()
