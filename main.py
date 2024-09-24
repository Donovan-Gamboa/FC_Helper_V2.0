import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch


class JobManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Job Management System")

        # Menu Bar
        self.menu_bar = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Add New Job", command=self.open_add_job_dialog)
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
        
        self.columns = ("Sign Off Date", "Name", "Phone Number", "Location", "Production Date", "Price", "Notes", "Job Number", "Status", "Days in Shop")
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

        self.job_tree.tag_configure('done', background='lightgray')
        self.job_tree.tag_configure('not_done', background='white')

        self.file_path = "jobs.xlsx"  # File path for Excel

    def load_jobs_from_excel(self, file_path):
        # Load data from Excel file
        self.df = pd.read_excel(file_path)

        # Convert 'Production Date' column to datetime
        self.df['Production Date'] = pd.to_datetime(self.df['Production Date'], format="%Y-%m-%d")
        self.df['Sign Off Date'] = pd.to_datetime(self.df['Sign Off Date'], format="%Y-%m-%d").dt.normalize()  # Remove time component
        
        # Calculate 'Days in Shop'
        current_date = pd.Timestamp.now().normalize()  # Remove time component
        self.df['Days in Shop'] = (current_date - self.df['Production Date']).dt.days
        
        # Ensure 'Days in Shop' is an integer
        self.df['Days in Shop'] = self.df['Days in Shop'].astype(int)

        self.df['Production Date'] = self.df['Production Date'].dt.strftime("%Y-%m-%d")
        self.df['Sign Off Date'] = self.df['Sign Off Date'].dt.strftime("%Y-%m-%d")

        self.df['Notes'] = self.df['Notes'].fillna("")
        self.df['Job Number'] = self.df['Job Number'].fillna("")


        # Update Treeview
        self.update_treeview()

    def on_job_select(self, event):
        selected_items = self.job_tree.selection()  # Get selected item IDs
        if not selected_items:  # Check if no items are selected
            return

        selected_item = selected_items[0]  # Get the first selected item ID
        job = self.job_tree.item(selected_item, 'values')  # Get job data

        # Populate Job Details Section
        for i, label in enumerate(self.labels):
            if label == "Notes":
                self.entries[label].config(state=tk.NORMAL)
                self.entries[label].delete(1.0, tk.END)  # Clear the existing text
                self.entries[label].insert(tk.END, job[i])
                self.entries[label].config(state=tk.DISABLED)
            else:
                self.entries[label].config(state=tk.NORMAL)
                self.entries[label].delete(0, tk.END)  # Clear the existing text
                self.entries[label].insert(0, job[i])
                self.entries[label].config(state=tk.DISABLED)

        # Enable Edit and Delete buttons and disable Save button
        self.edit_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.NORMAL)

        # Get the job status and update buttons
        job_number = job[self.labels.index("Job Number")]
        job_status = self.df[self.df["Job Number"].astype(str).str.strip() == job_number]["Status"].values[0]
        self.update_status_buttons(job_status)

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
        self.df.loc[self.df["Job Number"].astype(str).str.strip() == job_number, updated_values.keys()] = updated_values.values()
        
        # Save to Excel
        self.save_to_excel()

        self.load_jobs_from_excel("jobs.xlsx")
        
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
        current_status = self.df[self.df["Job Number"].astype(str).str.strip() == job_number]["Status"].values[0]
        if current_status == "Done":
            self.status_bar.config(text="Status: Job already marked as Done")
            return

        self.df.loc[self.df["Job Number"].astype(str).str.strip() == job_number, "Status"] = "Done"
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
        current_status = self.df[self.df["Job Number"].astype(str).str.strip() == job_number]["Status"].values[0]
        if current_status == "Not Done":
            self.status_bar.config(text="Status: Job already marked as Not Done")
            return

        self.df.loc[self.df["Job Number"].astype(str).str.strip() == job_number, "Status"] = "Not Done"
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
            self.df = self.df[self.df["Job Number"].astype(str).str.strip() != job_number]
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
        def get_gradient_color(days):
            if days <= 45:
                red = int((days / 45) * 255)
                green = 255
                blue = 0
                return f'#{red:02x}{green:02x}{blue:02x}'
            elif (days > 45) and (days <= 90):
                red = 255
                green = 255 - int(((days - 45) / 45) * 255)
                blue = 0
                return f'#{red:02x}{green:02x}{blue:02x}'
            elif (days >= 90):
                return '#ff0000'
            

        # Clear the Treeview
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)

        # Sort DataFrame by 'Status' and then by 'Days in Shop' in ascending order
        self.df['Status'] = pd.Categorical(self.df['Status'], categories=['Not Done', 'Done'], ordered=True)
        sorted_df = self.df.sort_values(by=['Status', 'Days in Shop'], ascending=[True, True])

        # Insert updated data and color-code based on status
        for _, row in sorted_df.iterrows():
            values = row.tolist()
            days_in_shop = row['Days in Shop']
            color = get_gradient_color(days_in_shop)
            tags = (f'days_{days_in_shop}',)
            self.job_tree.tag_configure(f'days_{days_in_shop}', background=color)
            if row['Status'] == 'Done':
                tags = ('done',)
            self.job_tree.insert("", tk.END, values=values, tags=tags)
        
        # Apply tag configuration to gray out 'Done' jobs
        self.job_tree.tag_configure('done', foreground='gray')

    def open_add_job_dialog(self):
        add_job_dialog = AddJobDialog(self)
        self.root.wait_window(add_job_dialog.top)

    def print_pdf(self):
        # Create a PDF with landscape orientation and setup
        pdf_file = "Undone_Jobs_Report.pdf"
        doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))

        # Define styles for the table
        styles = getSampleStyleSheet()
        normal_style = styles["Normal"]

        # Table data: column headers
        data = [["Job Number", "Name", "Phone Number", "Location", "Sign Off Date", "Production Date", "Notes", "Days in Shop"]]

        sorted_df = self.df[self.df["Status"] == "Not Done"].sort_values(by="Days in Shop")


        # Add each job's details to the table with text wrapping
        for _, row in sorted_df.iterrows():
            job_details = [
                Paragraph(str(row["Job Number"]), normal_style),
                Paragraph(str(row["Name"]), normal_style),
                Paragraph(str(row["Phone Number"]), normal_style),
                Paragraph(str(row["Location"]), normal_style),
                Paragraph(str(row["Sign Off Date"]), normal_style),
                Paragraph(str(row["Production Date"]), normal_style),
                Paragraph(str(row["Notes"]), normal_style),
                Paragraph(str(row["Days in Shop"]), normal_style)
            ]
            data.append(job_details)

        # Create the table
        table = Table(data, colWidths=[60, 110, 80, 90, 80, 80, 200, 60])  # Adjust column widths as necessary

        # Style the table
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Header background color
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header text color
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center alignment
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Top align for wrapping content
            ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines
            ('FONTSIZE', (0, 0), (-1, -1), 10),  # Font size for all cells
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Row background color
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),  # Normal text color
        ])

        # Apply the gradient for "Days in Shop" based on the values
        for i in range(1, len(data)):
            days_in_shop = int(sorted_df.iloc[i - 1]["Days in Shop"])

            if days_in_shop <= 30:
                bg_color = colors.green
            elif (days_in_shop > 30) and (days_in_shop <= 70):
                bg_color = colors.yellow
            elif (days_in_shop >= 70):
                bg_color = colors.red

            style.add('BACKGROUND', (7, i), (7, i), bg_color)
        
        table.setStyle(style)

        # Allow the table to split across multiple pages
        elements = [table]

        # Build the document (multi-page support)
        doc.build(elements)

        # Notify user that PDF generation is complete
        messagebox.showinfo("Success", "PDF generated successfully!")

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

        if not new_job["Notes"]:
            new_job["Notes"] = ""  # Set to empty string instead of NaN
        if not new_job["Job Number"]:
            new_job["Job Number"] = ""  # Set to empty string instead of NaN

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


#Price not Needed for PDF
#Grid for PDF if possible