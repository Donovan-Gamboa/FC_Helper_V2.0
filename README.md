
# Job Management System

This project is a simple GUI-based job management application built using Python's `tkinter` for the user interface and `pandas` for data handling. It allows users to manage a list of jobs, update job details, mark jobs as done or not done, and save this data to an Excel file. Additionally, users can export a list of undone jobs to a PDF file.

## Features
- **Add New Job**: Add a new job entry.
- **Edit and Update Jobs**: Modify job details and save changes.
- **Mark as Done / Not Done**: Update the status of jobs.
- **Delete Job**: Remove a job from the list.
- **Print Undone Jobs to PDF**: Export a list of jobs that are not marked as done.
- **Excel Integration**: Load job data from an Excel file (`jobs.xlsx`) and save any changes back to it.

## Dependencies
- **Python 3.x**
- **Tkinter**: Built-in GUI library in Python.
- **Pandas**: For managing job data in a DataFrame (`pip install pandas`).
- **ReportLab**: For creating PDF files (`pip install reportlab`).

## Installation
1. Clone the repository or download the code files.
2. Install the required dependencies:
   ```bash
   pip install pandas reportlab
   ```
3. Ensure that `jobs.xlsx` (an Excel file with job data) exists in the same directory.

## Usage
1. Run the script:
   ```bash
   python job_management_app.py
   ```
2. The main window will open with a job list and detailed view for job management.

### Main Interface Sections
1. **Menu Bar**: Options to add new jobs, print undone jobs to PDF, or exit the application.
2. **Job List**: Displays all jobs, including columns like name, phone number, location, etc.
3. **Job Details**: Shows detailed information for a selected job, allowing edits and updates.
4. **Buttons**: Options to edit, save, mark as done/not done, and delete jobs.
5. **Status Bar**: Provides feedback on actions taken within the app.

## File Descriptions
- **jobs.xlsx**: The Excel file where job data is stored and loaded.
- **job_management_app.py**: The main Python file containing the code for the Job Management System.

## PDF Export
The app includes a feature to export all jobs that are not marked as done into a PDF. The PDF lists each job's details and is generated in landscape format.

## Contributing
Feel free to submit issues or pull requests for improvements.

## License
This project is open-source and available under the MIT License.
