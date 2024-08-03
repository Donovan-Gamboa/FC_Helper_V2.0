import pandas as pd

# Create a DataFrame with the dummy job data
data = {
    "Sign Off Date": ["2024-07-10", "2024-07-22", "2024-07-25"],
    "Name": ["Alice Smith", "Bob Johnson", "Carol White"],
    "Phone Number": ["555-5678", "555-8765", "555-4321"],
    "Location": ["456 Maple Avenue, Springfield", "789 Oak Street, Springfield", "321 Pine Lane, Springfield"],
    "Production Date": ["2024-07-01", "2024-07-10", "2024-07-20"],
    "Price": ["$200.00", "$350.00", "$175.00"],
    "Notes": [
        "Urgent job completed ahead of schedule. No issues reported.",
        "Job had some delays due to supply chain issues. Extra charges applied.",
        "Completed on time. Minor adjustments requested post-delivery."
    ],
    "Job Number": ["J0002", "J0003", "J0004"],
    "Days in Shop": ["9", "17", "14"],
    "Status": ["Not Done", "Done", "Not Done"]
}

df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
df.to_excel("jobs.xlsx", index=False)
