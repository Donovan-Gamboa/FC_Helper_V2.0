import pandas as pd
from random import randint
from datetime import datetime, timedelta

# Generate random dates within the specified range
def generate_random_date(start_date, end_date):
    delta = end_date - start_date
    random_days = randint(0, delta.days)
    return start_date + timedelta(days=random_days)

start_date = datetime.strptime("2024-01-15", "%Y-%m-%d")
end_date = datetime.strptime("2024-07-20", "%Y-%m-%d")

# Create a DataFrame with the dummy job data
data = {
    "Sign Off Date": [generate_random_date(start_date, end_date).strftime("%Y-%m-%d") for _ in range(15)],
    "Name": [
        "Alice Smith", "Bob Johnson", "Carol White", "David Brown", "Eve Black",
        "Frank Green", "Grace Blue", "Hank Pink", "Ivy Red", "Jack Gold",
        "Kelly Silver", "Leo Copper", "Mona Bronze", "Nina Violet", "Oscar Crimson"
    ],
    "Phone Number": [
        "555-5678", "555-8765", "555-4321", "555-1234", "555-6789",
        "555-2468", "555-1357", "555-9876", "555-6543", "555-3210",
        "555-2468", "555-7531", "555-3698", "555-1597", "555-7530"
    ],
    "Location": [
        "456 Maple Avenue, Springfield", "789 Oak Street, Springfield", "321 Pine Lane, Springfield", "654 Birch Boulevard, Springfield", "987 Cedar Circle, Springfield",
        "123 Aspen Drive, Springfield", "246 Elm Court, Springfield", "369 Willow Way, Springfield", "753 Cherry Road, Springfield", "951 Spruce Street, Springfield",
        "357 Poplar Place, Springfield", "159 Fir Terrace, Springfield", "753 Redwood Row, Springfield", "159 Cypress Crescent, Springfield", "753 Alder Alley, Springfield"
    ],
    "Production Date": [generate_random_date(start_date, end_date).strftime("%Y-%m-%d") for _ in range(15)],
    "Price": [
        "$200.00", "$350.00", "$175.00", "$400.00", "$225.00",
        "$300.00", "$250.00", "$375.00", "$500.00", "$275.00",
        "$320.00", "$340.00", "$360.00", "$380.00", "$400.00"
    ],
    "Notes": [
        "Urgent job completed ahead of schedule. No issues reported.",
        "Job had some delays due to supply chain issues. Extra charges applied.",
        "Completed on time. Minor adjustments requested post-delivery.",
        "Quality assurance passed with no issues.",
        "Customer requested an additional feature, causing a delay.",
        "Job completed smoothly without any problems.",
        "Some rework needed, but job finished on time.",
        "Customer was very satisfied with the outcome.",
        "Minor issues detected, quickly resolved.",
        "Job was delayed due to inclement weather.",
        "Client requested early delivery, accommodated successfully.",
        "Job went as planned, no deviations.",
        "Slight delay due to material shortage.",
        "Finished ahead of schedule, no problems.",
        "Customer provided positive feedback, no issues."
    ],
    "Job Number": [f"J{str(i).zfill(4)}" for i in range(1, 16)],
    "Days in Shop": [str(randint(0, 90)) for _ in range(15)],
    "Status": ["Not Done" if i % 2 == 0 else "Done" for i in range(15)]
}

df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
df.to_excel("jobs.xlsx", index=False)
