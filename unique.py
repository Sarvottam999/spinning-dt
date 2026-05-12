import pandas as pd

# Load Excel file
df = pd.read_excel("/home/sarvottam/Documents/pasted2.xlsx")

selected_columns = [
    "Plant",
    "Department",
    "Section",
    "Functional Location",
    "Head reason",
    "Sub Head reason",
    "Reason Desc"
]

# Get unique values
for col in selected_columns:
    unique_values = df[col].dropna().unique()

    print(f"\n===== {col} =====")
    print(unique_values)