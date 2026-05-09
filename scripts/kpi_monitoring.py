import pandas as pd
import os

# Configuration
file_path = r"C:\Users\Hp\Downloads\UDISE_plus-16-17_infra_kanpur_dehat_up.csv"
output_name = "Kanpur_Dehat_KPI_Monitor.xlsx"

def generate_monitoring_report(path):
    if not os.path.exists(path):
        print(f"Error: File not found at {path}")
        return

    try:
        df = pd.read_csv(path)
        # Clean column names to prevent KeyErrors
        df.columns = df.columns.str.strip()
        print(f"Dataset loaded: {len(df)} records found.")
    except Exception as e:
        print(f"Process failed during file read: {e}")
        return

    # --- KPI Calculation Logic ---

    # KPI 1: Sanitation Parity (Girl vs Boy toilets, capped at 1.0)
    df['Sanitation_Parity'] = (df['Functional_Girl_Toilet'] / (df['Functional_Boy_Toilet'] + 0.1)).clip(upper=1)

    # KPI 2: Digital Readiness (Computers and Internet availability)
    df['Digital_Readiness'] = ((df['Computer_Available'] + df['Internet']) / (df['Total_Number_of_Schools'] * 2)) * 100

    # KPI 3: Basic Infrastructure (Water, Electricity, and Building status)
    df['Basic_Infra_Saturation'] = ((df['Functional_Drinking_Water'] + df['Functional_Electricity'] + df['Building']) / (df['Total_Number_of_Schools'] * 3)) * 100

    # KPI 4: Quality Learning Environment (Library, Playground, and Furniture)
    df['Quality_Index'] = ((df['Library_or_Reading_Corner_or_Book_Bank'] + df['Playground'] + df['Furniture']) / (df['Total_Number_of_Schools'] * 3)) * 100

    # --- Aggregation & Monitoring ---

    # Aggregate performance by Administrative Block
    report = df.groupby('Udise_Block_Name').agg({
        'Sanitation_Parity': 'mean',
        'Digital_Readiness': 'mean',
        'Basic_Infra_Saturation': 'mean',
        'Quality_Index': 'mean'
    }).reset_index()

    # Define status based on fundamental infrastructure threshold (75%)
    report['Status'] = report['Basic_Infra_Saturation'].apply(
        lambda x: 'Satisfactory' if x >= 75 else 'Immediate Action Required'
    )

    # --- Export ---
    try:
        report.to_excel(output_name, index=False)
        print(f"Success: Monitoring Dashboard generated as '{output_name}'")
    except PermissionError:
        print(f"Error: Close '{output_name}' before running the script.")

if __name__ == "__main__":
    generate_monitoring_report(file_path)
