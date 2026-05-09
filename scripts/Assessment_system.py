import pandas as pd
import os
from xlsxwriter.utility import xl_col_to_name

# Configuration
input_path = r"C:\Users\Hp\Downloads\UDISE_plus-16-17_infra_kanpur_dehat_up.csv"
output_file = 'Assessment_system_report.xlsx'

def process_infrastructure_data(path):
    if not os.path.exists(path):
        print(f"Error: File not found at {path}")
        return None

    df = pd.read_csv(path)

    # 1. Calculate metrics
    metrics = ['Functional_Drinking_Water', 'Functional_Electricity', 'Functional_Toilet_Facility']
    total_col = 'Total_Number_of_Schools'

    for col in metrics:
        df[f'{col}_Saturation'] = (df[col] / df[total_col]) * 100

    # 2. Generate Infrastructure Score and Grading
    sat_cols = [f'{c}_Saturation' for c in metrics]
    df['Infra_Score'] = df[sat_cols].mean(axis=1)

    def get_grade(score):
        if score >= 90: return 'A'
        if score >= 70: return 'B'
        return 'C'

    df['Final_Grade'] = df['Infra_Score'].apply(get_grade)
    
    # 3. Map Action Plan
    plan_map = {
        'A': 'Maintain & Digitalize',
        'B': 'Plug Infrastructure Gaps',
        'C': 'Urgent Budget Allocation Needed'
    }
    df['Action_Plan'] = df['Final_Grade'].map(plan_map)
    
    return df

def export_to_excel(df, filename):
    try:
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Assessment')

        workbook = writer.book
        worksheet = writer.sheets['Assessment']

        # Define styles
        formats = {
            'A': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}),
            'B': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}),
            'C': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        }

        # Locate Final_Grade column
        grade_idx = df.columns.get_loc("Final_Grade")
        col_letter = xl_col_to_name(grade_idx)
        row_count = len(df) + 1

        # Apply formatting strictly to the Grade column
        for grade, fmt in formats.items():
            worksheet.conditional_format(f'{col_letter}2:{col_letter}{row_count}', {
                'type':     'cell',
                'criteria': 'equal to',
                'value':    f'"{grade}"',
                'format':   fmt
            })

        writer.close()
        print(f"Success: Report generated at {filename}")

    except PermissionError:
        print(f"Error: Close '{filename}' before running the script.")

# Execute
data = process_infrastructure_data(input_path)
if data is not None:
    export_to_excel(data, output_file)
