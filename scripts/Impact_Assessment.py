import pandas as pd
import matplotlib.pyplot as plt
import os
from xlsxwriter.utility import xl_col_to_name

# Configuration: File paths (Using v2 files for MDM data)
infra_path    = r"C:\Users\Hp\Downloads\UDISE_plus-16-17_infra_kanpur_dehat_up.csv"
baseline_path = r"C:\Users\Hp\Downloads\Attendance_Baseline_2023_v2.csv"
endline_path  = r"C:\Users\Hp\Downloads\Attendance_Endline_2025_v2.csv"
output_report = "MoE_Final_Comprehensive_Report.xlsx"

def run_integrated_analysis():
    # 1. Validation: Ensure all 3 CSV sources exist
    if not all(os.path.exists(f) for f in [infra_path, baseline_path, endline_path]):
        print("Error: Missing data files. Ensure all v2 CSVs are in the Downloads folder.")
        return

    df_infra = pd.read_csv(infra_path)
    df_base = pd.read_csv(baseline_path)
    df_end = pd.read_csv(endline_path)

    # 2. Metric Calculation: Infrastructure & Nutrition
    # Calculating Infra Saturation based on UDISE CSV
    df_infra['Infra_Saturation'] = ((df_infra['Functional_Drinking_Water'] + 
                                     df_infra['Functional_Electricity'] + 
                                     df_infra['Functional_Toilet_Facility']) / 
                                    (df_infra['Total_Number_of_Schools'] * 3)) * 100
    
    block_infra = df_infra.groupby('Udise_Block_Name')['Infra_Saturation'].mean().reset_index()

    # 3. Data Integration: Merging Multi-Departmental Sources
    merged_data = pd.merge(df_base, df_end, on='Udise_Block_Name')
    master_df = pd.merge(merged_data, block_infra, on='Udise_Block_Name')

    # 4. Impact Analysis: Attendance & MDM Gains
    master_df['Attendance_Gain'] = master_df['Avg_Attendance_2025'] - master_df['Avg_Attendance_2023']
    master_df['MDM_Improvement'] = master_df['MDM_Quality_Score_2025'] - master_df['MDM_Quality_Score_2023']

    # 5. Visualization: Multi-Factor Impact Chart (The Image Part)
    plt.figure(figsize=(12, 7))
    blocks = master_df['Udise_Block_Name']
    x = range(len(blocks))
    width = 0.35

    plt.bar(x, master_df['Attendance_Gain'], width, label='Attendance Gain (%)', color='#2ca02c', alpha=0.8)
    plt.bar([i + width for i in x], master_df['MDM_Improvement'], width, label='MDM Quality Improvement', color='#1f77b4', alpha=0.8)
    plt.plot([i + width/2 for i in x], master_df['Infra_Saturation'], color='#d62728', marker='s', linewidth=2, label='Infra Saturation Score')

    plt.title('Impact Assessment: Infrastructure & MDM Quality vs Attendance Trends')
    plt.xlabel('Administrative Blocks')
    plt.ylabel('Performance Score / Percentage')
    plt.xticks([i + width/2 for i in x], blocks, rotation=45)
    plt.legend(loc='upper right')
    plt.grid(axis='y', linestyle='--', alpha=0.5)
    plt.tight_layout()
    
    chart_name = 'Final_Impact_Assessment_Chart.png'
    plt.savefig(chart_name)
    print(f"Chart generated: {chart_name}")

    # 6. Professional Excel Export: Traffic Light Formatting
    try:
        writer = pd.ExcelWriter(output_report, engine='xlsxwriter')
        master_df.to_excel(writer, index=False, sheet_name='Impact_Study')

        workbook = writer.book
        worksheet = writer.sheets['Impact_Study']

        # Define Formats
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        red_fmt    = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Negative
        yellow_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # 0 to 10
        green_fmt  = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # Above 10

        # Apply Header Style
        for col_num, value in enumerate(master_df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Apply Traffic Light Logic to Attendance_Gain Column
        gain_col_idx = master_df.columns.get_loc("Attendance_Gain")
        col_letter = xl_col_to_name(gain_col_idx)
        data_range = f'{col_letter}2:{col_letter}{len(master_df) + 1}'

        # Formatting Logic
        worksheet.conditional_format(data_range, {'type': 'cell', 'criteria': 'less than', 'value': 0, 'format': red_fmt})
        worksheet.conditional_format(data_range, {'type': 'cell', 'criteria': 'between', 'minimum': 0, 'maximum': 10, 'format': yellow_fmt})
        worksheet.conditional_format(data_range, {'type': 'cell', 'criteria': 'greater than', 'value': 10, 'format': green_fmt})

        writer.close()
        print(f"Excel Report generated: {output_report}")

    except PermissionError:
        print(f"Error: Please close {output_report} before running the script.")

if __name__ == "__main__":
    run_integrated_analysis()
