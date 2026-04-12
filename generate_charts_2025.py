import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = 'white'
plt.rcParams['font.size'] = 9

def merge_headers(df):
    """Combine row 0 and row 1 into column names."""
    row1 = df.iloc[0].fillna('').astype(str)
    row2 = df.iloc[1].fillna('').astype(str)

    column_names = []
    for i in range(len(row1)):
        if row2[i] and row2[i] != 'nan' and row2[i].strip() != '':
            if row1[i] and row1[i] != 'nan':
                combined = f"{row1[i]} | {row2[i]}"
            else:
                combined = row2[i]
        else:
            combined = row1[i] if row1[i] and row1[i] != 'nan' else f"Column_{i}"
        column_names.append(combined)

    df.columns = column_names
    return df.iloc[2:].reset_index(drop=True)

def create_pie_chart(data, title, filename, output_dir='charts'):
    """Pie chart for single-choice questions."""
    value_counts = data.value_counts()
    if len(value_counts) == 0:
        return

    labels = [label[:40] + '...' if len(str(label)) > 40 else str(label) for label in value_counts.index]

    fig, ax = plt.subplots(figsize=(10, 8))
    colors = plt.cm.Set3(np.linspace(0, 1, len(value_counts)))

    wedges, texts, autotexts = ax.pie(value_counts.values, labels=labels, autopct='%1.1f%%',
                                       startangle=90, colors=colors)

    for text in texts:
        text.set_fontsize(9)
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(8)

    ax.set_title(title[:70], fontsize=11, fontweight='bold', pad=20)
    plt.tight_layout()
    plt.savefig(f'{output_dir}/{filename}', dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Created: {output_dir}/{filename}")

def create_bar_chart(cols, title, filename, labels, df, output_dir='charts'):
    """Horizontal bar chart for multi-choice questions."""
    counts = [df[col].notna().sum() for col in cols]

    if sum(counts) == 0:
        return

    fig, ax = plt.subplots(figsize=(14, 8))
    colors = plt.cm.Spectral(np.linspace(0.1, 0.9, len(counts)))
    bars = ax.barh(range(len(counts)), counts, color=colors)

    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlabel('Number of Responses', fontsize=10)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=15)

    total = len(df)
    for i, (bar, count) in enumerate(zip(bars, counts)):
        ax.text(count + 1, bar.get_y() + bar.get_height()/2,
                f'{count} ({count/total*100:.1f}%)',
                ha='left', va='center', fontsize=9)

    ax.invert_yaxis()
    plt.tight_layout()
    plt.savefig(f'{output_dir}/{filename}', dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Created: {output_dir}/{filename}")

def main():
    input_file = 'Dempsey Center 2025 Annual Client Survey.xlsx'
    output_dir = 'charts'

    os.makedirs(output_dir, exist_ok=True)

    df_raw = pd.read_excel(input_file, header=None)
    df = merge_headers(df_raw)
    df.to_excel('processed_data.xlsx', index=False)

    # Single-choice charts
    single_choice_questions = [
        ('1. Overall, please rate your satisfaction with the services you received from the Dempsey Center in 2025. | Response',
         'Q1: Overall Satisfaction', 'Q1_Satisfaction_Pie.png'),
        ('2. Overall, did/does the Dempsey Center make your life better as it relates to your cancer impact? | Response',
         'Q2: Does Dempsey Center Make Your Life Better?', 'Q2_Make_Life_Better_Pie.png'),
        ('4. How often did you utilize services in 2025? | Response',
         'Q4: Service Utilization Frequency', 'Q4_Utilize_Frequency_Pie.png'),
        ("13. At what stage was your or your loved one's cancer when you/they last recieved or were offered treatment? | Response",
         'Q13: Cancer Stage', 'Q13_Cancer_Stage_Pie.png'),
        ('14. What is your age? | Response',
         'Q14: Age Distribution', 'Q14_Age_Pie.png'),
        ('16. What was your total household income before taxes in 2025? | Response',
         'Q16: Household Income', 'Q16_Income_Pie.png'),
        ('18. How would you describe your current gender identity? | Response',
         'Q18: Gender Identity', 'Q18_Gender_Pie.png'),
        ('19. What is your ethnicity? | Response',
         'Q19: Ethnicity', 'Q19_Ethnicity_Pie.png'),
        ('21. How would you describe your current sexual orientation/identity? | Response',
         'Q21: Sexual Orientation', 'Q21_Orientation_Pie.png'),
        ('23. Which of the following best describes the type of health insurance coverage you currently have (select one)? | Response',
         'Q23: Health Insurance', 'Q23_Insurance_Pie.png'),
        ('24. What is the last year of schooling that you have completed? | Response',
         'Q24: Education Level', 'Q24_Education_Pie.png'),
    ]

    for col, title, filename in single_choice_questions:
        if col in df.columns:
            create_pie_chart(df[col].dropna(), title, filename, output_dir)

    # Multi-choice charts
    q3_cols = [c for c in df.columns if c.startswith('In what ways did the Dempsey Center make your life better')]
    q3_labels = [
        'Reduced physical side effects', 'Reduced emotional distress',
        'Reduced isolation; created community', 'Improved relationships',
        'Helped adhere to treatment plan', 'Regained ownership of choices',
        'Healthier lifestyle', 'Rebuilt sense of self/identity', 'None of the above'
    ]
    create_bar_chart(q3_cols, 'Q3: How Did Dempsey Center Make Your Life Better?', 'Q3_Life_Better_Bar.png', q3_labels, df, output_dir)

    q5_cols = [c for c in df.columns if 'negatively affect your ability to access' in c]
    q5_labels = [
        'None of the above', 'Long wait to become client', 'Wait time for services',
        'Service times not suitable', 'Physical location', 'Transportation',
        'Internet/hardware issues', 'Did not feel welcomed', 'Other'
    ]
    create_bar_chart(q5_cols, 'Q5: Barriers to Accessing Services', 'Q5_Barriers_Bar.png', q5_labels, df, output_dir)

    q10_cols = [c for c in df.columns if 'Where did you receive services' in c]
    q10_labels = [
        'In-Person Lewiston', 'In-Person South Portland', 'In-Person Westbrook (Rock Row)',
        'Virtual Live (Dempsey Connects)', 'Virtual On-Demand', "Clayton's House"
    ]
    create_bar_chart(q10_cols, 'Q10: Where Did You Receive Services?', 'Q10_Service_Locations_Bar.png', q10_labels, df, output_dir)

    q20_cols = [c for c in df.columns if 'What racial identity best describes you' in c]
    q20_labels = [
        'American Indian/Alaskan Native', 'Asian', 'Black/African American',
        'Middle Eastern/Northern African', 'Multiracial', 'Native Hawaiian/Pacific Islander',
        'White', 'Prefer to self-describe'
    ]
    create_bar_chart(q20_cols, 'Q20: Racial Identity', 'Q20_Race_Bar.png', q20_labels, df, output_dir)

    q22_cols = [c for c in df.columns if 'Do you live with any of the following disabilities' in c]
    q22_labels = [
        'No known disabilities', 'Hearing Difficulty', 'Vision Difficulty',
        'Cognitive Difficulty', 'Ambulatory Difficulty', 'Self-Care Difficulty',
        'Independent Living Difficulty', 'Prefer not to answer'
    ]
    create_bar_chart(q22_cols, 'Q22: Disabilities', 'Q22_Disabilities_Bar.png', q22_labels, df, output_dir)

    print("\nAll charts generated successfully.")

if __name__ == '__main__':
    main()
