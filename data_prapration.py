import pandas as pd
import numpy as np
import os

# This line finds the folder where this script is saved
base_path = os.path.dirname(os.path.abspath(__file__))

# UPDATE: Match the exact names from your os.listdir()
person_file = os.path.join(base_path, 'color_score.csv')
job_file = os.path.join(base_path, 'job_score.xlsx') # Use the .xlsx file

def run_analysis():
    # Check if files exist
    if not os.path.exists(person_file):
        print(f"ERROR: Cannot find {person_file}")
        return
    if not os.path.exists(job_file):
        print(f"ERROR: Cannot find {job_file}")
        return

    print("Files found! Loading data...")
    
    # Load Person Data (CSV)
    p_df = pd.read_csv(person_file)
    
    # Load Job Data (Excel)
    # Note: Using sheet_name='Sheet1' and header=1 to match your data structure
    try:
        j_df = pd.read_excel(job_file, sheet_name='Sheet1', header=1)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        print("Try: pip install openpyxl")
        return

    print("Data loaded successfully!")

    # --- Processing Logic ---
    p_cols = ['RED Score', 'BLUE Score', 'WHITE Score', 'YELLOW Score']
    j_cols = ['Red_Score', 'Blue_Score', 'White_Score', 'Yellow_Score']

    # Normalize Function
    def norm(df, cols):
        d = df[cols].apply(pd.to_numeric, errors='coerce').dropna()
        return (d - d.min()) / (d.max() - d.min())

    # Process Person Scores
    p_norm = norm(p_df, p_cols)
    p_norm['Color'] = p_df.loc[p_norm.index, 'Final color']
    avg_p = p_norm.groupby('Color')[p_cols].mean()

    # Process Job Scores
    j_norm = norm(j_df, j_cols)
    j_norm.columns = p_cols # Standardize names to match Person scores
    j_titles = j_df.loc[j_norm.index, 'Row Labels']

    # Stress/Burnout Calculation
    analysis = []
    for idx, job_row in j_norm.iterrows():
        title = j_titles[idx]
        for color, profile in avg_p.iterrows():
            diff = job_row - profile
            
            # Identify Stress Factors (Where Job demand > Person's natural score)
            stress_factors = diff[diff > 0.2]
            
            # Calculate Burnout Risk (0 to 1+)
            risk_score = np.sqrt(np.sum(np.square(stress_factors))) if not stress_factors.empty else 0
            
            analysis.append({
                'Job': title,
                'Your Color': color,
                'Burnout Risk Score': round(risk_score, 3),
                'Burnout Reason': " / ".join([f"High {k.split(' ')[0]}" for k in stress_factors.index]) if not stress_factors.empty else "Natural Fit"
            })

    # Save output
    output_path = os.path.join(base_path, 'burnout_risk_report.csv')
    pd.DataFrame(analysis).to_csv(output_path, index=False)
    print(f"\nSUCCESS! Report created: {output_path}")
    print("You can now open 'burnout_risk_report.csv' to see the burnout scores.")

if __name__ == "__main__":
    run_analysis()