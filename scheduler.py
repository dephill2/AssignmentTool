import xlwings as xw
import pandas as pd
from ortools.linear_solver import pywraplp

# ---------------------------------------------------------
# HISTORICAL ASSIGNMENTS (Keep hardcoded or move to another sheet if needed later)
# ---------------------------------------------------------
historical_assignments = {
    (16, 4): {
        'ABQ': 11, 'ALN': 15, 'AMA': 5, 'ANN': 8, 'ASH': 13, 'ATG': 8, 'AUG': 2,
        'BAL': 3, 'BAY': 4, 'BEC': 11, 'BHH': 13, 'BIL': 15, 'BIR': 9, 'BOI': 7,
        'BRX': 15, 'BUF': 2, 'BYN': 6, 'CAV': 4, 'CHY': 14, 'CIN': 10, 'CLA': 7,
        'CLE': 12, 'CMO': 10, 'CMS': 6, 'CON': 12, 'CTX': 13, 'DAY': 7, 'DEN': 8,
        'DES': 3, 'DET': 10, 'DOD': 3, 'FAR': 1, 'FAV': 14, 'FHM': 13, 'FNC': 1,
        'FRE': 9, 'GLA': 0, 'GRJ': 5, 'HAM': 15, 'HOU': 5, 'HUN': 2, 'IND': 14, 'IOW': 9,
        'JAC': 7, 'KAN': 11, 'LAS': 1, 'LEA': 13, 'LEB': 14, 'LIT': 7, 'LKC': 5,
        'LOM': 4, 'MAC': 2, 'MAR': 5, 'MEM': 11, 'MIN': 6, 'MOU': 15, 'MUS': 13,
        'MWV': 6, 'NIN': 4, 'NJH': 10, 'NOL': 3, 'NOP': 12, 'NYN': 6, 'OKL': 13,
        'OMA': 6, 'ORL': 11, 'PAL': 1, 'PHO': 9, 'REN': 8, 'RIC': 2, 'SAN': 4,
        'SBY': 12, 'SHR': 0, 'SPO': 3, 'STL': 7, 'SUX': 13, 'SYR': 1, 'TOG': 10,
        'TOP': 13, 'TUC': 13, 'WAS': 6, 'WBP': 6, 'WIC': 9, 'WIM': 13, 'WRJ': 12
    }
}

def run_optimization():
    # 1. CONNECT TO EXCEL
    wb = xw.Book.caller()
    main_sheet = wb.sheets[0]  # The Dashboard/Button sheet
    data_sheet = wb.sheets['SiteData'] # The new Data sheet

    # 2. READ INPUTS
    try:
        num_rns = int(main_sheet.range('C1').value)
        aprn = int(main_sheet.range('C2').value)
    except:
        main_sheet.range('A6').value = "Error: Check inputs in C1 and C2."
        return

    # Clear previous status/results
    main_sheet.range('A6').value = "Running optimization..."
    main_sheet.range('A6').font.color = (0, 0, 0) # Black
    main_sheet.range('A8:Z1000').clear_contents()

    # 3. READ DATA DYNAMICALLY FROM 'SiteData' SHEET
    # This reads the table starting at A1 until it hits empty cells
    site_data = data_sheet.range('A1').options(pd.DataFrame, expand='table', index=False).value

    # Basic data cleaning to prevent crashes
    if site_data is None or site_data.empty:
        main_sheet.range('A6').value = "Error: No data found in 'SiteData' sheet."
        main_sheet.range('A6').font.color = (255, 0, 0)
        return
        
    # Ensure census is numeric (handle accidental text/spaces in Excel)
    site_data['TotalCensus'] = pd.to_numeric(site_data['TotalCensus'], errors='coerce').fillna(0)
    
    # 4. PREPARE DATA (Logic from your original script)
    site_data_copy = site_data.copy()
    site_data_copy['OriginalSiteCode'] = site_data_copy['SiteCode']
    # Handle the specific SPO/DOD grouping logic
    site_data_copy['SiteCode'] = site_data_copy['SiteCode'].apply(lambda x: 'SPO_DOD' if x in ['SPO', 'DOD'] else x)
    
    # Group sites if needed (e.g. SPO and DOD become one entry)
    site_data_grouped = site_data_copy.groupby('SiteCode').agg({
        'TotalCensus': 'sum',
        'facilityName': lambda x: list(x),
        'OriginalSiteCode': lambda x: list(x)
    }).reset_index()

    site_data_to_use = site_data_grouped.copy()
    non_aprn_sites = site_data_to_use.reset_index(drop=True)

    # 5. RUN OR-TOOLS OPTIMIZATION
    solver = pywraplp.Solver.CreateSolver('SCIP')
    if not solver:
        main_sheet.range('A6').value = "Error: SCIP solver not available."
        return

    NUM_RNS = num_rns
    MAX_PATIENTS_PER_RN = 55

    assign = {}
    penalty = {}
    
    # Logic to fetch historical assignment based on inputs
    # If exact key (rns, docs) doesn't exist, fallback to default or empty
    base_key = (16, 4) 
    base_assignment = historical_assignments.get(base_key, {})

    # Create Variables
    for i in range(len(non_aprn_sites)):
        for rn in range(NUM_RNS):
            assign[(i, rn)] = solver.IntVar(0, 1, f'site_{i}_to_rn_{rn}')
            # Simplified Penalty Logic for migration
            penalty[(i, rn)] = solver.IntVar(0, 1, f'penalty_{i}_{rn}')
            solver.Add(penalty[(i, rn)] >= assign[(i, rn)])

    # Constraint: Every site must be assigned to exactly one RN
    for i in range(len(non_aprn_sites)):
        solver.Add(sum(assign[(i, n)] for n in range(NUM_RNS)) == 1)

    # Constraint: Max Patients and Max Sites
    for n in range(NUM_RNS):
        total_patients = sum(assign[(i, n)] * non_aprn_sites.iloc[i]['TotalCensus'] for i in range(len(non_aprn_sites)))
        solver.Add(total_patients <= MAX_PATIENTS_PER_RN)
        
        total_sites = sum(assign[(i, n)] for i in range(len(non_aprn_sites)))
        solver.Add(total_sites <= 7)

    # Constraint: Balance Patients (Minimize variance)
    avg_patients = sum(non_aprn_sites['TotalCensus']) / NUM_RNS
    patient_diff = {}
    for n in range(NUM_RNS):
        patient_diff[n] = solver.NumVar(0, solver.infinity(), f'patient_diff_{n}')
        total_patients = sum(assign[(i, n)] * non_aprn_sites.iloc[i]['TotalCensus'] for i in range(len(non_aprn_sites)))
        solver.Add(patient_diff[n] >= total_patients - avg_patients)
        solver.Add(patient_diff[n] >= avg_patients - total_patients)

    # Constraint: Specific Site Groupings (Fayetteville, Columbia, etc)
    fayetteville_sites = ['FAV', 'FNC']
    columbia_sites = ['CMO', 'CMS']
    
    for rn in range(NUM_RNS):
        fayetteville_vars = [assign[(i, rn)] for i in range(len(non_aprn_sites)) 
                             if any(code in fayetteville_sites for code in non_aprn_sites.iloc[i]['OriginalSiteCode'])]
        solver.Add(sum(fayetteville_vars) <= 1)

    # Objective: Minimize patient difference
    solver.Minimize(sum(patient_diff[n] for n in range(NUM_RNS)))

    # 6. SOLVE
    status = solver.Solve()

    # 7. OUTPUT RESULTS
    if status == pywraplp.Solver.OPTIMAL:
        main_sheet.range('A6').value = "Optimization Successful!"
        main_sheet.range('A6').font.color = (0, 128, 0) # Green
        
        output_rows = []
        for n in range(NUM_RNS):
            assigned_sites = []
            assigned_census = 0
            
            for i in range(len(non_aprn_sites)):
                if assign[(i, n)].solution_value() > 0.5:
                    site_names = non_aprn_sites.iloc[i]['facilityName']
                    census = non_aprn_sites.iloc[i]['TotalCensus']
                    
                    # Convert list to string if it was grouped
                    if isinstance(site_names, list):
                        site_names = " & ".join(site_names)
                    
                    assigned_sites.append(f"{site_names} ({int(census)})")
                    assigned_census += census
            
            # Format: RN Name | Total Census | Sites
            output_rows.append([f"RN {n+1}", int(assigned_census), ", ".join(assigned_sites)])
        
        # Write Headers
        main_sheet.range('A8').value = ["RN Name", "Total Census", "Assigned Sites"]
        main_sheet.range('A8:C8').font.bold = True
        main_sheet.range('A8:C8').color = (217, 225, 242) # Light blue background
        
        # Write Data
        if output_rows:
            main_sheet.range('A9').value = output_rows
            
    else:
        main_sheet.range('A6').value = "No optimal solution found. Try increasing RN count."
        main_sheet.range('A6').font.color = (255, 0, 0)

if __name__ == "__main__":
    xw.Book("Scheduler.xlsm").set_mock_caller()
    run_optimization()