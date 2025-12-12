import xlwings as xw
import pandas as pd
from ortools.linear_solver import pywraplp

# ---------------------------------------------------------
# DATA & MAPPINGS (Provider Groups & Historicals)
# ---------------------------------------------------------
site_groups = {
    2: {
        1: ['Ann Arbor', 'Asheville', 'Atlanta', 'Augusta', 'Baltimore', 'Bay Pines', 'Biloxi', 'Black Hills', 'Beckley', 'Birmingham', 'Bronx', 'Central Alabama', 'Cincinnati', 'Clarksburg', 'Cleveland', 'Columbia, SC', 'Connecticut', 'Dayton', 'Detroit', 'Fayetteville, NC', 'Huntington', 'Hampton', 'Indianapolis', 'Lebanon', 'Martinsburg', 'Memphis', 'Mountain Home', 'North Florida/South Georgia', 'Northern Indiana', 'New Jersey', 'Northport', 'NYN-New York Harbor', 'BYN-New York Harbor', 'Orlando', 'Richmond', 'Salisbury', 'San Juan', 'Togus', 'ALN-Upstate New York', 'BUF-Upstate New York', 'SYR-Upstate New York', 'Washington DC', 'White River Junction', 'Wilkes-Barre', 'Wilmington'],
        2: ['Albuquerque', 'Amarillo', 'Boise', 'Cheyenne', 'Columbia, MO', 'Denver', 'Des Moines', 'Fargo', 'Fayetteville, AR', 'Fresno', 'Fort Harrison', 'Grand Junction', 'Houston', 'Iowa City', 'Jackson', 'Kansas City', 'Las Vegas', 'Leavenworth', 'Little Rock', 'Loma Linda', 'Marion', 'Minneapolis', 'Muskogee', 'Nellis AFB DOD', 'New Orleans', 'Oklahoma City', 'Omaha', 'Palo Alto', 'Phoenix', 'Reno', 'Sacramento', 'Shreveport', 'Sioux Falls', 'Spokane', 'St. Louis', 'Temple (Central Texas)', 'Topeka', 'Tucson', 'West Los Angeles', 'Wichita']
    },
    3: {
        1: ['Ann Arbor', 'Asheville', 'Augusta', 'Bay Pines', 'Beckley', 'Bronx', 'Cincinnati', 'Clarksburg', 'Cleveland', 'Columbia, SC', 'Connecticut', 'Dayton', 'Detroit', 'Fayetteville, NC', 'Huntington','Hampton', 'Lebanon', 'Memphis', 'Northern Indiana', 'Northport', 'NYN-New York Harbor','BYN-New York Harbor', 'Salisbury', 'ALN-Upstate New York', 'BUF-Upstate New York', 'SYR-Upstate New York', 'White River Junction', 'Wilkes-Barre', 'Wilmington'],
        2: ['Albuquerque', 'Biloxi', 'Fort Harrison', 'Grand Junction', 'Houston', 'Iowa City', 'Jackson', 'Kansas City', 'Las Vegas', 'Leavenworth', 'Little Rock', 'Marion', 'Muskogee', 'Nellis AFB DOD', 'Oklahoma City', 'Palo Alto', 'Reno', 'San Juan', 'Sacramento', 'Shreveport', 'Sioux Falls', 'Spokane', 'St. Louis', 'Temple (Central Texas)', 'Topeka', 'Tucson', 'West Los Angeles', 'Wichita'],
        3: ['Atlanta', 'Baltimore', 'Birmingham', 'Central Alabama', 'Indianapolis', 'Martinsburg', 'Mountain Home', 'North Florida/South Georgia', 'New Jersey', 'Orlando', 'Richmond', 'Togus', 'Washington DC', 'Amarillo', 'Black Hills', 'Boise', 'Cheyenne', 'Columbia, MO', 'Denver', 'Des Moines', 'Fargo', 'Fayetteville, AR', 'Fresno', 'Loma Linda', 'Minneapolis', 'New Orleans', 'Omaha', 'Phoenix']
    },
    4: {
        1: ['Ann Arbor', 'Asheville', 'Augusta', 'Bay Pines', 'Beckley', 'Bronx', 'Cincinnati', 'Clarksburg', 'Columbia, SC', 'Connecticut', 'Dayton', 'Detroit', 'Huntington','Hampton', 'Lebanon', 'Northern Indiana', 'NYN-New York Harbor','BYN-New York Harbor', 'Salisbury', 'ALN-Upstate New York', 'BUF-Upstate New York', 'SYR-Upstate New York','White River Junction', 'Wilmington', 'Wilkes-Barre'],
        2: ['Albuquerque', 'Amarillo', 'Cheyenne', 'Denver', 'Grand Junction', 'Houston', 'Iowa City', 'Marion', 'Muskogee', 'Nellis AFB DOD', 'Oklahoma City', 'Palo Alto', 'San Juan', 'Reno', 'Spokane', 'St. Louis', 'Topeka', 'Tucson', 'Wichita'],
        3: ['Atlanta', 'Baltimore', 'Birmingham', 'Central Alabama', 'Fayetteville, NC', 'Indianapolis', 'Martinsburg', 'Memphis', 'Mountain Home', 'North Florida/South Georgia', 'New Jersey', 'Northport', 'Orlando', 'Richmond', 'Togus', 'Washington DC', 'Black Hills', 'Loma Linda'],
        4: ['Biloxi', 'Boise', 'Columbia, MO', 'Des Moines', 'Fargo', 'Fayetteville, AR', 'Fresno', 'Fort Harrison', 'Jackson', 'Kansas City', 'Las Vegas', 'Leavenworth', 'Little Rock', 'Minneapolis', 'New Orleans', 'Omaha', 'Phoenix', 'Sacramento', 'Shreveport', 'Sioux Falls', 'Temple (Central Texas)', 'West Los Angeles', 'Cleveland']
    },
    5: {
        1: ['Ann Arbor', 'Augusta', 'Asheville', 'Bay Pines', 'Bronx', 'Clarksburg', 'Connecticut', 'Detroit', 'Huntington','Hampton', 'Northern Indiana', 'Lebanon', 'NYN-New York Harbor','BYN-New York Harbor', 'ALN-Upstate New York', 'BUF-Upstate New York', 'SYR-Upstate New York', 'White River Junction', 'Wilmington'],
        2: ['Albuquerque', 'Amarillo', 'Cheyenne', 'Denver', 'Grand Junction', 'Houston', 'Muskogee', 'Nellis AFB DOD', 'Oklahoma City', 'Palo Alto', 'Reno', 'Spokane', 'St. Louis', 'Topeka', 'Tucson', 'Wichita'],
        3: ['Atlanta', 'Baltimore', 'Birmingham', 'Central Alabama', 'Fayetteville, NC', 'Indianapolis', 'Martinsburg', 'Mountain Home', 'North Florida/South Georgia', 'New Jersey', 'Richmond', 'Togus', 'Washington DC', 'Loma Linda'],
        4: ['Biloxi', 'Boise', 'Columbia, MO', 'Des Moines', 'Fargo', 'Fayetteville, AR', 'Fresno', 'Fort Harrison', 'Jackson', 'Kansas City', 'Little Rock', 'New Orleans', 'Omaha', 'Phoenix', 'Sacramento', 'Shreveport', 'Sioux Falls', 'Temple (Central Texas)', 'Cleveland', 'West Los Angeles'],
        5: ['Beckley', 'Cincinnati', 'Columbia, SC', 'Dayton', 'San Juan', 'Minneapolis', 'Memphis', 'Northern Indiana', 'Northport', 'Orlando', 'Salisbury', 'Wilkes-Barre', 'Grand Junction', 'Iowa City', 'Las Vegas', 'Leavenworth', 'Marion']
    },
    6: {
        1: ['Asheville', 'Bay Pines', 'Bronx', 'Clarksburg', 'Connecticut', 'Detroit', 'Huntington', 'Hampton', 'Lebanon', 'NYN-New York Harbor','BYN-New York Harbor', 'ALN-Upstate New York', 'BUF-Upstate New York', 'SYR-Upstate New York', 'White River Junction', 'Wilmington', 'Augusta'],
        2: ['Albuquerque', 'Amarillo', 'Cheyenne', 'Denver', 'Nellis AFB DOD', 'Oklahoma City', 'Palo Alto', 'Reno', 'Spokane', 'St. Louis', 'Topeka', 'Tucson', 'Wichita'],
        3: ['Baltimore', 'Atlanta', 'Birmingham', 'Central Alabama', 'Fayetteville, NC', 'Martinsburg', 'Mountain Home', 'North Florida/South Georgia', 'New Jersey', 'Richmond', 'Togus', 'Washington DC'],
        4: ['Biloxi', 'Boise', 'Columbia, MO', 'Des Moines', 'Fargo', 'Fresno', 'Jackson', 'Kansas City', 'Little Rock', 'New Orleans', 'Omaha', 'Phoenix', 'Sacramento', 'Shreveport', 'Sioux Falls', 'Temple (Central Texas)', 'Cleveland'],
        5: ['Beckley', 'Cincinnati', 'Columbia, SC', 'Dayton', 'Minneapolis', 'Memphis', 'Northern Indiana', 'Northport', 'Salisbury', 'Wilkes-Barre', 'Grand Junction', 'Iowa City', 'Las Vegas', 'Leavenworth', 'Marion'],
        6: ['Ann Arbor', 'Orlando', 'San Juan', 'Fayetteville, AR', 'Black Hills', 'Fort Harrison', 'Houston', 'Indianapolis', 'Muskogee', 'West Los Angeles', 'Loma Linda']
    },
    7: {
        1: ['Asheville', 'Bay Pines', 'Bronx', 'Clarksburg', 'Connecticut', 'Detroit', 'Huntington', 'Hampton', 'Lebanon', 'ALN-Upstate New York', 'BUF-Upstate New York', 'SYR-Upstate New York', 'White River Junction', 'Wilmington', 'Augusta'],
        2: ['Albuquerque', 'Amarillo', 'Cheyenne', 'Nellis AFB DOD', 'Oklahoma City', 'Palo Alto', 'Reno', 'Spokane', 'St. Louis', 'Topeka', 'Tucson'],
        3: ['Atlanta', 'Birmingham', 'Central Alabama', 'Cleveland', 'Mountain Home', 'North Florida/South Georgia', 'New Jersey', 'Richmond', 'Washington DC', 'Fayetteville, NC'],
        4: ['Biloxi', 'Boise', 'Columbia, MO', 'Des Moines', 'Kansas City', 'Little Rock', 'New Orleans', 'Omaha', 'Phoenix', 'Sacramento', 'Shreveport', 'Sioux Falls', 'Temple (Central Texas)'],
        5: ['Beckley', 'Cincinnati', 'Columbia, SC', 'Dayton', 'Minneapolis', 'Memphis', 'Northern Indiana', 'Northport', 'Salisbury', 'Wilkes-Barre', 'Grand Junction', 'Las Vegas', 'Leavenworth', 'Marion', 'West Los Angeles'],
        6: ['Ann Arbor', 'San Juan', 'Fayetteville, AR', 'Black Hills', 'Fort Harrison', 'Indianapolis', 'Togus', 'Orlando', 'Loma Linda'],
        7: ['NYN-New York Harbor', 'BYN-New York Harbor','Martinsburg', 'Iowa City', 'Denver', 'Fargo', 'Fresno', 'Houston', 'Jackson', 'Wichita', 'Muskogee', 'Baltimore']
    },
    8: {
        1: ['Asheville', 'Bay Pines', 'Bronx', 'Clarksburg', 'Connecticut', 'Detroit', 'Huntington', 'Hampton','Lebanon', 'ALN-Upstate New York','BUF-Upstate New York', 'SYR-Upstate New York',],
        2: ['Albuquerque', 'Amarillo', 'Cheyenne', 'Nellis AFB DOD', 'Oklahoma City', 'Palo Alto', 'Reno', 'Spokane', 'St. Louis'],
        3: ['Atlanta', 'Birmingham', 'Central Alabama', 'Cleveland', 'Mountain Home', 'North Florida/South Georgia', 'New Jersey', 'Richmond'],
        4: ['Biloxi', 'Boise', 'Columbia, MO', 'Des Moines', 'Kansas City', 'Little Rock', 'New Orleans', 'Omaha', 'Phoenix', 'Sacramento'],
        5: ['Beckley', 'Cincinnati', 'Columbia, SC', 'Dayton', 'Minneapolis', 'Memphis', 'Northern Indiana', 'Northport', 'Salisbury', 'Wilkes-Barre', 'Grand Junction', 'Las Vegas', 'Leavenworth', 'Marion'],
        6: ['Ann Arbor', 'San Juan', 'Black Hills', 'Fort Harrison', 'Indianapolis', 'Togus', 'Orlando', 'Loma Linda', 'Fayetteville, AR'],
        7: ['NYN-New York Harbor','BYN-New York Harbor', 'Martinsburg', 'Iowa City', 'Denver', 'Fargo', 'Houston', 'Muskogee', 'Baltimore', 'Fresno', 'Wichita'],
        8: ['White River Junction', 'Wilmington', 'Topeka', 'Fayetteville, NC', 'Sioux Falls', 'Temple (Central Texas)', 'Tucson', 'Shreveport', 'Jackson', 'West Los Angeles', 'Washington DC', 'Augusta']
    }
}

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

aprn_sites = ['ASH', 'BEC', 'BHH', 'CAV', 'CHY', 'FAR', 'FAV', 'FHM', 'LEA', 'LEB', 'LKC', 'MAR', 'MUS', 'SUX', 'TOP', 'WIM', 'BIL', 'CLA', 'CTX', 'GRJ', 'MOU', 'WIC', 'BRX', 'CMS', 'DET', 'MWV', 'TOG', 'SBY']

def run_optimization():
    # 1. SETUP & READ
    wb = xw.Book.caller()
    main_sheet = wb.sheets[0]
    data_sheet = wb.sheets['SiteData']
    
    # Inputs
    try:
        num_rns = int(main_sheet.range('C1').value)
        aprn = int(main_sheet.range('C2').value)
        num_providers = int(main_sheet.range('C3').value)
    except:
        main_sheet.range('A6').value = "Error: Check inputs (C1, C2, C3)."
        return

    main_sheet.range('A6').value = "Running optimization..."
    main_sheet.range('A6').font.color = (0, 0, 0)
    # Clear previous output
    main_sheet.range('A8:M500').clear_contents()
    main_sheet.range('A8:M500').color = None
    main_sheet.range('A8:M500').font.bold = False

    # Read Data
    site_data = data_sheet.range('A1').options(pd.DataFrame, expand='table', index=False).value
    if site_data is None or site_data.empty:
        main_sheet.range('A6').value = "Error: SiteData empty."
        return
    site_data['TotalCensus'] = pd.to_numeric(site_data['TotalCensus'], errors='coerce').fillna(0)
    
    # *** KEY FIX: Add 'OriginalSiteCode' to main dataframe immediately ***
    site_data['OriginalSiteCode'] = site_data['SiteCode']

    # ---------------------------------------------------------
    # PART 1: RN OPTIMIZATION (OR-TOOLS)
    # ---------------------------------------------------------
    site_data_copy = site_data.copy()
    # Logic for grouping SPO/DOD
    site_data_copy['SiteCode'] = site_data_copy['SiteCode'].apply(lambda x: 'SPO_DOD' if x in ['SPO', 'DOD'] else x)
    
    site_data_grouped = site_data_copy.groupby('SiteCode').agg({
        'TotalCensus': 'sum',
        'facilityName': lambda x: list(x),
        'OriginalSiteCode': lambda x: list(x)
    }).reset_index()

    site_data_to_use = site_data_grouped.copy()
    non_aprn_sites = site_data_to_use.reset_index(drop=True)

    solver = pywraplp.Solver.CreateSolver('SCIP')
    NUM_RNS = num_rns
    MAX_PATIENTS_PER_RN = 55
    assign = {}
    
    for i in range(len(non_aprn_sites)):
        for rn in range(NUM_RNS):
            assign[(i, rn)] = solver.IntVar(0, 1, f'site_{i}_to_rn_{rn}')
    
    for i in range(len(non_aprn_sites)):
        solver.Add(sum(assign[(i, n)] for n in range(NUM_RNS)) == 1)

    for n in range(NUM_RNS):
        total_patients = sum(assign[(i, n)] * non_aprn_sites.iloc[i]['TotalCensus'] for i in range(len(non_aprn_sites)))
        solver.Add(total_patients <= MAX_PATIENTS_PER_RN)
        solver.Add(sum(assign[(i, n)] for i in range(len(non_aprn_sites))) <= 7)

    avg_patients = sum(non_aprn_sites['TotalCensus']) / NUM_RNS
    patient_diff = {}
    for n in range(NUM_RNS):
        patient_diff[n] = solver.NumVar(0, solver.infinity(), f'patient_diff_{n}')
        total_patients = sum(assign[(i, n)] * non_aprn_sites.iloc[i]['TotalCensus'] for i in range(len(non_aprn_sites)))
        solver.Add(patient_diff[n] >= total_patients - avg_patients)
        solver.Add(patient_diff[n] >= avg_patients - total_patients)
    
    fayetteville_sites = ['FAV', 'FNC']
    for rn in range(NUM_RNS):
        fayetteville_vars = [assign[(i, rn)] for i in range(len(non_aprn_sites)) 
                             if any(code in fayetteville_sites for code in non_aprn_sites.iloc[i]['OriginalSiteCode'])]
        solver.Add(sum(fayetteville_vars) <= 1)

    solver.Minimize(sum(patient_diff[n] for n in range(NUM_RNS)))
    status = solver.Solve()

    if status != pywraplp.Solver.OPTIMAL:
        main_sheet.range('A6').value = "No optimal solution found."
        main_sheet.range('A6').font.color = (255, 0, 0)
        return

    # Store RN Assignments for Cross-Walk
    rn_lookup = {} # Maps Site Name -> RN Number
    
    # ---------------------------------------------------------
    # PART 2: WRITE RN OUTPUT (Column A)
    # ---------------------------------------------------------
    main_sheet.range('A8').value = "RN Distribution"
    main_sheet.range('A8').font.bold = True
    main_sheet.range('A8').font.size = 14
    
    row = 10
    for n in range(NUM_RNS):
        main_sheet.range(f'A{row}').value = f"RN {n+1}"
        main_sheet.range(f'A{row}').font.bold = True
        main_sheet.range(f'A{row}').color = (217, 225, 242) # Light Blue
        row += 1
        
        assigned_census = 0
        current_site_start_row = row
        
        for i in range(len(non_aprn_sites)):
            if assign[(i, n)].solution_value() > 0.5:
                # Handle Grouped Sites
                names = non_aprn_sites.iloc[i]['facilityName']
                census = non_aprn_sites.iloc[i]['TotalCensus']
                
                # If grouped (list), iterate
                if not isinstance(names, list):
                    names = [names]
                
                for site_name in names:
                    # Find individual census if grouped
                    site_row = site_data[site_data['facilityName'] == site_name]
                    
                    # Safe access to TotalCensus
                    ind_census = int(site_row['TotalCensus'].iloc[0]) if not site_row.empty else 0
                    
                    # Safe access to OriginalSiteCode
                    orig_code = site_row['OriginalSiteCode'].iloc[0] if not site_row.empty else ""
                    
                    # Write Site
                    main_sheet.range(f'B{row}').value = site_name
                    main_sheet.range(f'C{row}').value = ind_census
                    
                    # Blue if APRN
                    if aprn == 1 and orig_code in aprn_sites:
                         main_sheet.range(f'B{row}').font.color = (0, 0, 255)
                    
                    # Store for Cross-Walk
                    rn_lookup[site_name] = f"RN {n+1}"
                    row += 1
                
                assigned_census += census

        # Add Total
        main_sheet.range(f'B{row}').value = f"Total: {int(assigned_census)}"
        main_sheet.range(f'B{row}').font.bold = True
        row += 2

    # ---------------------------------------------------------
    # PART 3: PROVIDER TEAMS (Column E)
    # ---------------------------------------------------------
    main_sheet.range('E8').value = "Provider Team Detail"
    main_sheet.range('E8').font.bold = True
    main_sheet.range('E8').font.size = 14
    
    if num_providers in site_groups:
        prov_teams = site_groups[num_providers]
        row = 10
        provider_lookup = {} # Maps Site Name -> Provider Team
        
        for team_num, sites_in_team in prov_teams.items():
            main_sheet.range(f'E{row}').value = f"Team {team_num}"
            main_sheet.range(f'E{row}').font.bold = True
            main_sheet.range(f'E{row}').color = (217, 225, 242)
            row += 1
            
            total_prov_census = 0
            # Sort by category for display
            grouped_by_cat = {}
            
            for site_name in sites_in_team:
                site_row = site_data[site_data['facilityName'] == site_name]
                if not site_row.empty:
                    c = int(site_row['TotalCensus'].iloc[0])
                    cat = int(site_row['Categories'].iloc[0])
                    grouped_by_cat.setdefault(cat, []).append((site_name, c))
                    total_prov_census += c
                    provider_lookup[site_name] = f"Team {team_num}"

            for cat in sorted(grouped_by_cat.keys()):
                main_sheet.range(f'F{row}').value = f"Category {cat}"
                # Red for cat 1/2
                if cat <= 2:
                    main_sheet.range(f'F{row}').font.color = (255, 0, 0) 
                    main_sheet.range(f'F{row}').color = (252, 228, 214) # Light red bg
                else:
                    main_sheet.range(f'F{row}').color = (165, 165, 165) # Grey bg
                    main_sheet.range(f'F{row}').font.color = (255, 255, 255)

                row += 1
                for s_name, c_count in sorted(grouped_by_cat[cat]):
                    main_sheet.range(f'F{row}').value = s_name
                    main_sheet.range(f'G{row}').value = c_count
                    row += 1
            
            main_sheet.range(f'F{row}').value = f"Est. Total: {total_prov_census}"
            main_sheet.range(f'F{row}').font.bold = True
            row += 2
    else:
        main_sheet.range('E10').value = "Provider grouping not defined for this count."

    # ---------------------------------------------------------
    # PART 4: CROSS-WALK TABLE (Column I)
    # ---------------------------------------------------------
    main_sheet.range('I8').value = "Cross-Walk Reference"
    main_sheet.range('I8').font.bold = True
    main_sheet.range('I8').font.size = 14
    
    # Headers
    headers = ["Site Name", "Assigned RN", "Assigned Provider"]
    main_sheet.range('I10').value = headers
    main_sheet.range('I10:K10').font.bold = True
    main_sheet.range('I10:K10').color = (217, 225, 242)
    
    # Build Table Data
    cross_walk_data = []
    all_sites = site_data['facilityName'].tolist()
    all_sites.sort()
    
    for s in all_sites:
        rn = rn_lookup.get(s, "Unassigned")
        prov = provider_lookup.get(s, "Unassigned")
        cross_walk_data.append([s, rn, prov])
        
    main_sheet.range('I11').value = cross_walk_data
    
    # ---------------------------------------------------------
    # FINISH
    # ---------------------------------------------------------
    main_sheet.range('A6').value = "Optimization Successful!"
    main_sheet.range('A6').font.color = (0, 128, 0)
    main_sheet.autofit()

if __name__ == "__main__":
    xw.Book("Scheduler.xlsm").set_mock_caller()
    run_optimization()