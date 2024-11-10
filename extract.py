import re
import pandas as pd
import json

# Load the entire Excel file
file_path = r"Copy of Data_for_Time_Table_software_-27_july_20(1).xlsx"

# Load specific sheets
sheet1 = pd.read_excel(file_path, sheet_name="5CDCS of sem I")
sheet2 = pd.read_excel(file_path, sheet_name="2.course load")
sheet3 = pd.read_excel(file_path, sheet_name="1.Time Table")  # Additional sheet for section and STAT information

# Define branch codes and their corresponding names
branches = {
    "CS": "A7",
    "BIO": "B1",
    "CHEM": "B2",
    "ECON": "B3",
    "MATH": "B4",
    "PHY": "B5"
}

branches["CHE"] = "A1"  # Using "GEN" for general courses marked under A1

# Special cases for A7 and A1
cs_branch = "CS"
cs_code = "A7"
che_branch = "CHE"
che_code = "A1"

# Create combined branch names like B1A7, B2A7, B1A1, B2A1, etc., and standalone A7 and A1
combined_branches = {
    f"{value}{cs_code}": (value, cs_code) for value in branches.values() if value not in {cs_code, che_code}
}
combined_branches.update({
    f"{value}{che_code}": (value, che_code) for value in branches.values() if value not in {cs_code, che_code}
})
combined_branches[cs_code] = (cs_branch, cs_code)
combined_branches[che_code] = (che_branch, che_code)

# Filter and concatenate courses for each combined branch
sheet1_filtered = pd.concat([
    pd.concat([sheet1[sheet1['branch'] == cs_branch], sheet1[sheet1['branch'] == base_branch]])
    for base_branch in branches.keys()
])

# Extract course codes and years
course_codes = sheet1_filtered['SEM 1']
years = sheet1_filtered['Year']

# Create a mapping of course codes to years and branches
course_to_year_branch = {
    row['SEM 1']: (row['Year'], branches.get(row['branch'], cs_code))
    for _, row in sheet1_filtered.iterrows()
}

# Filter rows in the second sheet where course codes match
matching_courses = sheet2[sheet2['courseno'].isin(course_codes)]

# Debug: Print if no matching courses found
if matching_courses.empty:
    print("No matching courses found. Check the course codes and 'courseno' column in sheet2.")
else:
    print("Matching Courses (first 5 rows):")
    print(matching_courses.head())

# Create a dictionary to store the final JSON structure for combined branches like B1A7, B2A7, etc.
result = {
    combined_branch: {
        f"Year {year} Sem 1": {} for year in range(1, 5)
    }
    for combined_branch in combined_branches.keys()
}

# Function to map LPU values to structured format
def parse_lpu(lpu):
    # If LPU is not a string or is NaN, return default structure
    if not isinstance(lpu, str) or pd.isna(lpu):
        return {"lectures": 0, "tutorials": 0, "labs": 0}
    
    parts = lpu.split()
    if len(parts) > 2:
        lectures = int(parts[0])
        if int(parts[1]) == 0:
            tutorials = 1
            labs = 0
        else:
            tutorials = 0
            labs = int(parts[1])
    else:
        lectures = 0
        if int(parts[0]) == 0:
            tutorials = 1
            labs = 0
        else:
            tutorials = 0
            labs = int(parts[0])
    return {"lectures": lectures, "tutorials": tutorials, "labs": labs}


# Extract lab hours based on STAT and DAYS/H columns# Define a helper function to retrieve lab hours

# check here if STAT is 'P'


# Extract section counts based on SEC and STAT columns
def get_section_counts(course_code):
    sections = {"lectures": 0, "tutorials": 0, "labs": 0}
    relevant_rows = sheet3[sheet3['COURSENO'] == course_code]
    
    for _, row in relevant_rows.iterrows():
        stat = row['STAT']
        if stat == 'L':
            sections["lectures"] += 1
        elif stat == 'T':
            sections["tutorials"] += 1
        elif stat == 'P':
            sections["labs"] += 1
    return sections

def get_lab_hours(course_code):
    lab_rows = sheet3[sheet3['STAT'] == 'P'][sheet3['COURSENO'] == course_code].head(1)
    # Standardize column names
    lab_rows.columns = lab_rows.columns.str.strip()

    # Check for variations of 'DAYS/H' column
    if 'DAYS/ H' in lab_rows.columns:
        days_h_values = lab_rows['DAYS/ H'].astype(str)
    else:
        raise KeyError("Neither 'DAYS/H' nor 'DAY/H' column found in lab rows.")

    # Extract lab hours by counting unique digits in one entry
    def count_lab_hours(entry):
        # Track unique hour values by using a set
        unique_hours = set()
        for part in entry.split():
            if part.isdigit():
                unique_hours.add(part)  # Add only unique hours
        return len(unique_hours)

    # Apply the count function to each entry in days_h_values
    total_lab_hours = days_h_values.apply(count_lab_hours).sum()
    
    return total_lab_hours
# Iterate through matching courses and fill the result dictionary
for _, row in matching_courses.iterrows():
    course_code = row['courseno']
    lpu = row['LPU']
    course_name = row['coursetitle']

    year_branch = course_to_year_branch.get(course_code, (1, "B2"))
    year, base_branch_code = year_branch

    # Find the corresponding combined branch code (e.g., B1A7, B2A7, etc.)
    for combined_branch, (base_code, target_code) in combined_branches.items():
        if base_branch_code == base_code or base_branch_code == target_code:
            if (base_branch_code == target_code and base_branch_code in {"A7", "A1"}) and combined_branch not in {cs_code, che_code}:
                
                adjusted_year = year + 1
            else:
                adjusted_year = year

            key = f"Year {adjusted_year} Sem 1"
            
            if key in result[combined_branch]:
                lpu_data = parse_lpu(lpu)
                lpu_data['lab_hours'] = get_lab_hours(course_code)
                
                sections = get_section_counts(course_code)
                # if combined_branch not in {cs_code, che_code}:
                #     if sections['labs'] > 0:
                #         sections['labs'] = 1
                #     if sections['tutorials'] > 0:
                #         sections['tutorials'] = 1
                #     if sections['lectures'] > 0:
                #         sections['lectures'] = 1
                    
                
                result[combined_branch][key][course_name] = {
                    "LPU": lpu_data,
                    "Sections": sections
                }

import numpy as np

def convert_to_native(obj):
    if isinstance(obj, dict):
        return {k: convert_to_native(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_native(i) for i in obj]
    elif isinstance(obj, np.int64):  # Check for numpy int64
        return int(obj)
    elif isinstance(obj, np.float64):  # Check for numpy float64
        return float(obj)
    else:
        return obj

result_json = json.dumps(convert_to_native(result), indent=4)


with open('result.txt', 'w') as json_file:
    json_file.write(result_json)

print(result_json)
