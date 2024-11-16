import re
import pandas as pd
import json
import numpy as np

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

branches["CHE"] = "A1"  # Using "A1" for general courses marked under A1

# Special cases for A7 and A1
cs_branch = "CS"
cs_code = "A7"
che_branch = "CHE"
che_code = "A1"

# Create combined branch names like B1A7, B2A7, B1A1, B2A1, etc.
combined_branches_ordered = []

# Add standalone A7 and A1 first
combined_branches_ordered.append(cs_code)   # A7
combined_branches_ordered.append(che_code)  # A1

# Add BXA7 branches
for value in branches.values():
    if value not in {cs_code, che_code}:
        combined_branches_ordered.append(f"{value}{cs_code}")  # e.g., B1A7

# Add BXA1 branches
for value in branches.values():
    if value not in {cs_code, che_code}:
        combined_branches_ordered.append(f"{value}{che_code}")  # e.g., B1A1

# Create a dictionary mapping combined branch codes to their base and target codes
combined_branches = {}

# Add A7 and A1
combined_branches[cs_code] = (cs_branch, cs_code)
combined_branches[che_code] = (che_branch, che_code)

# Add BXA7 and BXA1
for branch_code in combined_branches_ordered[2:]:
    base_code = branch_code[:-2]
    target_code = branch_code[-2:]
    combined_branches[branch_code] = (base_code, target_code)

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

# Initialize the result structure with ordered combined branches
result = {
    combined_branch: {
        f"Year {year} Sem 1": {} for year in range(1, 5)
    }
    for combined_branch in combined_branches_ordered
}

# Function to map LPU values to structured format
def parse_lpu(lpu):
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
    lab_rows = sheet3[(sheet3['STAT'] == 'P') & (sheet3['COURSENO'] == course_code)].head(1)
    lab_rows.columns = lab_rows.columns.str.strip()

    if 'DAYS/ H' in lab_rows.columns:
        days_h_values = lab_rows['DAYS/ H'].astype(str)
    elif 'DAY/H' in lab_rows.columns:
        days_h_values = lab_rows['DAY/H'].astype(str)
    else:
        raise KeyError("Neither 'DAYS/H' nor 'DAY/H' column found in lab rows.")

    def count_lab_hours(entry):
        unique_hours = set()
        for part in entry.split():
            if part.isdigit():
                unique_hours.add(part)
        return len(unique_hours)

    total_lab_hours = days_h_values.apply(count_lab_hours).sum()
    
    return total_lab_hours

# Iterate through matching courses and fill the result dictionary
for _, row in matching_courses.iterrows():
    course_code = row['courseno']
    lpu = row['LPU']
    course_name = row['coursetitle']

    year_branch = course_to_year_branch.get(course_code, (1, "B2"))
    year, base_branch_code = year_branch

    for combined_branch in combined_branches_ordered:
        base_code, target_code = combined_branches[combined_branch]
        
        is_target_course = base_branch_code == target_code
        is_base_course = base_branch_code == base_code

        if is_target_course or is_base_course:
            adjusted_year = year + 1 if (base_branch_code == target_code and combined_branch not in {cs_code, che_code}) else year
            key = f"Year {adjusted_year} Sem 1"

            if key in result[combined_branch]:
                lpu_data = parse_lpu(lpu)
                try:
                    lpu_data['lab_hours'] = get_lab_hours(course_code)
                except KeyError as e:
                    print(f"Error processing lab hours for course {course_code}: {e}")
                    lpu_data['lab_hours'] = 0

                sections = get_section_counts(course_code)

                # For Year 1 Sem 1, directly add courses without "A7 Courses" or "Branch-Specific Courses"
                if adjusted_year == 1:
                    result[combined_branch][key][course_name] = {
                        "LPU": lpu_data,
                        "Sections": sections
                    }
                else:
                    # Ensure "A7 Courses" appear before "Branch-Specific Courses"
                    if is_target_course:
                        result[combined_branch][key].setdefault("A7 Courses", {})[course_name] = {
                            "LPU": lpu_data,
                            "Sections": sections
                        }
                    elif is_base_course:
                        result[combined_branch][key].setdefault("Branch-Specific Courses", {})[course_name] = {
                            "LPU": lpu_data,
                            "Sections": sections
                        }
                
                # Switch the order so "A7 Courses" appears before "Branch-Specific Courses" in each year
if "A7 Courses" in result[combined_branch][key] or "Branch-Specific Courses" in result[combined_branch][key]:
    reordered = {}
    if "A7 Courses" in result[combined_branch][key]:
        reordered["A7 Courses"] = result[combined_branch][key].pop("A7 Courses")
    if "Branch-Specific Courses" in result[combined_branch][key]:
        reordered["Branch-Specific Courses"] = result[combined_branch][key].pop("Branch-Specific Courses")
    result[combined_branch][key].update(reordered)


# Convert to JSON format and save
def convert_to_native(obj):
    if isinstance(obj, dict):
        return {k: convert_to_native(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_native(i) for i in obj]
    elif isinstance(obj, np.int64):
        return int(obj)
    elif isinstance(obj, np.float64):
        return float(obj)
    else:
        return obj

result_json = json.dumps(convert_to_native(result), indent=4)

with open('result.txt', 'w') as json_file:
    json_file.write(result_json)

print(result_json)
