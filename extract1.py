import pandas as pd
import json

# Load the Excel file
file_path = r"C:\Users\mundh\Documents\DOP-timetable-generator\Copy of Data_for_Time_Table_software_-27_july_20(1).xlsx"

# Load sheets into DataFrames
sheet1 = pd.read_excel(file_path, sheet_name="5CDCS of sem I")
sheet2 = pd.read_excel(file_path, sheet_name="2.course load")
sheet3 = pd.read_excel(file_path, sheet_name='1.Time Table')

# Get the maximum number of sections for each course from 'SEC' column in sheet3
max_sections_per_course = sheet3.groupby('COURSETITLE')['SEC'].max().to_dict()

# Define branch codes and their corresponding names
branches = {
    "CS": "A7",
    "BIO": "B1",
    "CHEM": "B2",
    "ECON": "B3",
    "MATH": "B4",
    "PHY": "B5",
    "CHE": "A1"
}

# Create combined branch names
combined_branches = {
    f"{value}{branches['CS']}": (value, branches['CS']) for value in branches.values() if value != branches['CS']
}
combined_branches.update({
    f"{value}{branches['CHE']}": (value, branches['CHE']) for value in branches.values() if value != branches['CHE']
})
combined_branches[branches['CS']] = (branches['CS'], branches['CS'])
combined_branches[branches['CHE']] = (branches['CHE'], branches['CHE'])

# Filter and concatenate courses for each combined branch
sheet1_filtered = pd.concat([
    pd.concat([sheet1[sheet1['branch'] == branches['CS']], sheet1[sheet1['branch'] == base_branch]])
    for base_branch in branches.keys()
])

# Extract course codes and years
course_codes = sheet1_filtered['SEM 1']
years = sheet1_filtered['Year']

# Create a mapping of course codes to years and branches
course_to_year_branch = {
    row['SEM 1']: (row['Year'], branches.get(row['branch'], branches['CS']),
                   max_sections_per_course.get(row['coursename'], 1))
    for _, row in sheet1_filtered.iterrows()
}

# Filter rows in the second sheet where course codes match
matching_courses = sheet2[sheet2['courseno'].isin(course_codes)]

if matching_courses.empty:
    print("No matching courses found. Check the course codes and 'courseno' column in sheet2.")
else:
    print("Matching Courses (first 5 rows):")
    print(matching_courses.head())

# Prepare the result dictionary to be converted to JSON
result = {
    combined_branch: {
        f"Year {year} Sem 1": {} for year in range(1, 5)
    }
    for combined_branch in combined_branches.keys()
}

# Function to map LPU values to structured format
def parse_lpu(lpu):
    """Parse the LPU string and return a dictionary of lectures, tutorials, and labs."""
    if not isinstance(lpu, str) or pd.isna(lpu):
        return {"lectures": 0, "tutorials": 0, "labs": 0}
    
    parts = lpu.split()
    if len(parts) > 2:
        lectures = int(parts[0])
        tutorials = 1 if int(parts[1]) == 0 else 0
        labs = 0 if tutorials == 1 else int(parts[1])
    else:
        lectures = 0
        tutorials = 1 if int(parts[0]) == 0 else 0
        labs = 0 if tutorials == 1 else int(parts[0])
    
    return {"lectures": lectures, "tutorials": tutorials, "labs": labs}

# Iterate through matching courses and fill the result dictionary
for _, row in matching_courses.iterrows():
    course_code = row['courseno']
    lpu = row['LPU']
    course_name = row['coursetitle'] 

    # Get the year, base branch code, and max number of sections
    year, base_branch_code, num_sections = course_to_year_branch.get(course_code, (1, "CS", 1))
    

    # Find the corresponding combined branch code (e.g., B1A7, B2A7, etc.)
    for combined_branch, (base_code, target_code) in combined_branches.items():
        if base_branch_code in (base_code, target_code):
            # Adjust num_sections for combined branches to 1
            if combined_branch not in {branches['CS'], branches['CHE']}:
                num_sections = 1

            # Adjust the year for combined branches if needed
            adjusted_year = year if not (base_branch_code == target_code and base_branch_code in {branches['CS'], branches['CHE']}) else year + 1

            # Create the key with the adjusted year
            key = f"Year {adjusted_year} Sem 1"
            
            # Initialize branch and year structure in result if not present
            result.setdefault(combined_branch, {}).setdefault(key, {})

            # Add course details
            result[combined_branch][key][course_name] = {
                **parse_lpu(lpu),
                "number of sections": num_sections,
                "number of sections parallel": 1  # Single parallel section for combined branches
            }
            


# Convert the result to JSON format
result_json = json.dumps(result, indent=4)

# Write the result to a JSON file
with open('result1.json', 'w') as f:
    json.dump(result, f, indent=4)

print(result_json)
