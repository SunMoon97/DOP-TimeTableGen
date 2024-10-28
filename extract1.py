import pandas as pd
import json

# Load the entire Excel file
file_path = r"/Users/adi-2310/College/CS DOP/DOP-TimeTableGen/Copy of Data_for_Time_Table_software_-27_july_20(1).xlsx"

# Load specific sheets
sheet1 = pd.read_excel(file_path, sheet_name="5CDCS of sem I")
sheet2 = pd.read_excel(file_path, sheet_name="2.course load")
sheet3 = pd.read_excel(file_path, sheet_name="1.Time Table")  # Sheet containing 'SEC' column for number of sections

# Debug: Check columns in sheet3
print("Columns in sheet3:", sheet3.columns.tolist())

# Define branch codes and their corresponding names
branches = {
    "CS": "A7",
    "BIO": "B1",
    "CHEM": "B2",
    "ECON": "B3",
    "MATH": "B4",
    "PHY": "B5"
}
branches["CHE"] = "A1"  # General courses marked under A1

# Special cases for A7 and A1
cs_branch = "CS"
cs_code = "A7"
che_branch = "CHE"
che_code = "A1"

# Create combined branch names like B1A7, B2A7, etc., and standalone A7 and A1
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

# Calculate the maximum number of sections for each course from sheet3
# Ensure that 'courseno' and 'SEC' match the actual names in sheet3
max_sections = sheet3.groupby('courseno')['SEC'].max().to_dict()  # Group by 'courseno' to find max SEC values

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

# Iterate through matching courses and fill the result dictionary
for _, row in matching_courses.iterrows():
    course_code = row['courseno']
    lpu = row['LPU']
    course_name = row['coursetitle']  # Use 'coursetitle' as the course name
    max_sec = max_sections.get(course_code, 1)  # Get maximum sections for this course

    # Get the year and base branch code from the course_to_year_branch dictionary
    year_branch = course_to_year_branch.get(course_code, (1, "B2"))  # Default to Year 1, Branch B2 if not found
    year, base_branch_code = year_branch

    # Find the corresponding combined branch code (e.g., B1A7 , B2A7, B1A1, B2A1, etc.)
    for combined_branch, (base_code, target_code) in combined_branches.items():
        if base_branch_code == base_code or base_branch_code == target_code:
            # Adjust the year if it's a combined branch with A7 or A1
            if (base_branch_code == target_code and base_branch_code in {"A7", "A1"}) and combined_branch not in {cs_code, che_code}:
                adjusted_year = year + 1
            else:
                adjusted_year = year

            # Create the key with the adjusted year
            key = f"Year {adjusted_year} Sem 1"
            
            # Debug: Check if the key is in the result dictionary
            if key in result[combined_branch]:
                print(f"Adding course {course_name} to {combined_branch}, {key} with LPU: {lpu}")
                
                # Populate result dictionary with course data, including number of sections
                result[combined_branch][key][course_name] = {
                    **parse_lpu(lpu),  # Add LPU structure
                    "number_of_sections": max_sec  # Include the max number of sections
                }
            else:
                print(f"Key {key} not found in result dictionary for branch {combined_branch}.")
   
# Convert the result to JSON format
result_json = json.dumps(result, indent=4)

# Save the result to a file or print it
with open('result.txt', 'w') as json_file:
    json_file.write(result_json)

print(result_json)