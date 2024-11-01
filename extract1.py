import pandas as pd
import json

# Load the entire Excel file
file_path = r"/Users/adi-2310/College/CS DOP/DOP-TimeTableGen/Copy of Data_for_Time_Table_software_-27_july_20(1).xlsx"

# Load specific sheets
sheet1 = pd.read_excel(file_path, sheet_name="5CDCS of sem I")
sheet2 = pd.read_excel(file_path, sheet_name="2.course load")
# Load sheet3 to get 'coursename' and 'SEC' columns for section counts
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
    row['SEM 1']: (row['Year'], branches.get(row['branch'], cs_code),
                   max_sections_per_course.get(row['coursename'], 1)  # Include max sections
                   )
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

# Prepare the result dictionary to be converted to JSON
result = {}

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

    # Get the year, base branch code, and max number of sections
    year, base_branch_code, num_sections = course_to_year_branch.get(course_code, (1, "B2", 1))

    # Find the corresponding combined branch code (e.g., B1A7, B2A7, B1A1, B2A1, etc.)
    for combined_branch, (base_code, target_code) in combined_branches.items():
        if base_branch_code == base_code or base_branch_code == target_code:
            # Adjust the year if it's a combined branch with A7 or A1
            if (base_branch_code == target_code and base_branch_code in {"A7", "A1"}) and combined_branch not in {cs_code, che_code}:
                adjusted_year = year + 1
            else:
                adjusted_year = year

            # Create the key with the adjusted year
            key = f"Year {adjusted_year} Sem 1"
            
            # Initialize branch and year structure in result if not present
            if combined_branch not in result:
                result[combined_branch] = {}
            if key not in result[combined_branch]:
                result[combined_branch][key] = {}

            # Add course details with number of sections
            result[combined_branch][key][course_name] = {
                **parse_lpu(lpu),
                "number of sections": num_sections
            }

# Convert the result to JSON format
result_json = json.dumps(result, indent=4)

# Write the result to a JSON file
with open('result1.txt', 'w') as f:
    json.dump(result, f, indent=4)

print(result_json)
