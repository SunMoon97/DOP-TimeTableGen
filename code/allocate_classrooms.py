import pandas as pd

# Load the input Excel files
combined_timetable = pd.read_excel("combined_timetable.xlsx", sheet_name="All Course Assignments")
erp_data = pd.read_excel("sem 1 23-24 erp tt& reg.xlsx", sheet_name="Sheet1", header=1)
rooms_data = pd.read_excel("data.xlsx", sheet_name="3.rooms")

# Merge Subject + Catalog to form the full course code
erp_data["Course"] = (erp_data["Subject"].astype(str).str.strip() + ' ' + erp_data["Catalog"].astype(str).str.strip()).str.replace('  ', ' ')

# Replace double space with single space

# Extract relevant columns
course_enrollment = erp_data[["Course", "No. Of Students"]]
print(course_enrollment["Course"][0])
room_capacity = rooms_data[["Room", "Seating Capacity"]]

# Convert seating capacity to numeric
room_capacity = room_capacity.copy()
room_capacity["Seating Capacity"] = pd.to_numeric(room_capacity["Seating Capacity"], errors='coerce')

# Sort rooms by ascending seating capacity for best-fit allocation
room_capacity = room_capacity.sort_values(by="Seating Capacity")

# Initialize classroom allocation dictionary
classroom_allocation = {}
room_schedule = {}

# Allocate classrooms based on best fit
for index, row in combined_timetable.iterrows():
    course = row["Course"]
    day = row["Day"]
    time = row["Time"]
    session_type = row["Type"]
    course = row["Course"].split('-')[0]
    # print(course)
    
    # Get student count for the course
    student_count = course_enrollment.loc[course_enrollment["Course"] == course, "No. Of Students"].values
    if len(student_count) == 0:
        continue  # Skip if course is not found in enrollment data
    student_count = student_count[0]
    
    # Find the best-fit room
    suitable_room = None
    for _, room in room_capacity.iterrows():
        if room["Seating Capacity"] >= student_count:
            suitable_room = room["Room"]
            break
    
    if suitable_room:
        classroom_allocation[(course, day, time, session_type)] = suitable_room
        room_schedule.setdefault((day, time), []).append((suitable_room, course))
    else:
        classroom_allocation[(course, day, time, session_type)] = "No Room Available"

# Create DataFrames for output
allocation_df = pd.DataFrame(
    [(k[0], k[1], k[2], k[3], v) for k, v in classroom_allocation.items()],
    columns=["Course", "Day", "Time", "Type", "Room"]
)
schedule_df = pd.DataFrame(
    [(k[0], k[1], v[0], v[1]) for k, val in room_schedule.items() for v in val],
    columns=["Day", "Time", "Room", "Course"]
)

print(allocation_df.head())
print(schedule_df.head())

# Save outputs to Excel
writer = pd.ExcelWriter("classroom_allocation.xlsx")
allocation_df.to_excel(writer, sheet_name="Course to Room Mapping", index=False)
schedule_df.to_excel(writer, sheet_name="Full Room Schedule", index=False)
writer.close()


print("Classroom allocation completed! Check classroom_allocation.xlsx")
