import random
import json
from tabulate import tabulate

# Days of the week and available time slots (excluding 1 PM - 2 PM)
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
time_slots = [f"{hour}:00" for hour in range(8, 13)] + [f"{hour + 1}:00" for hour in range(1, 6)]

# Dictionary to store assigned time slots for each course
timetable = {day: {slot: None for slot in time_slots} for day in days}

# Load course data from a file
def load_course_data(filename):
    with open(filename, 'r') as file:
        return json.load(file)

# Assign a course to the timetable, considering predefined lab hours
def assign_course_to_timetable(course, lectures, tutorials, labs):
    assigned_slots = []
    days_used = set()

    # Assign lectures (1 per day)
    for _ in range(lectures):
        while True:
            day = random.choice([d for d in days if d not in days_used])
            slot = random.choice(time_slots)
            if timetable[day][slot] is None:
                timetable[day][slot] = course
                assigned_slots.append((day, slot))
                days_used.add(day)
                break

    # Assign tutorials
    for _ in range(tutorials):
        while True:
            day = random.choice(days)
            slot = random.choice(time_slots)
            if timetable[day][slot] is None:
                timetable[day][slot] = course + " Tut"
                assigned_slots.append((day, slot))
                break

    # Assign lab (fixed duration, once a week)
    if labs > 0:
        while True:
            day = random.choice(days)
            
            if any(course and "Lab" in course for course in timetable[day].values()):
                continue
            valid_time_slots = [i for i in range(len(time_slots) - labs + 1) if time_slots[i] != "12:00" and time_slots[i] != "8:00"]
            if not valid_time_slots:
                continue
            slot_index = random.choice(valid_time_slots)
            if all(timetable[day][time_slots[slot_index + i]] is None for i in range(labs)):  # Check for consecutive free slots
                # Check if the entire lab duration is available
                for i in range(labs):
                    timetable[day][time_slots[slot_index + i]] = course + " Lab"
                assigned_slots.append((day, time_slots[slot_index]))
                break

    return assigned_slots

# Generate the timetable for all semesters
def generate_timetable(branch, course_data):
    for semester, courses in course_data[branch].items():
        print(f"\nAssigning courses for {branch} - {semester}...")

        # Reset timetable for each semester
        global timetable
        timetable = {day: {slot: None for slot in time_slots} for day in days}

        # Assign courses to the timetable
        for course, load in courses.items():
            assign_course_to_timetable(course, lectures=load['lectures'], tutorials=load['tutorials'], labs=load['labs'])

        # Save the timetable for this semester to a file
        save_timetable_to_file(semester)

# Display the generated timetable in tabular format (slots as rows, days as columns)
def save_timetable_to_file(semester):
    filename = 'output.txt'
    with open(filename, 'a') as file:
        file.write(f"\nTimetable for {semester}:\n")
        table = []
        for slot in time_slots:
            row = [slot]
            for day in days:
                course = timetable[day][slot]
                row.append(course if course else "")
            table.append(row)

        file.write(tabulate(table, headers=["Time Slot"] + days, tablefmt="grid"))
        file.write("\n")

# Main function to start the program
def main():
    # Load the course data from the file
    course_data = load_course_data('data.txt')

    # Generate timetables for all semesters of the branch
    branch = 'A7'  # For computer science branch
    generate_timetable(branch, course_data)

# Run the program
main()
