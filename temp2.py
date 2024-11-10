import os
import random
import json
import pandas as pd

# Define the days of the week in order
ordered_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# Separate days into Mon-Wed-Fri and Tue-Thu
days_mwf = ["Monday", "Wednesday", "Friday"]
days_tt = ["Tuesday", "Thursday"]

# Available time slots (keeping 1 PM - 2 PM visible but not assignable)
time_slots = [f"{hour}:00" for hour in range(8, 13)] + ["1:00"] + [f"{hour}:00" for hour in range(2, 6)]
# Remove 8 AM for lectures and labs
time_slots_without_8am = [slot for slot in time_slots if slot != "8:00"]

# Initialize a visited matrix to track assigned time slots
visited = {day: {slot: False for slot in time_slots} for day in ordered_days}

# Global variable to store all course assignments
all_course_assignments = {}

def load_course_data(filename):
    """Load course data from a JSON file."""
    with open(filename, 'r') as file:
        return json.load(file)

def is_valid_slot(slot, is_tutorial=False):
    """Check if the slot is valid for assignment (not 1 PM to 2 PM and not 8 AM for lectures/labs)."""
    if is_tutorial:
        return slot != "1:00"
    else:
        return slot != "1:00" and slot != "8:00"

def is_valid_day_for_course(course, day, assigned_slots):
    """Check if a day is valid for assigning a new session for a course."""
    return all(assigned_day != day for assigned_day, _ in assigned_slots)

import random

# Global variables and initializations
timetable = {}  # Your timetable dictionary
visited = {}  # Tracks visited slots
all_course_assignments = {}  # Tracks assigned slots for each course
ordered_days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
days_mwf = ["Mon", "Wed", "Fri"]
days_tt = ["Tue", "Thu"]
time_slots = ["8:00", "9:00", "10:00", "11:00", "12:00", "2:00", "3:00", "4:00", "5:00"]

def assign_session(course, session_type, count, series, assigned_slots):
    """Assigns sessions like lectures, tutorials, or labs to the timetable."""
    for _ in range(count):
        while True:
            day = random.choice([d for d in series if is_valid_day_for_course(course, d, assigned_slots)])
            valid_slots = [slot for slot in time_slots if not visited[day][slot] and is_valid_slot(slot, session_type == "Tut")]
            if valid_slots:
                slot = '8:00' if session_type == "Tut" else random.choice(valid_slots)
                visited[day][slot] = True  # Mark the slot as visited
                timetable[day][slot] = f"{course} {session_type}"
                assigned_slots.append((day, slot))
                all_course_assignments.setdefault(course, []).append((day, slot, session_type))
                break

def assign_lab(course, labs, assigned_slots):
    """Assigns lab sessions for a course."""
    attempts = 0
    while attempts < 10 and labs > 0:
        day = random.choice(ordered_days)
        for start_index in range(len(time_slots) - labs + 1):
            if all(is_valid_slot(time_slots[start_index + j]) and not visited[day][time_slots[start_index + j]] for j in range(labs)):
                for j in range(labs):
                    timetable[day][time_slots[start_index + j]] = f"{course} Lab"
                    visited[day][time_slots[start_index + j]] = True
                    all_course_assignments.setdefault(course, []).append((day, time_slots[start_index + j], "Lab"))
                assigned_slots.append((day, time_slots[start_index]))
                labs -= len(range(labs))
                break
        attempts += 1

def assign_course_to_timetable(course, lectures, tutorials, lab_hours, lecture_series, tutorial_series):
    """Assigns all sessions for a course to the timetable."""
    assigned_slots = []
    remaining_lectures = lectures
    remaining_tutorials = tutorials

    if course in all_course_assignments:
        for day, slot, session_type in all_course_assignments[course]:
            if session_type == "Lecture" and remaining_lectures != 0:
                remaining_lectures -= 1
            elif session_type == "Tut":
                remaining_tutorials -= 1
            elif session_type == "Lab":
                lab_hours = 0
            timetable[day][slot] = f"{course} {session_type}"
            visited[day][slot] = True
            assigned_slots.append((day, slot))

    for day in lecture_series:
        if remaining_lectures == 0:
            break
        available_slots = ["9:00", "10:00", "4:00", "5:00"] if day in days_mwf else ["11:00", "12:00", "2:00", "3:00"]
        valid_slots = [slot for slot in available_slots if not visited[day][slot] and is_valid_slot(slot)]
        if valid_slots:
            lecture_time = random.choice(valid_slots)
            visited[day][lecture_time] = True
            timetable[day][lecture_time] = f"{course} Lecture"
            assigned_slots.append((day, lecture_time))
            all_course_assignments.setdefault(course, []).append((day, lecture_time, "Lecture"))
            remaining_lectures -= 1

            for series_day in lecture_series[1:]:
                if is_valid_day_for_course(course, series_day, assigned_slots) and not visited[series_day][lecture_time]:
                    timetable[series_day][lecture_time] = f"{course} Lecture"
                    visited[series_day][lecture_time] = True
                    assigned_slots.append((series_day, lecture_time))
                    all_course_assignments.setdefault(course, []).append((series_day, lecture_time, "Lecture"))
                    remaining_lectures -= 1

    while remaining_tutorials > 0:
        assign_session(course, "Tut", 1, tutorial_series, assigned_slots)
        remaining_tutorials -= 1

    if lab_hours > 0:
        assign_lab(course, lab_hours, assigned_slots)

    return remaining_lectures, remaining_tutorials

def assign_courses(courses):
    """Assigns all courses to the timetable according to series constraints."""
    unassigned_courses = {}
    half_courses = len(courses) // 2
    course_list = list(courses.items())
    mwf_courses = course_list[:half_courses]
    tt_courses = course_list[half_courses:]

    def generate_component_name(course, component, index):
        return f"{course}_{component}_{index}"

    # Assign MWF courses
    for course, load in mwf_courses:
        # Assign lectures
        lecture_sections = load['Sections'].get('lectures', 0)
        for i in range(1, lecture_sections + 1):
            course_variant = generate_component_name(course, "Lecture", i)
            lectures_left, _ = assign_course_to_timetable(
                course_variant, load['LPU']['lectures'], 0, 0, days_mwf, days_tt
            )
            if lectures_left > 0:
                unassigned_courses[course_variant] = {'lectures': lectures_left}

        # Assign tutorials
        tutorial_sections = load['Sections'].get('tutorials', 0)
        for i in range(1, tutorial_sections + 1):
            course_variant = generate_component_name(course, "Tutorial", i)
            _, tutorials_left = assign_course_to_timetable(
                course_variant, 0, load['LPU']['tutorials'], 0, days_mwf, days_tt
            )
            if tutorials_left > 0:
                unassigned_courses[course_variant] = {'tutorials': tutorials_left}

        # Assign labs
        lab_sections = load['Sections'].get('labs', 0)
        for i in range(1, lab_sections + 1):
            course_variant = generate_component_name(course, "Lab", i)
            _, _ = assign_course_to_timetable(
                course_variant, 0, 0, load['LPU']['lab_hours'], days_mwf, days_tt
            )

    # Assign TT courses
    for course, load in tt_courses:
        # Assign lectures
        lecture_sections = load['Sections'].get('lectures', 0)
        for i in range(1, lecture_sections + 1):
            course_variant = generate_component_name(course, "Lecture", i)
            lectures_left, _ = assign_course_to_timetable(
                course_variant, load['LPU']['lectures'], 0, 0, days_tt, days_mwf
            )
            if lectures_left > 0:
                unassigned_courses[course_variant] = {'lectures': lectures_left}

        # Assign tutorials
        tutorial_sections = load['Sections'].get('tutorials', 0)
        for i in range(1, tutorial_sections + 1):
            course_variant = generate_component_name(course, "Tutorial", i)
            _, tutorials_left = assign_course_to_timetable(
                course_variant, 0, load['LPU']['tutorials'], 0, days_tt, days_mwf
            )
            if tutorials_left > 0:
                unassigned_courses[course_variant] = {'tutorials': tutorials_left}

        # Assign labs
        lab_sections = load['Sections'].get('labs', 0)
        for i in range(1, lab_sections + 1):
            course_variant = generate_component_name(course, "Lab", i)
            _, _ = assign_course_to_timetable(
                course_variant, 0, 0, load['LPU']['lab_hours'], days_tt, days_mwf
            )

    # Attempt to reassign any unassigned components
    reassign_unassigned_courses(unassigned_courses)

def reassign_unassigned_courses(unassigned_courses):
    """Reassign any remaining unassigned sessions to available time slots."""
    for course, load in unassigned_courses.items():
        print(f"Reassigning remaining sessions for {course}...")

        lectures_left = load.get('lectures', 0)
        tutorials_left = load.get('tutorials', 0)
        labs_left = load.get('labs', 0)  # Use .get() to avoid KeyError

        # Try to reassign remaining lectures
        while lectures_left > 0:
            available_slots = [(day, slot) for day in ordered_days for slot in time_slots_without_8am
                               if not visited[day][slot] and is_valid_slot(slot)]
            if available_slots:
                day, slot = random.choice(available_slots)
                timetable[day][slot] = f"{course} Lecture"
                visited[day][slot] = True
                all_course_assignments.setdefault(course, []).append((day, slot, "Lecture"))
                lectures_left -= 1
            else:
                break  # No more available slots

        # Try to reassign remaining tutorials
        while tutorials_left > 0:
            available_slots = [(day, slot) for day in ordered_days for slot in time_slots if
                               not visited[day][slot] and is_valid_slot(slot, True)]
            if available_slots:
                day, slot = random.choice(available_slots)
                timetable[day][slot] = f"{course} Tut"
                visited[day][slot] = True
                all_course_assignments.setdefault(course, []).append((day, slot, "Tut"))
                tutorials_left -= 1
            else:
                break  # No more available slots

def generate_timetable(branch, course_data):
    """Generate the timetable for all semesters of a specific branch."""
    all_semester_data = {}
    
    if branch[0]=='A':
        for semester, courses in course_data[branch].items():
            # print(courses,semester)

            print(f"\nAssigning courses for {branch} - {semester}...")
            global timetable
            timetable = {day: {slot: None for slot in time_slots} for day in ordered_days}
            
            global visited
            visited = {day: {slot: False for slot in time_slots} for day in ordered_days}  # Reset visited matrix

            assign_courses(courses)
            
            # Store each semester with branch information
            if semester not in all_semester_data:
                all_semester_data[semester] = {}    

            # Combine multiple branches under the same semester
            all_semester_data[semester][branch] = timetable.copy()
            
        
    else:
        for semester, courses in course_data[branch].items():
            print(f"\nAssigning courses for {branch} - {semester}...")
            
            timetable = {day: {slot: None for slot in time_slots} for day in ordered_days}
            
            
            visited = {day: {slot: False for slot in time_slots} for day in ordered_days}  # Reset visited matrix

            assign_courses(courses)
            
            # Store each semester with branch information
            if semester not in all_semester_data:
                all_semester_data[semester] = {}

            # Combine multiple branches under the same semester
            all_semester_data[semester][branch] = timetable.copy()

    return all_semester_data

def save_timetable_to_excel(all_semester_data):
    """Save the combined timetable for all branches into an Excel file, separated by semesters."""
    with pd.ExcelWriter("combined_timetable2.xlsx", engine='xlsxwriter') as writer:
        for semester, branches in all_semester_data.items():
            # Create a DataFrame for each semester, combining branches with an empty row separating them
            combined_df = pd.DataFrame()

            for branch, timetable in branches.items():
                df = pd.DataFrame(timetable)
                df.index.name = 'Day'
                df.columns.name = 'Time Slot'
                
                # Add branch name as a header to separate data visually
                branch_header = pd.DataFrame([[f"Branch: {branch}"]], columns=[None])
                combined_df = pd.concat([combined_df, branch_header, df, pd.DataFrame([[]])])
                    # Add an empty row after each branch's data
            
            # Write the combined data for the semester into the sheet
            combined_df.to_excel(writer, sheet_name=f"Semester {semester}", index=True, header=True)

        # Add a new sheet with all course assignments
        course_assignments_df = pd.DataFrame(
            [(course, day, slot, session_type) for course, assignments in all_course_assignments.items() for day, slot, session_type in assignments],
            columns=['Course', 'Day', 'Time', 'Type']
        )
        course_assignments_df.to_excel(writer, sheet_name="All Course Assignments", index=False)

def main():
    """Main function to start the program."""
    course_data = load_course_data('result.txt')
    all_branch_data = {}

    # Generate timetables for all branches and semesters
    for branch in course_data:
        branch_timetable = generate_timetable(branch, course_data)
        
        # Merge branch data into the overall semester structure
        for semester, semester_data in branch_timetable.items():
            if semester not in all_branch_data:
                all_branch_data[semester] = {}
            all_branch_data[semester].update(semester_data)

    save_timetable_to_excel(all_branch_data)

if __name__ == "__main__":
    main()