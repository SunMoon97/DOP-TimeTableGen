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

def load_course_data(filename):
    """Load course data from a JSON file."""
    with open(filename, 'r') as file:
        return json.load(file)

def is_valid_slot(slot, is_tutorial=False):
    """Check if the slot is valid for assignment (not 1 PM to 2 PM and not 8 AM for lectures/labs)."""
    #if there is a tutorial i want it to assihn  to the 8 am slot only
        
    if is_tutorial:
        return slot != "1:00"
    else:
        return slot != "1:00" and slot != "8:00"

def is_valid_day_for_course(course, day, assigned_slots):
    """Check if a day is valid for assigning a new session for a course."""
    return all(assigned_day != day for assigned_day, _ in assigned_slots)

def assign_session(course, session_type, count, series, assigned_slots):
    """Assign a session (lecture, tutorial, or lab) to the timetable."""
    for _ in range(count):
        while True:
            day = random.choice([d for d in series if is_valid_day_for_course(course, d, assigned_slots)])
            valid_slots = [slot for slot in time_slots if not visited[day][slot] and is_valid_slot(slot, session_type == "Tut")]
            if valid_slots:
                slot = random.choice(valid_slots)
                visited[day][slot] = True  # Mark the slot as visited
                timetable[day][slot] = f"{course} {session_type}"
                assigned_slots.append((day, slot))
                break

def assign_lab(course, labs, assigned_slots):
    """Assign lab sessions for a course."""
    attempts = 0
    while attempts < 10 and labs > 0:
        day = random.choice(ordered_days)
        # Check for a continuous block of slots for labs
        for start_index in range(len(time_slots) - labs + 1):
            if all(is_valid_slot(time_slots[start_index + j]) and not visited[day][time_slots[start_index + j]] for j in range(labs)):
                for j in range(labs):
                    timetable[day][time_slots[start_index + j]] = f"{course} Lab"
                    visited[day][time_slots[start_index + j]] = True  # Mark as visited
                assigned_slots.append((day, time_slots[start_index]))  # Record the starting time slot
                labs -= len(range(labs))  # Deduct the assigned labs
                break
        attempts += 1

def assign_course_to_timetable(course, lectures, tutorials, lab_hours, lecture_series, tutorial_series):
    """Assign all sessions for a course to the timetable."""
    assigned_slots = []
    remaining_lectures = lectures
    remaining_tutorials = tutorials

    # Assign lectures
    for day in lecture_series:
        if remaining_lectures == 0:
            break
        valid_slots = [slot for slot in time_slots_without_8am if not visited[day][slot] and is_valid_slot(slot)]
        if valid_slots:
            lecture_time = random.choice(valid_slots)
            visited[day][lecture_time] = True  # Mark as visited
            timetable[day][lecture_time] = f"{course} Lecture"
            assigned_slots.append((day, lecture_time))
            remaining_lectures -= 1  # Decrement lecture counter

            # Assign the same lecture slot to corresponding series days
            for series_day in lecture_series[1:]:
                if is_valid_day_for_course(course, series_day, assigned_slots) and not visited[series_day][lecture_time]:
                    timetable[series_day][lecture_time] = f"{course} Lecture"
                    visited[series_day][lecture_time] = True  # Mark as visited
                    assigned_slots.append((series_day, lecture_time))
                    remaining_lectures -= 1  # Decrement lecture counter

    # Assign tutorials
    while remaining_tutorials > 0:
        assign_session(course, "Tut", 1, tutorial_series, assigned_slots)  # Assign one tutorial at a time
        remaining_tutorials -= 1  # Decrement tutorial counter

    # Assign lab (labs are only one per course, with varying hours)
    if lab_hours > 0:
        assign_lab(course, lab_hours, assigned_slots)

    return remaining_lectures, remaining_tutorials

def assign_courses(courses):
    """Assign all courses to the timetable based on their series."""
    unassigned_courses = {}

    # Assign Mon-Wed-Fri and Tue-Thu courses
    half_courses = len(courses) // 2
    course_list = list(courses.items())
    mwf_courses = course_list[:half_courses]
    tt_courses = course_list[half_courses:]

    for course, load in mwf_courses:
        lectures_left, tutorials_left = assign_course_to_timetable(
            course, load['lectures'], load['tutorials'], load['labs'], days_mwf, days_tt
        )

        # Check if any unassigned sessions are left
        if lectures_left > 0 or tutorials_left > 0 :
            unassigned_courses[course] = {
                'lectures': lectures_left,
                'tutorials': tutorials_left,
                
            }
            
    for course, load in tt_courses:
        #in days tt series add one of the mwf days randomly
        days_tt.append(random.choice([d for d in days_mwf if is_valid_day_for_course(course, d, [])]))
        lectures_left, tutorials_left = assign_course_to_timetable(
            course, load['lectures'], load['tutorials'], load['labs'], days_tt, days_mwf
        )
        days_tt.pop()

        # Check if any unassigned sessions are left
        if lectures_left > 0 or tutorials_left > 0:
            unassigned_courses[course] = {
                'lectures': lectures_left,
                'tutorials': tutorials_left,
            }

    # Try to reassign unassigned sessions if needed
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
                lectures_left -= 1
            else:
                break  # No more available slots

        # Try to reassign remaining tutorials
        while tutorials_left > 0:
            assign_session(course, "Tut", 1, ordered_days, [])  # Assign tutorial to any available day
            tutorials_left -= 1

        # Try to reassign remaining labs
        while labs_left > 0:
            assign_lab(course, 1, [])  # Assign lab to any available day
            labs_left -= 1

        if lectures_left == 0 and tutorials_left == 0 and labs_left == 0:
            print(f"Successfully reassigned all sessions for {course}.")
        else:
            print(f"Could not reassign all sessions for {course}. Please check availability.")

def generate_timetable(branch, course_data):
    """Generate the timetable for all semesters of a specific branch."""
    all_semester_data = {}

    for semester, courses in course_data[branch].items():
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

    return all_semester_data

def save_timetable_to_excel(all_semester_data):
    """Save the combined timetable for all branches into an Excel file, separated by semesters."""
    with pd.ExcelWriter("combined_timetable.xlsx", engine='xlsxwriter') as writer:
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
