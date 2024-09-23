import random
import json
import pandas as pd
import time

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
    return slot != "1:00" and (slot != "8:00" or is_tutorial)

def is_valid_day_for_course(course, day, assigned_slots):
    """Check if a day is valid for assigning a new session for a course."""
    return all(assigned_day != day for assigned_day, _ in assigned_slots)

def block_consecutive_slots(day, start_slot, count):
    """Block consecutive time slots for the lab."""
    start_index = time_slots.index(start_slot)
    for i in range(start_index, start_index + count):
        visited[day][time_slots[i]] = True  # Mark consecutive slots as visited

def unassign_slots(day, start_slot, count):
    """Unassign consecutive time slots (for backtracking)."""
    start_index = time_slots.index(start_slot)
    for i in range(start_index, start_index + count):
        visited[day][time_slots[i]] = False  # Unmark the slots as visited

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
            else:
                if not backtrack_on_failure(course, assigned_slots):
                    return False  # If backtrack fails, return false

def backtrack_on_failure(course, assigned_slots):
    """Perform backtracking when no valid slot is found."""
    if not assigned_slots:
        return False  # No assignments left to backtrack from, failure

    last_assigned_day, last_assigned_slot = assigned_slots.pop()
    timetable[last_assigned_day][last_assigned_slot] = None  # Remove from timetable
    visited[last_assigned_day][last_assigned_slot] = False  # Mark slot as unvisited
    
    return True  # Backtracked successfully

def assign_lab(course, labs, no_of_labs, lecture_series, assigned_slots):
    """Assign lab sessions for a course, ensuring consecutive time slots are blocked."""
    attempts = 0
    while attempts < 20 and no_of_labs > 0:
        day = random.choice(lecture_series)
        if any(day == assigned_day for assigned_day, _ in assigned_slots):
            continue  # Avoid assigning labs on the same day as lectures

        for start_index in range(len(time_slots) - labs + 1):
            if all(is_valid_slot(time_slots[start_index + j]) and not visited[day][time_slots[start_index + j]] for j in range(labs)):
                for j in range(labs):
                    timetable[day][time_slots[start_index + j]] = f"{course} Lab"
                    visited[day][time_slots[start_index + j]] = True  # Mark as visited
                assigned_slots.append((day, time_slots[start_index]))  # Record the starting time slot
                no_of_labs -= 1  # Deduct one session for this lab assignment
                return True  # Successfully assigned lab
        attempts += 1

    # If no valid assignment is found, attempt backtracking
    return backtrack_on_failure(course, assigned_slots)

def assign_course_to_timetable(course, lectures, tutorials, labs, no_of_labs, lecture_series, tutorial_series):
    """Assign all sessions for a course to the timetable."""
    assigned_slots = []
    remaining_lectures = lectures

    # Assign lectures
    for day in lecture_series:
        if remaining_lectures == 0 or not is_valid_day_for_course(course, day, assigned_slots):
            break
        valid_slots = [slot for slot in time_slots_without_8am if not visited[day][slot] and is_valid_slot(slot)]
        if valid_slots:
            lecture_time = random.choice(valid_slots)
            visited[day][lecture_time] = True  # Mark as visited
            timetable[day][lecture_time] = f"{course} Lecture"
            assigned_slots.append((day, lecture_time))
            remaining_lectures -= 1

            # Assign the same lecture slot to corresponding series days
            for series_day in lecture_series[1:]:
                if is_valid_day_for_course(course, series_day, assigned_slots) and not visited[series_day][lecture_time]:
                    timetable[series_day][lecture_time] = f"{course} Lecture"
                    visited[series_day][lecture_time] = True  # Mark as visited
                    assigned_slots.append((series_day, lecture_time))

    # Assign tutorials
    if tutorials > 0:
        assign_session(course, "Tut", tutorials, tutorial_series, assigned_slots)

    # Assign labs
    if labs > 0 and no_of_labs > 0:
        if not assign_lab(course, labs, no_of_labs, lecture_series, assigned_slots):
            while backtrack_on_failure(course, assigned_slots):
                if assign_lab(course, labs, no_of_labs, lecture_series, assigned_slots):
                    break

    return assigned_slots

def assign_courses(courses):
    """Assign all courses to the timetable based on their series."""
    assigned_mwf = []
    assigned_tt = []

    half_courses = len(courses) // 2
    course_list = list(courses.items())
    mwf_courses = course_list[:half_courses]
    tt_courses = course_list[half_courses:]

    # Assign Mon-Wed-Fri courses
    for course, load in mwf_courses:
        assign_course_to_timetable(course, load['lectures'], load['tutorials'], load['labs'], load['no_of_labs'], days_mwf, days_tt)

    # Assign Tue-Thu courses
    for course, load in tt_courses:
        assign_course_to_timetable(course, load['lectures'], load['tutorials'], load['labs'], load['no_of_labs'], days_tt, days_mwf)

def save_timetable_to_excel(all_semester_data):
    """Save the combined timetable for all semesters into an Excel file."""
    with pd.ExcelWriter("complete_timetableb2a7.xlsx", engine='xlsxwriter') as writer:
        for semester, timetable in all_semester_data.items():
            df = pd.DataFrame(timetable)
            df.index.name = 'Day'
            df.columns.name = 'Time Slot'
            df.to_excel(writer, sheet_name=semester)

def generate_timetable(branch, course_data):
    """Generate the timetable for all semesters of a specific branch."""
    all_semester_data = {}

    for semester, courses in course_data[branch].items():
        print(f"\nAssigning courses for {branch} - {semester}...")
        global timetable
        timetable = {day: {slot: None for slot in time_slots} for day in ordered_days}
        
        global visited
        visited = {day: {slot: False for slot in time_slots} for day in ordered_days}  # Reset visited matrix

        # Add timeout mechanism
        start_time = time.time()
        timeout_duration = 10  # Set a timeout duration (in seconds)

        try:
            assign_courses(courses)
            all_semester_data[semester] = timetable.copy()

            # Check if timeout occurred
            if time.time() - start_time > timeout_duration:
                print("Timeout occurred while assigning courses.")
                return None

        except Exception as e:
            print(f"An error occurred during course assignment: {e}")
            return None

    save_timetable_to_excel(all_semester_data)
    return all_semester_data

def main():
    """Main function to start the program."""
    course_data = load_course_data('B2A7.txt')
    branch = 'B2A7'  # For computer science branch

    try:
        all_semester_data = generate_timetable(branch, course_data)
        
        if all_semester_data and all(all_semester_data[semester].values() for semester in all_semester_data):
            print("Complete timetable found!")

    except Exception as e:
        print(f"An error occurred: {e}.")

if __name__ == "__main__":
    main()
