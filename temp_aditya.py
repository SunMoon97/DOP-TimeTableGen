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
    return slot != "1:00" and (slot != "8:00" or is_tutorial)

def is_valid_day_for_course(course, day, assigned_slots):
    """Check if a day is valid for assigning a new session for a course."""
    return all(assigned_day != day for assigned_day, _ in assigned_slots)

def assign_session(course, session_type, count, series, assigned_slots):
    """Assign a session (lecture, tutorial, or lab) to the timetable."""
    for _ in range(count):
        while True:
            day = random.choice([d for d in series if is_valid_day_for_course(course, d, assigned_slots)])
            valid_slots = [slot for slot in time_slots if not visited[day][slot] and is_valid_slot(slot, session_type == "Tut")]
            if valid_slots :
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
                    if visited[day][time_slots[start_index + j]] == False:
                        timetable[day][time_slots[start_index + j]] = f"{course} Lab"
                        visited[day][time_slots[start_index + j]] = True  # Mark as visited
                assigned_slots.append((day, time_slots[start_index]))  # Record the starting time slot
                labs -= len(range(labs))  # Deduct the assigned labs
                break
        attempts += 1

    # If not all labs are assigned, attempt to find slots on alternate days
    # while labs > 0:
    #     day = random.choice(ordered_days)
    #     valid_slots = [slot for slot in time_slots if is_valid_slot(slot) and not visited[day][slot]]
    #     if valid_slots:
    #         slot = random.choice(valid_slots)
    #         timetable[day][slot] = f"{course} Lab"
    #         visited[day][slot] = True  # Mark as visited
    #         assigned_slots.append((day, slot))
    #         labs -= 1  # Deduct assigned lab

def assign_course_to_timetable(course, lectures, tutorials, lab_hours, lecture_series_mwf, lecture_series_tt, tutorial_series_mwf, tutorial_series_tt):
    """Assign all sessions for a course to the timetable."""
    assigned_slots = []
    remaining_lectures = lectures
    remaining_tutorials = tutorials

    # Assign lectures
    for day in lecture_series_mwf:
        if remaining_lectures == 0:
            break
        valid_slots = [slot for slot in time_slots_without_8am if not visited[day][slot] and is_valid_slot(slot)]
        if valid_slots:
            lecture_time = random.choice(valid_slots)
            visited[day][lecture_time] = True  # Mark as visited
            timetable[day][lecture_time] = f"{course} Lecture"
            assigned_slots.append((day, lecture_time))
            remaining_lectures -= 1  # Decrement lecture counter

            # Assign the same lecture slot to corresponding series days (MWF)
            for series_day in lecture_series_mwf[1:]:
                if is_valid_day_for_course(course, series_day, assigned_slots) and not visited[series_day][lecture_time]:
                    timetable[series_day][lecture_time] = f"{course} Lecture"
                    visited[series_day][lecture_time] = True  # Mark as visited
                    assigned_slots.append((series_day, lecture_time))
                    remaining_lectures -= 1  # Decrement lecture counter

    # If lectures are still remaining, assign to Tue/Thu
    while remaining_lectures > 0:
        assign_session(course, "Lecture", remaining_lectures, lecture_series_tt, assigned_slots)
        remaining_lectures = 0  # All lectures assigned

    # Assign tutorials
    # MWF tutorials are assigned to 8:00 AM
    for day in tutorial_series_mwf:
        if remaining_tutorials == 0:
            break
        if not visited[day]["8:00"] and is_valid_slot("8:00", is_tutorial=True):
            visited[day]["8:00"] = True
            timetable[day]["8:00"] = f"{course} Tut"
            assigned_slots.append((day, "8:00"))
            remaining_tutorials -= 1

    # If MWF tutorials are still remaining or for Tue/Thu tutorials, assign to any available slot
    while remaining_tutorials > 0:
        assign_session(course, "Tut", 1, tutorial_series_tt, assigned_slots)
        remaining_tutorials -= 1

    # Assign lab (labs are only one per course, with varying hours)
    if lab_hours > 0:
        assign_lab(course, lab_hours, assigned_slots)

    return remaining_lectures, remaining_tutorials



def assign_courses(courses):
    """Assign all courses to the timetable based on their series."""
    unassigned_courses = {}

    # Assign Mon-Wed-Fri and Tue-Thu courses
    for course, load in courses.items():
        lectures_left, tutorials_left = assign_course_to_timetable(
            course, load['lectures'], load['tutorials'], load['labs'], days_mwf, days_tt,days_mwf,days_tt
        )

        # Check if any unassigned sessions are left
        if lectures_left > 0 or tutorials_left > 0:
            unassigned_courses[course] = {
                'lectures': lectures_left,
                'tutorials': tutorials_left
            }

    # Try to reassign unassigned sessions if needed
    reassign_unassigned_courses(unassigned_courses)



def reassign_unassigned_courses(unassigned_courses):
    """Reassign any remaining unassigned sessions to available time slots."""
    for course, load in unassigned_courses.items():
        print(f"Reassigning remaining sessions for {course}...")

        lectures_left, tutorials_left, labs_left = load['lectures'], load['tutorials'], load['labs']

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
        all_semester_data[semester] = timetable.copy()

    save_timetable_to_excel(all_semester_data)


def save_timetable_to_excel(all_semester_data):
    """Save the combined timetable for all semesters into an Excel file."""
    with pd.ExcelWriter("complete_timetableb2a72_aditya.xlsx", engine='xlsxwriter') as writer:
        for semester, timetable in all_semester_data.items():
            df = pd.DataFrame(timetable)
            df.index.name = 'Day'
            df.columns.name = 'Time Slot'
            df.to_excel(writer, sheet_name=semester)

def main():
    """Main function to start the program."""
    course_data = load_course_data('B2A7.txt')
    branch = 'B2A7'  # For computer science branch
    generate_timetable(branch, course_data)

if __name__ == "__main__":
    main()