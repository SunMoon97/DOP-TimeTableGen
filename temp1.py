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
        return slot != "8:00"
    else:
        return slot != "1:00" and slot != "8:00"

def is_valid_day_for_course(course, day, assigned_slots):
    """Check if a day is valid for assigning a new session for a course."""
    return all(assigned_day != day for assigned_day, _ in assigned_slots)


def not_present_here(day,slot,course):
    """Check if the course is already assigned in the same day and slot."""
    #extract all courses in the day
    if not course[-1].isdigit() and timetable[day][slot].count(course) > 0:
            return False
    for all_slots in timetable[day]:
        for c in all_slots:
            if c[10:] == course[10:]:
                return False
            
    for c in timetable[day][slot]:
    # if the section of the c course of component same as course and the number of section is 1 then return false
        
        end = c[-1]
        # Check if the last character 'end' is not a digit
        if not end.isdigit():
            return False
        with open('course_info.txt','w') as f:
            f.write(f"End character of course {c}: {end}\n")
            f.write(f"Last character of course {course}: {course[-1]}\n")
        if  end == course[-1]:
            return False
        if c[10:] == course[10:]:
            return False
        
    return True
        
        
    
def assign_session(course, session_type, count, series, assigned_slots):
    """Assign a session (lecture, tutorial, or lab) to the timetable."""
    for _ in range(count):
        while True:
            day = random.choice([d for d in series if is_valid_day_for_course(course, d, assigned_slots)])
            valid_slots = [slot for slot in time_slots if not_present_here(day,slot,course) and is_valid_slot(slot, session_type == "Tut")]
            if valid_slots:
                slot = "8:00"
                
                timetable[day][slot].append(f"{course}")
                assigned_slots.append((day, slot))
                # Add to all_course_assignments
                all_course_assignments.setdefault(course, []).append((day, slot, session_type))
                break

def assign_lab(course, labs, assigned_slots):
    """Assign lab sessions for a course."""
    attempts = 0
    while attempts < 10 and labs > 0:
        day = random.choice(ordered_days)
        # Check for a continuous block of slots for labs
        for start_index in range(len(time_slots) - labs + 1):
            #use function of not_present_here instead of visited
            if all(is_valid_slot(time_slots[start_index + j]) and not visited[day][time_slots[start_index + j]] for j in range(labs)):
                for j in range(labs):
                    timetable[day][time_slots[start_index + j]].append(f"{course}")
                    visited[day][time_slots[start_index + j]] = True  # Mark as visited
                    # Add to all_course_assignments
                    all_course_assignments.setdefault(course, []).append((day, time_slots[start_index + j], "Lab"))
                assigned_slots.append((day, time_slots[start_index]))  # Record the starting time slot
                labs -= len(range(labs))  # Deduct the assigned labs
                break
        attempts += 1

def assign_course_to_timetable(course, lectures, tutorials, lab_hours, lecture_series, tutorial_series):
    """Assign all sessions for a course to the timetable."""
    assigned_slots = []
    remaining_lectures = lectures
    remaining_tutorials = tutorials

    # Check if course is already assigned
    if course in all_course_assignments:
        for day, slot, session_type in all_course_assignments[course]:
            if session_type == "Lecture" and remaining_lectures!=0:
                print(f"{course} already assigned on {day} at {slot}")
                remaining_lectures -= 1
                print(f"Remaining lectures: {remaining_lectures}")
            elif session_type == "Tut":
                remaining_tutorials -= 1
            elif session_type == "Lab":
                lab_hours = 0
            timetable[day][slot].append(f"{course}")
            visited[day][slot] = True
            assigned_slots.append((day, slot))
    # Assign remaining lectures
    for day in lecture_series:
        if remaining_lectures == 0:
            print("No more lectures for " + course)
            break
        available_slots=["9:00", "10:00", "4:00", "5:00"] if day in days_mwf else ["11:00", "12:00", "2:00", "3:00"]
        valid_slots = [slot for slot in available_slots if not_present_here(day,slot,course) and is_valid_slot(slot)]
        if valid_slots:
            lecture_time = random.choice(valid_slots)
            visited[day][lecture_time] = True  # Mark as visited
            timetable[day][lecture_time].append(f"{course}")
            assigned_slots.append((day, lecture_time))
            all_course_assignments.setdefault(course, []).append((day, lecture_time, "Lecture"))
            remaining_lectures -= 1  # Decrement lecture counter

            # Assign the same lecture slot to corresponding series days
            for series_day in lecture_series[1:]:
                if is_valid_day_for_course(course, series_day, assigned_slots) and not_present_here(series_day,lecture_time,course):
                    timetable[series_day][lecture_time].append(f"{course}")
                    visited[series_day][lecture_time] = True  # Mark as visited
                    assigned_slots.append((series_day, lecture_time))
                    all_course_assignments.setdefault(course, []).append((series_day, lecture_time, "Lecture"))
                    remaining_lectures -= 1  # Decrement lecture counter

    # Assign remaining tutorials
    while remaining_tutorials > 0:
        assign_session(course, "Tut", 1, tutorial_series, assigned_slots)  # Assign one tutorial at a time
        remaining_tutorials -= 1  # Decrement tutorial counter

    # Assign lab (labs are only one per course, with varying hours)
    if lab_hours > 0:
        assign_lab(course, lab_hours, assigned_slots)

    return remaining_lectures, remaining_tutorials

#  def assign_courses(courses):
    """Assign all courses to the timetable based on their series."""
    unassigned_courses = {}

    # Assign Mon-Wed-Fri and Tue-Thu courses
    half_courses = len(courses) // 2
    course_list = list(courses.items())
    mwf_courses = course_list[:half_courses]
    tt_courses = course_list[half_courses:]

    for course, load in mwf_courses:
        
        lectures_left, tutorials_left = assign_course_to_timetable(
            course, load['LPU']['lectures'], load['LPU']['tutorials'], load['LPU']['lab_hours'], days_mwf, days_tt
        )

        # Check if any unassigned sessions are left
        if lectures_left > 0 or tutorials_left > 0:
            unassigned_courses[course] = {
                'lectures': lectures_left,
                'tutorials': tutorials_left,
            }
            
    for course, load in tt_courses:
        # Temporarily add a random MWF day to TT series for better distribution
        random_mwf_day = random.choice(days_mwf)
        days_tt.append(random_mwf_day)
        
        lectures_left, tutorials_left = assign_course_to_timetable(
            course, load['LPU']['lectures'], load['LPU']['tutorials'], load['LPU']['lab_hours'], days_tt, days_mwf
        )
        
        days_tt.remove(random_mwf_day)

        # Check if any unassigned sessions are left
        if lectures_left > 0 or tutorials_left > 0:
            unassigned_courses[course] = {
                'lectures': lectures_left,
                'tutorials': tutorials_left,
            }
            

    # Try to reassign unassigned sessions if needed
    reassign_unassigned_courses(unassigned_courses)
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
        lecture_parallel = load['No of sections parallel'][0]  # For lectures
        tutorial_parallel = load['No of sections parallel'][1]  # For tutorials
        lab_parallel = load['No of sections parallel'][2]  # For labs

        lecture_sections = load['Sections'].get('lectures', 0)
        for i in lecture_parallel:
            course_variant = generate_component_name(course, "Lecture", i[0])
            if lecture_sections == 1:
                course_variant = course
            lectures_left, _ = assign_course_to_timetable(
                course_variant, load['LPU']['lectures'], 0, 0, days_mwf, days_tt
            )
            if lectures_left > 0:
                unassigned_courses[course_variant] = {'lectures': lectures_left}

            # Get the day and time slot for the current iteration 'i'
            day, slot = next(((d, s) for d in ordered_days for s in time_slots if course_variant in timetable[d][s]), (None, None))
            if day and slot:
                # Assign all other 'j' to the same slot
                if isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = generate_component_name(course, "Lecture", j)
                        timetable[day][slot].append(course_variant_j)
                        
                        all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Lecture"))
                    # Assign the same lecture slot to corresponding series days
                        for series_day in days_mwf[1:]:
                            if is_valid_day_for_course(course, series_day, [(day, slot)]) and not_present_here(series_day, slot, course_variant_j):
                                timetable[series_day][slot].append(course_variant_j)
                                
                                all_course_assignments.setdefault(course_variant_j, []).append((series_day, slot, "Lecture"))
            
            
        

        # Assign tutorials
        tutorial_sections = load['Sections'].get('tutorials', 0)
        for i in tutorial_parallel:
            
            course_variant = generate_component_name(course, "Tutorial", i[0])
            
            _, tutorials_left = assign_course_to_timetable(
                course_variant, 0, load['LPU']['tutorials'], 0, days_mwf, days_tt
            )
            if tutorials_left > 0:
                unassigned_courses[course_variant] = {'tutorials': tutorials_left}
            day, slot = next(((d, s) for d in ordered_days for s in time_slots if course_variant in timetable[d][s]), (None, None))
            if day and slot:
                if isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = generate_component_name(course, "Tutorial", j)
                        timetable[day][slot].append(course_variant_j)
                        visited[day][slot] = True
                        all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Tutorial"))

        # Assign labs
        lab_sections = load['Sections'].get('labs', 0)
        for i in range(1, lab_sections + 1):
            course_variant = generate_component_name(course, "Lab", i)
            _, _ = assign_course_to_timetable(
                course_variant, 0, 0, load['LPU']['lab_hours'], days_mwf, days_tt
            )

    # Assign TT courses
    #for lectures in TT courses use the same logic as in MWF
    for course, load in tt_courses:
        
        random_mwf_day = random.choice(days_mwf)
        days_tt.append(random_mwf_day)
        lecture_parallel = load['No of sections parallel'][0]  # For lectures
        tutorial_parallel = load['No of sections parallel'][1]  # For tutorials
        lab_parallel = load['No of sections parallel'][2]  # For lab
        lecture_sections = load['Sections'].get('lectures', 0)
        for i in lecture_parallel:
            course_variant = generate_component_name(course, "Lecture", i[0])
            if lecture_sections == 1:
                course_variant = course
            lectures_left, _ = assign_course_to_timetable(
                course_variant, load['LPU']['lectures'], 0, 0, days_tt, days_mwf
            )
            if lectures_left > 0:
                unassigned_courses[course_variant] = {'lectures': lectures_left}

            # Get the day and time slot for the current iteration 'i'
            day, slot = next(((d, s) for d in ordered_days for s in time_slots if course_variant in timetable[d][s]), (None, None))
            
            # If a valid day and slot were found, assign all other 'j' to the same slot
            if day and slot:
                # Assign all other 'j' to the same slot
                if isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = generate_component_name(course, "Lecture", j)
                        timetable[day][slot].append(course_variant_j)
                        visited[day][slot] = True
                        all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Lecture"))
                    # Assign the same lecture slot to corresponding series days
                        for series_day in days_tt[1:]:
                            if is_valid_day_for_course(course, series_day, [(day, slot)]) and not_present_here(series_day, slot, course_variant_j):
                                timetable[series_day][slot].append(course_variant_j)
                                visited[series_day][slot] = True
                                all_course_assignments.setdefault(course_variant_j, []).append((series_day, slot, "Lecture"))
        # Assign tutorials
        # use same logic o for tutorials
        tutorial_parallel = load['No of sections parallel'][1]
        tutorial_sections = load['Sections'].get('tutorials', 0)
        for i in tutorial_parallel:
            course_variant = generate_component_name(course, "Tutorial", i[0])
            _, tutorials_left = assign_course_to_timetable(
                course_variant, 0, load['LPU']['tutorials'], 0, days_tt, days_mwf
            )
            if tutorials_left > 0:
                unassigned_courses[course_variant] = {'tutorials': tutorials_left}
                
            day, slot = next(((d, s) for d in ordered_days for s in time_slots if course_variant in timetable[d][s]), (None, None))
            
            if day and slot:
                if isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = generate_component_name(course, "Tutorial", j)
                        timetable[day][slot].append(course_variant_j)
                        visited[day][slot] = True
                        all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Tutorial"))
                        

        # Assign labs
        lab_sections = load['Sections'].get('labs', 0)
        for i in range(1, lab_sections + 1):
            course_variant = generate_component_name(course, "Lab", i)
            _, _ = assign_course_to_timetable(
                course_variant, 0, 0, load['LPU']['lab_hours'], days_tt, days_mwf
            )
        days_tt.remove(random_mwf_day)

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
                               if not_present_here(day, slot, course) and is_valid_slot(slot)]
            if available_slots:
                day, slot = random.choice(available_slots)
                timetable[day][slot].append(f"{course}")
                visited[day][slot] = True
                all_course_assignments.setdefault(course, []).append((day, slot, "Lecture"))
                lectures_left -= 1
            else:
                break  # No more available slots

        # Try to reassign remaining tutorials
        while tutorials_left > 0:
            available_slots = [(day, slot) for day in ordered_days for slot in time_slots if
                               not_present_here(day, slot, course) and is_valid_slot(slot, True)]
            if available_slots:
                day, slot = random.choice(available_slots)
                timetable[day][slot].append(f"{course}")
                visited[day][slot] = True
                all_course_assignments.setdefault(course, []).append((day, slot, "Tut"))
                tutorials_left -= 1
            else:
                break  # No more available slots


def generate_timetable(branch, course_data):
    """Generate the timetable for all semesters of a specific branch."""
    all_semester_data = {}

    for semester, courses in course_data[branch].items():
        print(f"\nAssigning courses for {branch} - {semester}...")
        global timetable
        timetable = {day: {slot: [] for slot in time_slots} for day in ordered_days}
        
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