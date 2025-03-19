import pandas as pd
import numpy as np
import random
import os
from timetable_generator import TimetableGenerator
from allocate_classrooms import ClassroomAllocator
from file_handler import FileHandler
from config import TimetableConfig
from data_structure import TimetableState
from course_assigner import CourseAssigner
from course_utils import CourseUtils

class BacktrackingScheduler:
    def __init__(self, course_data_file, erp_file, rooms_file, max_backtrack_attempts=50):
        """
        Initialize the backtracking scheduler
        
        Args:
            course_data_file (str): Path to the course data file
            erp_file (str): Path to the ERP data file
            rooms_file (str): Path to the rooms data file
            max_backtrack_attempts (int): Maximum number of backtracking attempts
        """
        self.course_data_file = course_data_file
        self.erp_file = erp_file
        self.rooms_file = rooms_file
        self.max_backtrack_attempts = max_backtrack_attempts
        
        # Load course data
        self.course_data = FileHandler.load_course_data(course_data_file)
        
        # Initialize the best solution variables
        self.best_solution = None
        self.best_unallocated_count = float('inf')
        self.best_timetable = None
        self.best_course_assignments = None
        
        # Preload room and enrollment data
        self._preload_room_and_enrollment_data()
    
    def _preload_room_and_enrollment_data(self):
        """Preload room and enrollment data for quicker access during backtracking"""
        # Load ERP data for enrollments
        erp_data = pd.read_excel(self.erp_file, sheet_name="Sheet1", header=1)
        erp_data["Course"] = (erp_data["Subject"].astype(str).str.strip() + 
                            ' ' + erp_data["Catalog"].astype(str).str.strip()).str.replace('  ', ' ')
        self.course_enrollment = erp_data[["Course", "No. Of Students"]]
        self.course_enrollment = self.course_enrollment.sort_values(
            by="No. Of Students", ascending=False
        )
        
        # Load room capacity data
        rooms_data = pd.read_excel(self.rooms_file, sheet_name="3.rooms")
        self.room_capacity = rooms_data[["Room", "Seating Capacity"]].copy()
        self.room_capacity["Seating Capacity"] = pd.to_numeric(
            self.room_capacity["Seating Capacity"], errors='coerce'
        )
        self.room_capacity = self.room_capacity.sort_values(
            by="Seating Capacity", ascending=False
        )
        
        # Create dictionaries for faster lookup
        self.room_capacity_dict = dict(zip(self.room_capacity["Room"], self.room_capacity["Seating Capacity"]))
        self.course_enrollment_dict = dict(zip(self.course_enrollment["Course"], self.course_enrollment["No. Of Students"]))
    
    def _generate_timetable(self):
        """
        Generate a course timetable using the original TimetableGenerator logic
        
        Returns:
            tuple: (all_branch_data, all_course_assignments) - The generated timetable and assignments
        """
        all_branch_data = {}
        
        # Use the original TimetableGenerator
        generator = TimetableGenerator()
        
        for branch in self.course_data:
            branch_timetable = generator.generate_timetable(branch, self.course_data)
            for semester, semester_data in branch_timetable.items():
                if semester not in all_branch_data:
                    all_branch_data[semester] = {}
                all_branch_data[semester].update(semester_data)
        
        return all_branch_data, generator.state.all_course_assignments
    
    def _get_student_count(self, course):
        """Get student count for a given course"""
        course_base = course.split('_')[0] if '_' in course else course
        course_base = course_base.split('-')[0] if '-' in course_base else course_base
        return self.course_enrollment_dict.get(course_base, 0)
    
    def _save_timetable(self, all_branch_data, all_course_assignments):
        """Save timetable to Excel file"""
        output_file = "combined_timetable_backtracking.xlsx"
        # Create the full path to the file
        full_path = os.path.join(os.getcwd(), output_file)
        FileHandler.save_timetable_to_excel(all_branch_data, all_course_assignments, output_file)
        
        # Verify file was created
        if not os.path.exists(full_path):
            print(f"Warning: File {full_path} was not created properly")
            # Use a fallback file
            output_file = "combined_timetable1.xlsx"
            
        return output_file
    
    def _evaluate_solution(self, timetable_file):
        """
        Evaluate a solution by allocating classrooms and counting unallocated courses
        
        Args:
            timetable_file (str): Path to the timetable Excel file
            
        Returns:
            tuple: (unallocated_count, allocation_df, schedule_df, room_specific_timetables)
        """
        try:
            allocator = ClassroomAllocator(timetable_file, self.erp_file, self.rooms_file)
            allocation_df, schedule_df, room_specific_timetables = allocator.allocate_classrooms()
            
            # Count unallocated courses
            unallocated_count = len(allocation_df[allocation_df['Room'] == "No Room Available"])
            print(f"Unallocated courses: {unallocated_count}")
            
            return unallocated_count, allocation_df, schedule_df, room_specific_timetables
        except FileNotFoundError:
            print(f"Error: Could not open timetable file {timetable_file}.")
            print("Using fallback timetable file 'combined_timetable1.xlsx'")
            
            # Try with a different file
            allocator = ClassroomAllocator("combined_timetable1.xlsx", self.erp_file, self.rooms_file)
            allocation_df, schedule_df, room_specific_timetables = allocator.allocate_classrooms()
            
            # Count unallocated courses
            unallocated_count = len(allocation_df[allocation_df['Room'] == "No Room Available"])
            print(f"Unallocated courses: {unallocated_count}")
            
            return unallocated_count, allocation_df, schedule_df, room_specific_timetables
    
    def _identify_problem_courses(self, allocation_df):
        """
        Identify courses that couldn't be allocated rooms
        
        Args:
            allocation_df (DataFrame): Course to room allocation dataframe
            
        Returns:
            list: List of problem courses with details (course, day, time, type)
        """
        problem_courses = []
        unallocated = allocation_df[allocation_df['Room'] == "No Room Available"]
        
        for _, row in unallocated.iterrows():
            problem_courses.append((
                row['Course'],
                row['Day'],
                row['Time'],
                row['Type']
            ))
        
        return problem_courses
    
    def _get_alternative_slots(self, course, session_type):
        """
        Get alternative slots for a course based on session type
        
        Args:
            course (str): Course name
            session_type (str): Session type (Lecture/Tutorial/Lab)
            
        Returns:
            list: List of (day, slot) tuples for alternative placements
        """
        alternatives = []
        
        # Determine valid days
        valid_days = TimetableConfig.ORDERED_DAYS
        
        # For lectures, determine whether course is a MWF or TT course
        if session_type == "Lecture":
            # Simple heuristic - check if the course contains MWF or TT in existing assignments
            if any("Monday" in d or "Wednesday" in d or "Friday" in d for d, _, _ in self.best_course_assignments.get(course, [])):
                valid_days = TimetableConfig.DAYS_MWF
            elif any("Tuesday" in d or "Thursday" in d for d, _, _ in self.best_course_assignments.get(course, [])):
                valid_days = TimetableConfig.DAYS_TT
        
        # Determine valid slots based on session type
        if session_type == "Lecture":
            valid_slots = TimetableConfig.TIME_SLOTS_WITHOUT_8AM
        elif session_type == "Tut":
            valid_slots = TimetableConfig.TIME_SLOTS  # Tutorials can be at 8 AM
        else:  # Lab
            valid_slots = TimetableConfig.TIME_SLOTS_WITHOUT_8AM
        
        # Generate all possible combinations
        for day in valid_days:
            for slot in valid_slots:
                alternatives.append((day, slot))
        
        # Shuffle for randomness
        random.shuffle(alternatives)
        
        return alternatives
    
    def _try_reschedule_course(self, course, day, time, session_type, all_branch_data, course_assignments):
        """
        Try to reschedule a course to a different time slot
        
        Args:
            course (str): Course name
            day (str): Current day
            time (str): Current time
            session_type (str): Session type
            all_branch_data (dict): Current timetable
            course_assignments (dict): Current course assignments
            
        Returns:
            bool: Whether rescheduling was successful
        """
        # Get alternative slots
        alternatives = self._get_alternative_slots(course, session_type)
        
        course_base = course.split('_')[0] if '_' in course else course
        
        # First, remove the course from its current slot
        for semester, branches in all_branch_data.items():
            for branch, timetable in branches.items():
                if course in timetable[day][time]:
                    timetable[day][time].remove(course)
                    break
        
        # Remove from course assignments
        if course in course_assignments:
            course_assignments[course] = [
                (d, t, st) for d, t, st in course_assignments[course]
                if not (d == day and t == time and st == session_type)
            ]
        
        # Try to place the course in each alternative slot
        for alt_day, alt_time in alternatives:
            # Check if the slot is valid for this course
            can_assign = True
            
            # Check if course with same base is already in this slot
            for semester, branches in all_branch_data.items():
                for branch, timetable in branches.items():
                    for existing_course in timetable[alt_day][alt_time]:
                        existing_base = existing_course.split('_')[0] if '_' in existing_course else existing_course
                        if existing_base == course_base:
                            can_assign = False
                            break
                    if not can_assign:
                        break
                if not can_assign:
                    break
            
            if can_assign:
                # Place the course in this new slot
                for semester, branches in all_branch_data.items():
                    for branch, timetable in branches.items():
                        if any(course in slot_courses for slot_courses in timetable[day].values()):
                            timetable[alt_day][alt_time].append(course)
                            # Add to course assignments
                            if course not in course_assignments:
                                course_assignments[course] = []
                            course_assignments[course].append((alt_day, alt_time, session_type))
                            return True
        
        # If we get here, we couldn't reschedule
        return False
    
    def _backtrack_problematic_courses(self, problem_courses, all_branch_data, course_assignments):
        """
        Apply backtracking to problematic courses
        
        Args:
            problem_courses (list): List of problem courses with details
            all_branch_data (dict): Current timetable
            course_assignments (dict): Current course assignments
            
        Returns:
            bool: Whether any changes were made
        """
        changes_made = False
        
        for course, day, time, session_type in problem_courses:
            print(f"Attempting to reschedule {course} on {day} at {time}")
            success = self._try_reschedule_course(
                course, day, time, session_type,
                all_branch_data, course_assignments
            )
            if success:
                print(f"Successfully rescheduled {course}")
                changes_made = True
            else:
                print(f"Could not reschedule {course}")
        
        return changes_made

    def generate_and_allocate(self):
        """
        Generate timetable and allocate rooms with backtracking
        
        Returns:
            int: Number of unallocated courses in the best solution
        """
        for attempt in range(self.max_backtrack_attempts):
            print(f"\nAttempt {attempt + 1}/{self.max_backtrack_attempts}")
            
            # 1. Generate a basic timetable using the original logic
            all_branch_data, course_assignments = self._generate_timetable()
            
            # Save the timetable
            try:
                timetable_file = self._save_timetable(all_branch_data, course_assignments)
                
                # 2. Try to allocate classrooms
                unallocated_count, allocation_df, schedule_df, room_specific_timetables = self._evaluate_solution(timetable_file)
                
                # Update best solution if this attempt is better
                if unallocated_count < self.best_unallocated_count:
                    self.best_unallocated_count = unallocated_count
                    self.best_solution = (allocation_df, schedule_df, room_specific_timetables)
                    self.best_timetable = all_branch_data
                    self.best_course_assignments = course_assignments.copy()
                    
                    # If perfect allocation is found, stop
                    if unallocated_count == 0:
                        print("\nPerfect allocation found!")
                        break
                
                # 3. If there are unallocated courses, try backtracking
                if unallocated_count > 0 and unallocated_count < 20:  # Only backtrack if not too many problems
                    print(f"Trying backtracking for {unallocated_count} unallocated courses")
                    
                    # Identify problematic courses
                    problem_courses = self._identify_problem_courses(allocation_df)
                    
                    # Create a copy of the current state to modify
                    backtrack_branch_data = {
                        semester: {
                            branch: {
                                day: {slot: list(courses) for slot, courses in day_data.items()}
                                for day, day_data in timetable.items()
                            }
                            for branch, timetable in branches.items()
                        }
                        for semester, branches in all_branch_data.items()
                    }
                    backtrack_assignments = {
                        course: list(assignments) 
                        for course, assignments in course_assignments.items()
                    }
                    
                    # Apply backtracking to problematic courses
                    changes_made = self._backtrack_problematic_courses(
                        problem_courses, backtrack_branch_data, backtrack_assignments
                    )
                    
                    if changes_made:
                        # Save the backtracked timetable
                        backtrack_file = self._save_timetable(backtrack_branch_data, backtrack_assignments)
                        
                        # Evaluate the backtracked solution
                        bt_unallocated, bt_allocation, bt_schedule, bt_rooms = self._evaluate_solution(backtrack_file)
                        
                        print(f"After backtracking: {bt_unallocated} unallocated courses")
                        
                        # Update best solution if backtracking improved things
                        if bt_unallocated < unallocated_count:
                            self.best_unallocated_count = bt_unallocated
                            self.best_solution = (bt_allocation, bt_schedule, bt_rooms)
                            self.best_timetable = backtrack_branch_data
                            self.best_course_assignments = backtrack_assignments
                            
                            # If perfect allocation found through backtracking, stop
                            if bt_unallocated == 0:
                                print("\nPerfect allocation found through backtracking!")
                                break
            except Exception as e:
                print(f"Error in attempt {attempt + 1}: {str(e)}")
                print("Continuing to next attempt...")
                
            if attempt == self.max_backtrack_attempts - 1:
                print("\nMaximum attempts reached. Using best allocation found.")
        
        # Save best result
        if self.best_solution:
            try:
                # First save the best timetable
                best_timetable_file = self._save_timetable(self.best_timetable, self.best_course_assignments)
                
                # Then save the classroom allocation
                allocator = ClassroomAllocator(best_timetable_file, self.erp_file, self.rooms_file)
                allocator.save_to_excel(
                    self.best_solution[0],
                    self.best_solution[1],
                    self.best_solution[2],
                    "final_classroom_allocation_backtracking.xlsx"
                )
                print("Final solution saved to final_classroom_allocation_backtracking.xlsx")
            except Exception as e:
                print(f"Error saving final solution: {str(e)}")
                print("You may need to manually save the solution by re-running the program.")
            
        return self.best_unallocated_count

def main():
    # Input files
    course_data_file = 'result.txt'
    erp_file = "sem 1 23-24 erp tt& reg.xlsx"
    rooms_file = "data.xlsx"
    
    # Create scheduler
    scheduler = BacktrackingScheduler(
        course_data_file, erp_file, rooms_file, max_backtrack_attempts=10
    )
    
    # Run integrated scheduling with backtracking
    unallocated_count = scheduler.generate_and_allocate()
    
    print(f"\nFinal Results:")
    print(f"Best unallocated count: {unallocated_count}")
    print("Check final_classroom_allocation_backtracking.xlsx for complete allocation details")

if __name__ == "__main__":
    main() 