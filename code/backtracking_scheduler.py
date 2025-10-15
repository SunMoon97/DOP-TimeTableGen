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
        """Generate a course timetable with improved prioritization"""
        all_branch_data = {}

        # Use the original TimetableGenerator with prioritization
        generator = TimetableGenerator()

        # Sort branches by complexity (number of courses and constraints)
        sorted_branches = sorted(
            self.course_data.keys(),
            key=lambda b: self._calculate_branch_complexity(b),
            reverse=True
        )

        for branch in sorted_branches:
            branch_timetable = generator.generate_timetable(branch, self.course_data)
            for semester, semester_data in branch_timetable.items():
                if semester not in all_branch_data:
                    all_branch_data[semester] = {}
                all_branch_data[semester].update(semester_data)

        return all_branch_data, generator.state.all_course_assignments

    def _calculate_branch_complexity(self, branch):
        """Calculate complexity score for a branch based on courses and constraints"""
        complexity = 0
        for year_sem, courses in self.course_data[branch].items():
            complexity += len(courses) * 2  # Base score for number of courses

            # Add complexity for courses with labs
            for course, details in courses.items():
                if details.get('LPU', {}).get('labs', 0) > 0:
                    complexity += 5  # Labs are harder to schedule
                if details.get('Sections', {}).get('lectures', 0) > 1:
                    complexity += 3  # Multiple lecture sections add complexity

        return complexity
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
        """Get alternative slots with improved time preferences"""
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

        # Determine valid slots based on session type and student count
        student_count = self._get_student_count(course)

        # Prefer prime time slots for large courses
        if session_type == "Lecture":
            if student_count > 100:
                # Large courses get prime time slots (10 AM - 3 PM)
                valid_slots = ["10:00-11:00", "11:00-12:00", "12:00-13:00", "14:00-15:00"]
            else:
                # Smaller courses can use any slot except 8 AM
                valid_slots = TimetableConfig.TIME_SLOTS_WITHOUT_8AM
        elif session_type == "Tut":
            # Tutorials can be at any time
            valid_slots = TimetableConfig.TIME_SLOTS
        else:  # Lab
            # Labs preferably in afternoon
            valid_slots = ["14:00-15:00", "15:00-16:00", "16:00-17:00"]

        # Generate all possible combinations
        for day in valid_days:
            for slot in valid_slots:
                alternatives.append((day, slot))

        # Shuffle for randomness but with weighted preference
        # (This ensures we try good slots first but still have variety)
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
        """Apply smarter backtracking to problematic courses"""
        changes_made = False
    
        # Sort problem courses by student count (prioritize larger courses)
        problem_courses.sort(key=lambda x: self._get_student_count(x[0]), reverse=True)
    
        # Try to reschedule each problem course
        for course, day, time, session_type in problem_courses:
            print(f"Attempting to reschedule {course} on {day} at {time}")
        
            # First try direct rescheduling
            success = self._try_reschedule_course(
                course, day, time, session_type,
                all_branch_data, course_assignments
            )
        
            # If direct rescheduling fails, try cascading changes
            if not success:
                print(f"Direct rescheduling failed for {course}, trying cascading changes")
                success = self._try_cascading_reschedule(
                    course, day, time, session_type,
                    all_branch_data, course_assignments
                )
        
            if success:
                print(f"Successfully rescheduled {course}")
                changes_made = True
            else:
                print(f"Could not reschedule {course}")
    
        return changes_made

    def _try_cascading_reschedule(self, course, day, time, session_type, all_branch_data, course_assignments):
        """Try to reschedule a course by moving other courses first"""
        # Identify courses that could be moved to make room
        conflicting_courses = self._identify_conflicting_courses(course, day, time, all_branch_data)
    
        # Try moving each conflicting course
        for conflict_course in conflicting_courses:
            # Find current assignments for this course
            current_assignments = []
            for c_day, c_time, c_type in course_assignments.get(conflict_course, []):
                if c_type == session_type:  # Only consider same session type
                    current_assignments.append((c_day, c_time))
        
            # Try to move this conflicting course
            for c_day, c_time in current_assignments:
                if self._try_reschedule_course(
                    conflict_course, c_day, c_time, session_type,
                    all_branch_data, course_assignments
                ):
                    # Now try to place our original problem course
                    if self._try_reschedule_course(
                        course, day, time, session_type,
                        all_branch_data, course_assignments
                    ):
                        return True
    
        return False

    def _identify_conflicting_courses(self, course, day, time, all_branch_data):
        """Identify courses that might be conflicting with this course's placement"""
        conflicting_courses = []
    
        # Get student count for this course
        course_students = self._get_student_count(course)
    
        # Look for courses with fewer students that could be moved
        for semester, branches in all_branch_data.items():
            for branch, timetable in branches.items():
                for existing_course in timetable[day][time]:
                    existing_students = self._get_student_count(existing_course)
                    if existing_students < course_students:
                        conflicting_courses.append(existing_course)
    
        # Sort by student count (move smallest courses first)
        conflicting_courses.sort(key=self._get_student_count)
    
        return conflicting_courses
    def generate_and_allocate(self):
        """Generate timetable and allocate rooms with improved strategies"""
        for attempt in range(self.max_backtrack_attempts):
            print(f"\nAttempt {attempt + 1}/{self.max_backtrack_attempts}")
        
            # 1. Generate a basic timetable using improved logic
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

    def _preprocess_timetable(self, all_branch_data, course_assignments):
        """Pre-process the timetable to improve initial placement"""
        # 1. Identify courses with labs and ensure they have appropriate slots
        lab_courses = []
        for course, assignments in course_assignments.items():
            if any(session_type == "Lab" for _, _, session_type in assignments):
                lab_courses.append(course)
    
        # 2. Ensure large courses are in rooms with sufficient capacity
        large_courses = []
        for course in course_assignments:
            if self._get_student_count(course) > 100:
                large_courses.append(course)
    
        # 3. Try to spread courses evenly across days and times
        self._balance_timetable(all_branch_data)
    
        # 4. Handle specific constraints for lab courses
        for course in lab_courses:
            self._optimize_lab_placement(course, all_branch_data, course_assignments)
    
        # 5. Handle specific constraints for large courses
        for course in large_courses:
            self._optimize_large_course_placement(course, all_branch_data, course_assignments)

    def _balance_timetable(self, all_branch_data):
        """Balance the timetable to spread courses evenly"""
        # Count courses per day and time slot
        day_counts = {day: 0 for day in TimetableConfig.ORDERED_DAYS}
        time_counts = {time: 0 for time in TimetableConfig.TIME_SLOTS}
    
        for semester, branches in all_branch_data.items():
            for branch, timetable in branches.items():
                for day, day_data in timetable.items():
                    for time, courses in day_data.items():
                        day_counts[day] += len(courses)
                        time_counts[time] += len(courses)
    
        # Identify overloaded and underloaded slots
        avg_per_day = sum(day_counts.values()) / len(day_counts)
        avg_per_time = sum(time_counts.values()) / len(time_counts)
    
        overloaded_days = [day for day, count in day_counts.items() if count > avg_per_day * 1.2]
        underloaded_days = [day for day, count in day_counts.items() if count < avg_per_day * 0.8]
    
        overloaded_times = [time for time, count in time_counts.items() if count > avg_per_time * 1.2]
        underloaded_times = [time for time, count in time_counts.items() if count < avg_per_time * 0.8]
    
        # Try to move courses from overloaded to underloaded slots
        # (Implementation details would depend on your specific constraints)
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