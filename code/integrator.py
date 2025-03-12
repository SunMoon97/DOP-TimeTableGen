import pandas as pd
import json
from file_handler import FileHandler
from timetable_generator import TimetableGenerator
import numpy as np
from allocate_classrooms import ClassroomAllocator

class IntegratedScheduler:
    def __init__(self, max_attempts=50):
        self.max_attempts = max_attempts
        self.best_allocation = None
        self.best_unallocated_count = float('inf')
        
    def generate_and_allocate(self, course_data_file, erp_file, rooms_file):
        """
        Generate timetable and allocate rooms with retry mechanism
        """
        for attempt in range(self.max_attempts):
            print(f"\nAttempt {attempt + 1}/{self.max_attempts}")
            
            # Generate timetable
            timetable_file = self._generate_timetable(course_data_file)
            
            # Try to allocate rooms
            allocator = ClassroomAllocator(timetable_file, erp_file, rooms_file)
            allocation_df, schedule_df, room_specific_timetables = allocator.allocate_classrooms()
            
            # Count unallocated courses
            unallocated_count = len(allocation_df[allocation_df['Room'] == "No Room Available"])
            print(f"Unallocated courses: {unallocated_count}")
            
            # Update best result if this attempt is better
            if unallocated_count < self.best_unallocated_count:
                self.best_unallocated_count = unallocated_count
                self.best_allocation = (allocation_df, schedule_df, room_specific_timetables)
                
                # If perfect allocation is found, stop
                if unallocated_count == 0:
                    print("\nPerfect allocation found!")
                    break
            
            if attempt == self.max_attempts - 1:
                print("\nMaximum attempts reached. Using best allocation found.")
        
        # Save best result
        if self.best_allocation:
            allocator.save_to_excel(
                self.best_allocation[0],
                self.best_allocation[1],
                self.best_allocation[2],
                "final_classroom_allocation.xlsx"
            )
            
        return self.best_unallocated_count
    
    def _generate_timetable(self, course_data_file):
        """
        Generate a new timetable
        """
        course_data = FileHandler.load_course_data(course_data_file)
        all_branch_data = {}
        
        generator = TimetableGenerator()
        
        for branch in course_data:
            branch_timetable = generator.generate_timetable(branch, course_data)
            for semester, semester_data in branch_timetable.items():
                if semester not in all_branch_data:
                    all_branch_data[semester] = {}
                all_branch_data[semester].update(semester_data)

        output_file = "combined_timetable1.xlsx"
        FileHandler.save_timetable_to_excel(all_branch_data, generator.state.all_course_assignments)
        
        return output_file

def main():
    # Input files
    course_data_file = 'result.txt'
    erp_file = "sem 1 23-24 erp tt& reg.xlsx"
    rooms_file = "data.xlsx"
    
    # Create scheduler
    scheduler = IntegratedScheduler(max_attempts=10)
    
    # Run integrated scheduling
    unallocated_count = scheduler.generate_and_allocate(
        course_data_file, erp_file, rooms_file
    )
    
    print(f"\nFinal Results:")
    print(f"Best unallocated count: {unallocated_count}")
    print("Check final_classroom_allocation.xlsx for complete allocation details")

if __name__ == "__main__":
    main()