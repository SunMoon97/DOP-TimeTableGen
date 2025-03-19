from file_handler import FileHandler
from timetable_generator import TimetableGenerator
from integrator import IntegratedScheduler
from backtracking_scheduler import BacktrackingScheduler
import argparse

def main():
    # Set up command line arguments
    parser = argparse.ArgumentParser(description='Timetable Generator with Classroom Allocation')
    parser.add_argument('--method', choices=['basic', 'integrated', 'backtracking'], default='backtracking',
                        help='Method to use: basic (no room allocation), integrated (random trials), or backtracking')
    parser.add_argument('--attempts', type=int, default=10,
                        help='Number of attempts for integrated or backtracking methods')
    args = parser.parse_args()
    
    # Input files
    course_data_file = 'result.txt'
    erp_file = "sem 1 23-24 erp tt& reg.xlsx"
    rooms_file = "data.xlsx"
    
    if args.method == 'basic':
        # Original basic method (no classroom allocation)
        course_data = FileHandler.load_course_data(course_data_file)
        all_branch_data = {}
        
        generator = TimetableGenerator()
        
        for branch in course_data:
            branch_timetable = generator.generate_timetable(branch, course_data)
            for semester, semester_data in branch_timetable.items():
                if semester not in all_branch_data:
                    all_branch_data[semester] = {}
                all_branch_data[semester].update(semester_data)

        FileHandler.save_timetable_to_excel(all_branch_data, generator.state.all_course_assignments)
        print("Basic timetable generation completed. Check combined_timetable.xlsx for results.")
        
    elif args.method == 'integrated':
        # Integrated scheduler with random trials
        scheduler = IntegratedScheduler(max_attempts=args.attempts)
        unallocated_count = scheduler.generate_and_allocate(
            course_data_file, erp_file, rooms_file
        )
        print(f"\nFinal Results (Integrated Method):")
        print(f"Best unallocated count: {unallocated_count}")
        print("Check final_classroom_allocation.xlsx for complete allocation details")
        
    elif args.method == 'backtracking':
        # Backtracking scheduler with classroom constraints
        scheduler = BacktrackingScheduler(
            course_data_file, erp_file, rooms_file, max_backtrack_attempts=args.attempts
        )
        unallocated_count = scheduler.generate_and_allocate()
        print(f"\nFinal Results (Backtracking Method):")
        print(f"Best unallocated count: {unallocated_count}")
        print("Check final_classroom_allocation_backtracking.xlsx for complete allocation details")

if __name__ == "__main__":
    main()