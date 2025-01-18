from file_handler import FileHandler
from timetable_generator import TimetableGenerator

def main():
    course_data = FileHandler.load_course_data('result.txt')
    all_branch_data = {}
    
    generator = TimetableGenerator()
    
    for branch in course_data:
        branch_timetable = generator.generate_timetable(branch, course_data)
        for semester, semester_data in branch_timetable.items():
            if semester not in all_branch_data:
                all_branch_data[semester] = {}
            all_branch_data[semester].update(semester_data)

    FileHandler.save_timetable_to_excel(all_branch_data, generator.state.all_course_assignments)

if __name__ == "__main__":
    main()