from config import TimetableConfig
from data_structure import TimetableState
from course_assigner import CourseAssigner

class TimetableGenerator:
    def __init__(self):
        self.state = TimetableState()
        self.assigner = CourseAssigner(self.state)

    def generate_timetable(self, branch, course_data):
        all_semester_data = {}

        for semester, courses in course_data[branch].items():
            print(f"\nAssigning courses for {branch} - {semester}...")
            self.state.timetable = {day: {slot: [] for slot in TimetableConfig.TIME_SLOTS} 
                                  for day in TimetableConfig.ORDERED_DAYS}
            self.state.visited = {day: {slot: False for slot in TimetableConfig.TIME_SLOTS} 
                                for day in TimetableConfig.ORDERED_DAYS}

            self.assigner.assign_courses(courses)
            
            if semester not in all_semester_data:
                all_semester_data[semester] = {}

            all_semester_data[semester][branch] = self.state.timetable.copy()

        return all_semester_data