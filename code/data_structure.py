from config import TimetableConfig
class TimetableState:
    def __init__(self):
        self.visited = {day: {slot: False for slot in TimetableConfig.TIME_SLOTS} 
                       for day in TimetableConfig.ORDERED_DAYS}
        self.timetable = {day: {slot: [] for slot in TimetableConfig.TIME_SLOTS} 
                         for day in TimetableConfig.ORDERED_DAYS}
        self.all_course_assignments = {}