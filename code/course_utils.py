class CourseUtils:
    @staticmethod
    def generate_component_name(course, component, index):
        return f"{course}_{component}_{index}"

    @staticmethod
    def is_valid_slot(slot, is_tutorial=False):
        if is_tutorial:
            return slot != "8:00"
        return slot != "1:00" and slot != "8:00"

    @staticmethod
    def is_valid_day_for_course(course, day, assigned_slots):
        return all(assigned_day != day for assigned_day, _ in assigned_slots)

    @staticmethod
    def not_present_here(day, slot, course, timetable):
        if not course[-1].isdigit() and timetable[day][slot].count(course) > 0:
            return False
            
        for all_slots in timetable[day].values():
            for c in all_slots:
                if c[10:] == course[10:]:
                    return False
                
        for c in timetable[day][slot]:
            end = c[-1]
            if not end.isdigit():
                return False
            if end == course[-1]:
                return False
            if c[10:] == course[10:]:
                return False
            
        return True