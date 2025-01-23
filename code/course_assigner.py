from config import TimetableConfig
from course_utils import CourseUtils
import random

class CourseAssigner:
    def __init__(self, state):
        self.state = state
        self.utils = CourseUtils()

    def assign_session(self, course, session_type, count, series, assigned_slots):
        for _ in range(count):
            while True:
                day = random.choice([d for d in series 
                                   if self.utils.is_valid_day_for_course(course, d, assigned_slots)])
                valid_slots = [slot for slot in TimetableConfig.TIME_SLOTS 
                             if self.utils.not_present_here(day, slot, course, self.state.timetable) 
                             and self.utils.is_valid_slot(slot, session_type == "Tut")]
                if valid_slots:
                    slot = "8:00" if session_type == "Tut" else random.choice(valid_slots)
                    self.state.timetable[day][slot].append(f"{course}")
                    assigned_slots.append((day, slot))
                    self.state.all_course_assignments.setdefault(course, []).append((day, slot, session_type))
                    break

    def assign_lab(self, course, labs, assigned_slots):
        attempts = 0
        while attempts < 10 and labs > 0:
            day = random.choice(TimetableConfig.ORDERED_DAYS)
            for start_index in range(len(TimetableConfig.TIME_SLOTS) - labs + 1):
                if all(self.utils.is_valid_slot(TimetableConfig.TIME_SLOTS[start_index + j]) 
                      and not self.state.visited[day][TimetableConfig.TIME_SLOTS[start_index + j]] 
                      for j in range(labs)):
                    for j in range(labs):
                        slot = TimetableConfig.TIME_SLOTS[start_index + j]
                        self.state.timetable[day][slot].append(f"{course}")
                        self.state.visited[day][slot] = True
                        self.state.all_course_assignments.setdefault(course, []).append((day, slot, "Lab"))
                    assigned_slots.append((day, TimetableConfig.TIME_SLOTS[start_index]))
                    labs -= labs
                    break
            attempts += 1

    def assign_course_to_timetable(self, course, lectures, tutorials, lab_hours, lecture_series, tutorial_series):
        assigned_slots = []
        remaining_lectures = lectures
        remaining_tutorials = tutorials

        if course in self.state.all_course_assignments:
            for day, slot, session_type in self.state.all_course_assignments[course]:
                if session_type == "Lecture" and remaining_lectures != 0:
                    print(f"{course} already assigned on {day} at {slot}")
                    remaining_lectures -= 1
                    print(f"Remaining lectures: {remaining_lectures}")
                elif session_type == "Tut":
                    remaining_tutorials -= 1
                elif session_type == "Lab":
                    lab_hours = 0
                self.state.timetable[day][slot].append(f"{course}")
                self.state.visited[day][slot] = True
                assigned_slots.append((day, slot))

        for day in lecture_series:
            if remaining_lectures == 0:
                print("No more lectures for " + course)
                break
            available_slots = ["9:00", "10:00", "11:00", "12:00", "2:00", "3:00", "4:00", "5:00"]
            valid_slots = [slot for slot in available_slots 
                          if self.utils.not_present_here(day, slot, course, self.state.timetable) 
                          and self.utils.is_valid_slot(slot)]
            if valid_slots:
                lecture_time = random.choice(valid_slots)
                self.state.visited[day][lecture_time] = True
                self.state.timetable[day][lecture_time].append(f"{course}")
                assigned_slots.append((day, lecture_time))
                self.state.all_course_assignments.setdefault(course, []).append((day, lecture_time, "Lecture"))
                remaining_lectures -= 1

                for series_day in lecture_series[1:]:
                    if (self.utils.is_valid_day_for_course(course, series_day, assigned_slots) 
                            and self.utils.not_present_here(series_day, lecture_time, course, self.state.timetable)):
                        self.state.timetable[series_day][lecture_time].append(f"{course}")
                        self.state.visited[series_day][lecture_time] = True
                        assigned_slots.append((series_day, lecture_time))
                        self.state.all_course_assignments.setdefault(course, []).append((series_day, lecture_time, "Lecture"))
                        remaining_lectures -= 1

        while remaining_tutorials > 0:
            self.assign_session(course, "Tut", 1, tutorial_series, assigned_slots)
            remaining_tutorials -= 1

        if lab_hours > 0:
            self.assign_lab(course, lab_hours, assigned_slots)

        return remaining_lectures, remaining_tutorials

    def assign_courses(self, courses):
        unassigned_courses = {}
        half_courses = len(courses) // 2
        course_list = list(courses.items())
        mwf_courses = course_list[:half_courses]
        tt_courses = course_list[half_courses:]

        # Assign MWF courses
        for course, load in mwf_courses:
            lecture_parallel = load['No of sections parallel'][0]
            tutorial_parallel = load['No of sections parallel'][1]
            lab_parallel = load['No of sections parallel'][2]

            lecture_sections = load['Sections'].get('lectures', 0)
            for i in lecture_parallel:
                if lecture_sections == 1:
                    course_variant = f"{course}_Lecture"
                else:
                    course_variant = self.utils.generate_component_name(course, "Lecture", i[0])
                lectures_left, _ = self.assign_course_to_timetable(
                    course_variant, load['LPU']['lectures'], 0, 0, TimetableConfig.DAYS_MWF, TimetableConfig.DAYS_TT
                )
                if lectures_left > 0:
                    unassigned_courses[course_variant] = {'lectures': lectures_left}

                day, slot = next(((d, s) for d in TimetableConfig.ORDERED_DAYS 
                                for s in TimetableConfig.TIME_SLOTS 
                                if course_variant in self.state.timetable[d][s]), (None, None))
                
                if day and slot and isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = self.utils.generate_component_name(course, "Lecture", j)
                        self.state.timetable[day][slot].append(course_variant_j)
                        self.state.all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Lecture"))
                        
                        for series_day in TimetableConfig.DAYS_MWF[1:]:
                            if (self.utils.is_valid_day_for_course(course, series_day, [(day, slot)]) 
                                    and self.utils.not_present_here(series_day, slot, course_variant_j, self.state.timetable)):
                                self.state.timetable[series_day][slot].append(course_variant_j)
                                self.state.all_course_assignments.setdefault(course_variant_j, []).append((series_day, slot, "Lecture"))

            # Assign tutorials
            tutorial_sections = load['Sections'].get('tutorials', 0)
            for i in tutorial_parallel:
                course_variant = self.utils.generate_component_name(course, "Tutorial", i[0])
                if tutorial_sections == 1:
                    course_variant = f"{course}_Tut"
                _, tutorials_left = self.assign_course_to_timetable(
                    course_variant, 0, load['LPU']['tutorials'], 0, TimetableConfig.DAYS_MWF, TimetableConfig.DAYS_TT
                )
                if tutorials_left > 0:
                    unassigned_courses[course_variant] = {'tutorials': tutorials_left}
                
                day, slot = next(((d, s) for d in TimetableConfig.ORDERED_DAYS 
                                for s in TimetableConfig.TIME_SLOTS 
                                if course_variant in self.state.timetable[d][s]), (None, None))
                
                if day and slot and isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = self.utils.generate_component_name(course, "Tutorial", j)
                        self.state.timetable[day][slot].append(course_variant_j)
                        self.state.visited[day][slot] = True
                        self.state.all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Tutorial"))

            # Assign labs
            lab_sections = load['Sections'].get('labs', 0)
            for i in range(1, lab_sections + 1):
                course_variant = self.utils.generate_component_name(course, "Lab", i)
                if lab_sections == 1:
                    course_variant = f"{course}_Lab"
                _, _ = self.assign_course_to_timetable(
                    course_variant, 0, 0, load['LPU']['lab_hours'], TimetableConfig.DAYS_MWF, TimetableConfig.DAYS_TT
                )

        # Assign TT courses
        for course, load in tt_courses:
            random_mwf_day = random.choice(TimetableConfig.DAYS_MWF)
            TimetableConfig.DAYS_TT.append(random_mwf_day)
            
            lecture_parallel = load['No of sections parallel'][0]
            tutorial_parallel = load['No of sections parallel'][1]
            lab_parallel = load['No of sections parallel'][2]
            
            lecture_sections = load['Sections'].get('lectures', 0)
            for i in lecture_parallel:
                course_variant = self.utils.generate_component_name(course, "Lecture", i[0])
                if lecture_sections == 1:
                    course_variant = f"{course}_Lecture"
                lectures_left, _ = self.assign_course_to_timetable(
                    course_variant, load['LPU']['lectures'], 0, 0, TimetableConfig.DAYS_TT, TimetableConfig.DAYS_MWF
                )
                if lectures_left > 0:
                    unassigned_courses[course_variant] = {'lectures': lectures_left}

                day, slot = next(((d, s) for d in TimetableConfig.ORDERED_DAYS 
                                for s in TimetableConfig.TIME_SLOTS 
                                if course_variant in self.state.timetable[d][s]), (None, None))
                
                if day and slot and isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = self.utils.generate_component_name(course, "Lecture", j)
                        self.state.timetable[day][slot].append(course_variant_j)
                        self.state.visited[day][slot] = True
                        self.state.all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Lecture"))
                        
                        for series_day in TimetableConfig.DAYS_TT[1:]:
                            if (self.utils.is_valid_day_for_course(course, series_day, [(day, slot)]) 
                                    and self.utils.not_present_here(series_day, slot, course_variant_j, self.state.timetable)):
                                self.state.timetable[series_day][slot].append(course_variant_j)
                                self.state.visited[series_day][slot] = True
                                self.state.all_course_assignments.setdefault(course_variant_j, []).append((series_day, slot, "Lecture"))

            # Assign tutorials
            tutorial_sections = load['Sections'].get('tutorials', 0)
            for i in tutorial_parallel:
                course_variant = self.utils.generate_component_name(course, "Tutorial", i[0])
                if tutorial_sections == 1:
                    course_variant = f"{course}_Tut"
                _, tutorials_left = self.assign_course_to_timetable(
                    course_variant, 0, load['LPU']['tutorials'], 0, TimetableConfig.DAYS_TT, TimetableConfig.DAYS_MWF
                )
                if tutorials_left > 0:
                    unassigned_courses[course_variant] = {'tutorials': tutorials_left}
                
                day, slot = next(((d, s) for d in TimetableConfig.ORDERED_DAYS 
                                for s in TimetableConfig.TIME_SLOTS 
                                if course_variant in self.state.timetable[d][s]), (None, None))
                
                if day and slot and isinstance(i, list):
                    for j in i[1:]:
                        course_variant_j = self.utils.generate_component_name(course, "Tutorial", j)
                        self.state.timetable[day][slot].append(course_variant_j)
                        self.state.visited[day][slot] = True
                        self.state.all_course_assignments.setdefault(course_variant_j, []).append((day, slot, "Tutorial"))

            # Assign labs
            lab_sections = load['Sections'].get('labs', 0)
            for i in range(1, lab_sections + 1):
                course_variant = self.utils.generate_component_name(course, "Lab", i)
                if lab_sections == 1:
                    course_variant = f"{course}_Lab"
                _, _ = self.assign_course_to_timetable(
                    course_variant, 0, 0, load['LPU']['lab_hours'], TimetableConfig.DAYS_TT, TimetableConfig.DAYS_MWF
                )
            
            TimetableConfig.DAYS_TT.remove(random_mwf_day)

        self.reassign_unassigned_courses(unassigned_courses)

    def reassign_unassigned_courses(self, unassigned_courses):
        for course, load in unassigned_courses.items():
            print(f"Reassigning remaining sessions for {course}...")

            lectures_left = load.get('lectures', 0)
            tutorials_left = load.get('tutorials', 0)
            labs_left = load.get('labs', 0)

            while lectures_left > 0:
                available_slots = [(day, slot) for day in TimetableConfig.ORDERED_DAYS 
                                 for slot in TimetableConfig.TIME_SLOTS_WITHOUT_8AM
                                 if self.utils.not_present_here(day, slot, course, self.state.timetable) 
                                 and self.utils.is_valid_slot(slot)]
                if available_slots:
                    day, slot = random.choice(available_slots)
                    self.state.timetable[day][slot].append(f"{course}")
                    self.state.visited[day][slot] = True
                    self.state.all_course_assignments.setdefault(course, []).append((day, slot, "Lecture"))
                    lectures_left -= 1
                else:
                    break

            while tutorials_left > 0:
                available_slots = [(day, slot) for day in TimetableConfig.ORDERED_DAYS 
                                 for slot in TimetableConfig.TIME_SLOTS
                                 if self.utils.not_present_here(day, slot, course, self.state.timetable) 
                                 and self.utils.is_valid_slot(slot, True)]
                if available_slots:
                    day, slot = random.choice(available_slots)
                    self.state.timetable[day][slot].append(f"{course}")
                    self.state.visited[day][slot] = True
                    self.state.all_course_assignments.setdefault(course, []).append((day, slot, "Tut"))
                    tutorials_left -= 1
                else:
                    break