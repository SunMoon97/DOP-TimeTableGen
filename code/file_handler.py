import pandas as pd
import json
class FileHandler:
    @staticmethod
    def load_course_data(filename):
        with open(filename, 'r') as file:
            return json.load(file)

    @staticmethod
    def save_timetable_to_excel(all_semester_data, all_course_assignments):
        with pd.ExcelWriter("combined_timetable.xlsx", engine='xlsxwriter') as writer:
            for semester, branches in all_semester_data.items():
                combined_df = pd.DataFrame()
                
                for branch, timetable in branches.items():
                    df = pd.DataFrame(timetable)
                    df.index.name = 'Day'
                    df.columns.name = 'Time Slot'
                    branch_header = pd.DataFrame([[f"Branch: {branch}"]], columns=[None])
                    combined_df = pd.concat([combined_df, branch_header, df, pd.DataFrame([[]])])
                
                combined_df.to_excel(writer, sheet_name=f"Semester {semester}", index=True, header=True)

            course_assignments_df = pd.DataFrame(
                [(course, day, slot, session_type) 
                 for course, assignments in all_course_assignments.items() 
                 for day, slot, session_type in assignments],
                columns=['Course', 'Day', 'Time', 'Type']
            )
            course_assignments_df.to_excel(writer, sheet_name="All Course Assignments", index=False)
