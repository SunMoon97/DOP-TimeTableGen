import pandas as pd
import numpy as np

class ClassroomAllocator:
    def __init__(self, timetable_file, erp_file, rooms_file):
        """
        Initialize the classroom allocator with input files
        
        Args:
            timetable_file (str): Path to the combined timetable Excel file
            erp_file (str): Path to the ERP data Excel file
            rooms_file (str): Path to the rooms data Excel file
        """
        # Load input files
        self.combined_timetable = pd.read_excel(timetable_file, sheet_name="All Course Assignments")
        self.erp_data = pd.read_excel(erp_file, sheet_name="Sheet1", header=1)
        self.rooms_data = pd.read_excel(rooms_file, sheet_name="3.rooms")
        
        # Preprocess data
        self._preprocess_data()
        
        # Track room usage
        self.room_usage = {}
        
    def _preprocess_data(self):
        """
        Preprocess and clean input data
        """
        # Create full course code
        self.erp_data["Course"] = (self.erp_data["Subject"].astype(str).str.strip() + 
                                   ' ' + self.erp_data["Catalog"].astype(str).str.strip()).str.replace('  ', ' ')
        
        # Extract course enrollments
        self.course_enrollment = self.erp_data[["Course", "No. Of Students"]]
        
        # Prepare room capacity data
        self.room_capacity = self.rooms_data[["Room", "Seating Capacity"]].copy()
        self.room_capacity["Seating Capacity"] = pd.to_numeric(
            self.room_capacity["Seating Capacity"], errors='coerce'
        )
        
        # Sort rooms by ascending seating capacity
        self.room_capacity = self.room_capacity.sort_values(by="Seating Capacity")
        
        # Filter out Lab sessions
        self.combined_timetable = self.combined_timetable[
            ~self.combined_timetable['Type'].str.contains('Lab', case=False, na=False)
        ]
    
    def _get_student_count(self, course):
        """
        Get student count for a given course
        
        Args:
            course (str): Course code
        
        Returns:
            int: Number of students or 0 if not found
        """
        try:
            return self.course_enrollment.loc[
                self.course_enrollment["Course"] == course, 
                "No. Of Students"
            ].values[0]
        except IndexError:
            return 0
    
    def _get_room_capacity(self, room):
        """
        Get seating capacity for a given room
        
        Args:
            room (str): Room name
        
        Returns:
            int: Room capacity or 0 if not found
        """
        try:
            return self.room_capacity.loc[
                self.room_capacity["Room"] == room, 
                "Seating Capacity"
            ].values[0]
        except IndexError:
            return 0
    
    def _check_room_availability(self, room, day, time):
        """
        Check if a room is available at a specific time
        
        Args:
            room (str): Room name
            day (str): Day of the week
            time (str): Time slot
        
        Returns:
            bool: True if room is available, False otherwise
        """
        key = (room, day, time)
        return key not in self.room_usage
    
    def _mark_room_usage(self, room, day, time, course):
        """
        Mark a room as used for a specific time slot
        
        Args:
            room (str): Room name
            day (str): Day of the week
            time (str): Time slot
            course (str): Course code
        """
        key = (room, day, time)
        self.room_usage[key] = course
    
    def allocate_classrooms(self):
        """
        Allocate classrooms using best-fit algorithm
        
        Returns:
            tuple: DataFrames for allocation and room schedule
        """
        classroom_allocation = {}
        room_schedule = {}
        room_specific_timetables = {}
        
        # Group timetable by unique course-day-time combinations
        unique_slots = self.combined_timetable.drop_duplicates(
            subset=['Course', 'Day', 'Time', 'Type']
        )
        
        for _, row in unique_slots.iterrows():
            course = row['Course'].split('-')[0]
            day = row['Day']
            time = row['Time']
            session_type = row['Type']
            
            # Get student count
            student_count = self._get_student_count(course)
            if student_count == 0:
                continue
            
            # Find best-fit room
            best_room = self._find_best_fit_room(student_count, day, time)
            
            if best_room:
                room_capacity = self._get_room_capacity(best_room)
                classroom_allocation[(course, day, time, session_type)] = {
                    'room': best_room,
                    'capacity': room_capacity,
                    'strength': student_count
                }
                room_schedule.setdefault((day, time), []).append((best_room, course, room_capacity, student_count))
                
                # Track room-specific timetable
                room_specific_timetables.setdefault(best_room, []).append({
                    'Day': day,
                    'Time': time,
                    'Course': course,
                    'Type': session_type,
                    'Room Capacity': room_capacity,
                    'Course Strength': student_count
                })
                
                # Mark room as used
                self._mark_room_usage(best_room, day, time, course)
            else:
                classroom_allocation[(course, day, time, session_type)] = {
                    'room': "No Room Available",
                    'capacity': 0,
                    'strength': student_count
                }
        
        # Create output DataFrames
        allocation_df = pd.DataFrame(
            [(k[0], k[1], k[2], k[3], v['room'], v['capacity'], v['strength']) 
             for k, v in classroom_allocation.items()],
            columns=["Course", "Day", "Time", "Type", "Room", "Room Capacity", "Course Strength"]
        )
        
        schedule_df = pd.DataFrame(
            [(k[0], k[1], v[0], v[1], v[2], v[3]) 
             for k, val in room_schedule.items() for v in val],
            columns=["Day", "Time", "Room", "Course", "Room Capacity", "Course Strength"]
        )
        
        return allocation_df, schedule_df, room_specific_timetables
    
    def _find_best_fit_room(self, student_count, day, time):
        """
        Find the smallest room that can accommodate students
        
        Args:
            student_count (int): Number of students in the course
            day (str): Day of the week
            time (str): Time slot
        
        Returns:
            str: Best-fit room name or None
        """
        # Sort rooms by seating capacity
        sorted_rooms = self.room_capacity.sort_values(by="Seating Capacity")
        
        for _, room_info in sorted_rooms.iterrows():
            room = room_info['Room']
            capacity = room_info['Seating Capacity']
            
            # Check if room can accommodate students and is available
            if (capacity >= student_count and 
                self._check_room_availability(room, day, time)):
                return room
        
        return None
    
    def save_to_excel(self, allocation_df, schedule_df, room_specific_timetables, output_file):
        """
        Save allocation, schedule, and room-specific timetables to Excel
        
        Args:
            allocation_df (pd.DataFrame): Course to room mapping
            schedule_df (pd.DataFrame): Room schedule
            room_specific_timetables (dict): Room-specific timetables
            output_file (str): Path to output Excel file
        """
        with pd.ExcelWriter(output_file) as writer:
            # Save main allocation sheets
            allocation_df.to_excel(writer, sheet_name="Course to Room Mapping", index=False)
            schedule_df.to_excel(writer, sheet_name="Full Room Schedule", index=False)
            
            # Create room-specific timetable sheets
            for room, timetable_data in room_specific_timetables.items():
                room_df = pd.DataFrame(timetable_data)
                
                # Create a pivot table with all info
                pivoted_df = room_df.pivot_table(
                    index='Time', 
                    columns='Day', 
                    values=['Course', 'Type', 'Room Capacity', 'Course Strength'], 
                    aggfunc=lambda x: x.iloc[0] if len(x) > 0 else ''
                )
                
                # Flatten multi-level column names
                pivoted_df.columns = [f'{col[1]}_{col[0]}' for col in pivoted_df.columns]
                
                # Sort the columns in a specific order
                day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
                final_cols = []
                for day in day_order:
                    final_cols.extend([
                        f'{day}_Course', 
                        f'{day}_Type',
                        f'{day}_Room Capacity',
                        f'{day}_Course Strength'
                    ])
                
                pivoted_df = pivoted_df.reindex(columns=final_cols)
                
                # Save to Excel with room name as sheet name
                pivoted_df.to_excel(writer, sheet_name=f"{room} Timetable")
        
        print(f"Classroom allocation completed! Check {output_file}")

# Main execution
def main():
    # File paths
    timetable_file = "combined_timetable.xlsx"
    erp_file = "sem 1 23-24 erp tt& reg.xlsx"
    rooms_file = "data.xlsx"
    output_file = "classroom_allocation.xlsx"
    
    # Initialize and run allocator
    allocator = ClassroomAllocator(timetable_file, erp_file, rooms_file)
    allocation_df, schedule_df, room_specific_timetables = allocator.allocate_classrooms()
    
    # Save results
    allocator.save_to_excel(allocation_df, schedule_df, room_specific_timetables, output_file)

if __name__ == "__main__":
    main()