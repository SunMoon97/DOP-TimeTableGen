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
        
        # Extract course enrollments and sort by student count descending
        self.course_enrollment = self.erp_data[["Course", "No. Of Students"]]
        self.course_enrollment = self.course_enrollment.sort_values(
            by="No. Of Students", ascending=False
        )
        
        # Prepare room capacity data
        self.room_capacity = self.rooms_data[["Room", "Seating Capacity"]].copy()
        self.room_capacity["Seating Capacity"] = pd.to_numeric(
            self.room_capacity["Seating Capacity"], errors='coerce'
        )
        
        # Sort rooms by descending seating capacity (largest first)
        self.room_capacity = self.room_capacity.sort_values(
            by="Seating Capacity", ascending=False
        )
        
        # Filter out Lab sessions
        self.combined_timetable = self.combined_timetable[
            ~self.combined_timetable['Type'].str.contains('Lab', case=False, na=False)
        ]
        
        # Sort timetable by course size
        self.combined_timetable['Student_Count'] = self.combined_timetable['Course'].apply(
            lambda x: self._get_student_count(x.split('-')[0])
        )
        self.combined_timetable = self.combined_timetable.sort_values(
            by='Student_Count', ascending=False
        )
    
    def _get_student_count(self, course):
        """Get student count for a given course"""
        try:
            return self.course_enrollment.loc[
                self.course_enrollment["Course"] == course, 
                "No. Of Students"
            ].values[0]
        except IndexError:
            return 0
    
    def _get_room_capacity(self, room):
        """Get seating capacity for a given room"""
        try:
            return self.room_capacity.loc[
                self.room_capacity["Room"] == room, 
                "Seating Capacity"
            ].values[0]
        except IndexError:
            return 0
    
    def _check_room_availability(self, room, day, time):
        """Check if a room is available at a specific time"""
        key = (room, day, time)
        return key not in self.room_usage
    
    def _mark_room_usage(self, room, day, time, course):
        """Mark a room as used for a specific time slot"""
        key = (room, day, time)
        self.room_usage[key] = course
    
    def _find_best_fit_room(self, student_count, day, time):
        """
        Find the best-fit room that can accommodate students
        using minimum wastage strategy
        """
        best_room = None
        minimum_wastage = float('inf')
        
        for _, room_info in self.room_capacity.iterrows():
            room = room_info['Room']
            capacity = room_info['Seating Capacity']
            
            # Check if room can accommodate students and is available
            if (capacity >= student_count and 
                self._check_room_availability(room, day, time)):
                
                # Calculate wastage (unused seats)
                wastage = capacity - student_count
                
                # Update best room if this room has less wastage
                if wastage < minimum_wastage:
                    minimum_wastage = wastage
                    best_room = room
                
                # If wastage is less than 10% of capacity, return immediately
                if wastage <= capacity * 0.1:
                    return room
        
        return best_room
    
    def allocate_classrooms(self):
        """
        Allocate classrooms using best-fit algorithm with improved allocation strategy
        """
        classroom_allocation = {}
        room_schedule = {}
        room_specific_timetables = {}
        failed_assignments = []

        # Process courses in order of decreasing class size
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
            
            # Try primary allocation
            best_room = self._find_best_fit_room(student_count, day, time)
            
            if best_room:
                self._allocate_room(
                    course, day, time, session_type, student_count,
                    best_room, classroom_allocation, room_schedule,
                    room_specific_timetables
                )
            else:
                failed_assignments.append((course, day, time, session_type, student_count))
        
        # Handle failed assignments with more aggressive strategies
        self._handle_failed_assignments(
            failed_assignments, classroom_allocation,
            room_schedule, room_specific_timetables
        )
        
        # Create output DataFrames
        allocation_df = self._create_allocation_df(classroom_allocation)
        schedule_df = self._create_schedule_df(room_schedule)
        
        return allocation_df, schedule_df, room_specific_timetables
    
    def _allocate_room(self, course, day, time, session_type, student_count,
                      room, classroom_allocation, room_schedule, room_specific_timetables):
        """Helper method to allocate a room and update all relevant data structures"""
        room_capacity = self._get_room_capacity(room)
        
        classroom_allocation[(course, day, time, session_type)] = {
            'room': room,
            'capacity': room_capacity,
            'strength': student_count
        }
        
        room_schedule.setdefault((day, time), []).append(
            (room, course, room_capacity, student_count)
        )
        
        room_specific_timetables.setdefault(room, []).append({
            'Day': day,
            'Time': time,
            'Course': course,
            'Type': session_type,
            'Room Capacity': room_capacity,
            'Course Strength': student_count
        })
        
        self._mark_room_usage(room, day, time, course)
    
    def _handle_failed_assignments(self, failed_assignments, classroom_allocation,
                                 room_schedule, room_specific_timetables):
        """Handle failed assignments with multiple strategies"""
        for course, day, time, session_type, student_count in failed_assignments:
            # Strategy 1: Try different time slots on the same day
            if self._try_different_time_slot(
                course, day, time, session_type, student_count,
                classroom_allocation, room_schedule, room_specific_timetables
            ):
                continue
                
            # Strategy 2: Try room splitting if possible
            if self._try_room_splitting(
                course, day, time, session_type, student_count,
                classroom_allocation, room_schedule, room_specific_timetables
            ):
                continue
                
            # Strategy 3: Try different day
            if self._try_different_day(
                course, day, time, session_type, student_count,
                classroom_allocation, room_schedule, room_specific_timetables
            ):
                continue
                
            # If all strategies fail, mark as no room available
            classroom_allocation[(course, day, time, session_type)] = {
                'room': "No Room Available",
                'capacity': 0,
                'strength': student_count
            }
    
    def _try_different_time_slot(self, course, day, time, session_type, student_count,
                               classroom_allocation, room_schedule, room_specific_timetables):
        """Try to find a different time slot on the same day"""
        for alt_time in self.combined_timetable['Time'].unique():
            if alt_time != time:
                best_room = self._find_best_fit_room(student_count, day, alt_time)
                if best_room:
                    self._allocate_room(
                        course, day, alt_time, session_type, student_count,
                        best_room, classroom_allocation, room_schedule,
                        room_specific_timetables
                    )
                    return True
        return False
    
    def _try_room_splitting(self, course, day, time, session_type, student_count,
                          classroom_allocation, room_schedule, room_specific_timetables):
        """Try to split the class into two rooms if possible"""
        # This is a placeholder for room splitting logic
        # Implementation would depend on specific requirements
        return False
    
    def _try_different_day(self, course, day, time, session_type, student_count,
                          classroom_allocation, room_schedule, room_specific_timetables):
        """Try to find a slot on a different day"""
        for alt_day in self.combined_timetable['Day'].unique():
            if alt_day != day:
                best_room = self._find_best_fit_room(student_count, alt_day, time)
                if best_room:
                    self._allocate_room(
                        course, alt_day, time, session_type, student_count,
                        best_room, classroom_allocation, room_schedule,
                        room_specific_timetables
                    )
                    return True
        return False
    
    def _create_allocation_df(self, classroom_allocation):
        """Create allocation DataFrame from classroom_allocation dictionary"""
        return pd.DataFrame(
            [(k[0], k[1], k[2], k[3], v['room'], v['capacity'], v['strength']) 
             for k, v in classroom_allocation.items()],
            columns=["Course", "Day", "Time", "Type", "Room", "Room Capacity", "Course Strength"]
        )
    
    def _create_schedule_df(self, room_schedule):
        """Create schedule DataFrame from room_schedule dictionary"""
        return pd.DataFrame(
            [(k[0], k[1], v[0], v[1], v[2], v[3]) 
             for k, val in room_schedule.items() for v in val],
            columns=["Day", "Time", "Room", "Course", "Room Capacity", "Course Strength"]
        )
    
    def save_to_excel(self, allocation_df, schedule_df, room_specific_timetables, output_file):
        """Save all results to Excel file"""
        with pd.ExcelWriter(output_file) as writer:
            allocation_df.to_excel(writer, sheet_name="Course to Room Mapping", index=False)
            schedule_df.to_excel(writer, sheet_name="Full Room Schedule", index=False)
            
            # Create room-specific timetable sheets
            for room, timetable_data in room_specific_timetables.items():
                room_df = pd.DataFrame(timetable_data)
                
                pivoted_df = room_df.pivot_table(
                    index='Time', 
                    columns='Day', 
                    values=['Course', 'Type', 'Room Capacity', 'Course Strength'], 
                    aggfunc=lambda x: x.iloc[0] if len(x) > 0 else ''
                )
                
                pivoted_df.columns = [f'{col[1]}_{col[0]}' for col in pivoted_df.columns]
                
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
                pivoted_df.to_excel(writer, sheet_name=f"{room} Timetable")

def main():
    timetable_file = "combined_timetable1.xlsx"
    erp_file = "sem 1 23-24 erp tt& reg.xlsx"
    rooms_file = "data.xlsx"
    output_file = "classroom_allocation1.xlsx"
    
    allocator = ClassroomAllocator(timetable_file, erp_file, rooms_file)
    allocation_df, schedule_df, room_specific_timetables = allocator.allocate_classrooms()
    allocator.save_to_excel(allocation_df, schedule_df, room_specific_timetables, output_file)

if __name__ == "__main__":
    main()