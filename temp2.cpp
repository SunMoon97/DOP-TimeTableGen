#include <iostream>
#include <vector>
#include <unordered_map>
#include <string>
#include <fstream>
#include <iomanip>
#include <random>
#include <unordered_set>
#include <bits/stdc++.h>

using namespace std;

// Days of the week and available time slots (excluding 1 PM - 2 PM)
vector<string> days = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"};
vector<string> time_slots = {"8:00", "9:00", "10:00", "11:00", "12:00", "2:00", "3:00", "4:00", "5:00"};

// Course data for different semesters and branches with course names and loads
unordered_map<string, unordered_map<string, unordered_map<string, unordered_map<string, int>>>> course_data = {
    {"A7", {
        {"Year 1 Sem 1", {
            {"General Biology", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Engineering Graphics", {{"lectures", 0}, {"tutorials", 1}, {"labs", 3}}},
            {"Tech Report Writing", {{"lectures", 3}, {"tutorials", 0}, {"labs", 0}}},
            {"Chemistry Lab", {{"lectures", 0}, {"tutorials", 0}, {"labs", 2}}},
            {"General Chemistry", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Mathematics-I", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Probability & Statistics", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Physics Lab", {{"lectures", 0}, {"tutorials", 0}, {"labs", 2}}},
        }},
        {"Year 1 Sem 2", {
            {"Biological Lab", {{"lectures", 0}, {"tutorials", 0}, {"labs", 2}}},
            {"Thermodynamics", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Computer Programming", {{"lectures", 3}, {"tutorials", 0}, {"labs", 2}}},
            {"Electrical Sciences", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Mathematics-II", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
            {"Workshop Practice", {{"lectures", 0}, {"tutorials", 1}, {"labs", 3}}},
            {"Mechanical Oscillation & Waves", {{"lectures", 3}, {"tutorials", 1}, {"labs", 0}}},
        }},
    }}
};

// Dictionary to store assigned time slots for each course
unordered_map<string, unordered_map<string, string>> timetable;

// Utility function to reset timetable
void reset_timetable() {
    timetable.clear();
    for (const auto& day : days) {
        for (const auto& slot : time_slots) {
            timetable[day][slot] = "Free";
        }
    }
}

// Assign a course to the timetable, considering predefined lab hours
void assign_course_to_timetable(const string& course, int lectures, int tutorials, int labs) {
    vector<pair<string, string>> assigned_slots;
    unordered_set<string> days_used;

    // Random engine for choosing time slots and days
    random_device rd;
    mt19937 gen(rd());
    
    // Assign lectures (1 per day)
    for (int i = 0; i < lectures; ++i) {
        while (true) {
            uniform_int_distribution<> day_dist(0, days.size() - 1);
            uniform_int_distribution<> slot_dist(0, time_slots.size() - 1);
            
            string day = days[day_dist(gen)];
            string slot = time_slots[slot_dist(gen)];
            
            if (timetable[day][slot] == "Free" && days_used.find(day) == days_used.end()) {
                timetable[day][slot] = course;
                assigned_slots.push_back(make_pair(day, slot));
                days_used.insert(day);
                break;
            }
        }
    }

    // Assign tutorials
    for (int i = 0; i < tutorials; ++i) {
        while (true) {
            uniform_int_distribution<> day_dist(0, days.size() - 1);
            uniform_int_distribution<> slot_dist(0, time_slots.size() - 1);
            
            string day = days[day_dist(gen)];
            string slot = time_slots[slot_dist(gen)];
            
            if (timetable[day][slot] == "Free") {
                timetable[day][slot] = course + " Tut";
                assigned_slots.push_back(make_pair(day, slot));
                break;
            }
        }
    }

    // Assign lab (fixed duration, once a week)
    if (labs > 0) {
        while (true) {
            uniform_int_distribution<> day_dist(0, days.size() - 1);
            string day = days[day_dist(gen)];
            
            if (none_of(timetable[day].begin(), timetable[day].end(),
                        [](const pair<string, string>& entry) { return entry.second.find("Lab") != string::npos; })) {
                vector<int> valid_slots;
                for (int i = 0; i < time_slots.size() - labs + 1; ++i) {
                    if (time_slots[i] != "12:00" && all_of(time_slots.begin() + i, time_slots.begin() + i + labs,
                        [&](const string& slot) { return timetable[day][slot] == "Free"; })) {
                        valid_slots.push_back(i);
                    }
                }
                if (!valid_slots.empty()) {
                    uniform_int_distribution<> slot_dist(0, valid_slots.size() - 1);
                    int slot_index = valid_slots[slot_dist(gen)];
                    
                    for (int i = 0; i < labs; ++i) {
                        timetable[day][time_slots[slot_index + i]] = course + " Lab";
                    }
                    assigned_slots.push_back(make_pair(day, time_slots[slot_index]));
                    break;
                }
            }
        }
    }
}

// Save the timetable for this semester to a file
void save_timetable_to_file(const string& semester) {
    ofstream file("output.txt", ios::app);
    if (file.is_open()) {
        file << "\nTimetable for " << semester << ":\n";
        file << left << setw(12) << "Time Slot";
        for (const auto& day : days) {
            file << setw(15) << day;
        }
        file << "\n";
        for (const auto& slot : time_slots) {
            file << left << setw(12) << slot;
            for (const auto& day : days) {
                file << setw(15) << timetable[day][slot];
            }
            file << "\n";
        }
        file.close();
    }
}

// Generate the timetable for all semesters
void generate_timetable(const string& branch) {
    auto branch_courses = course_data.find(branch);
    if (branch_courses != course_data.end()) {
        for (auto& semester_courses : branch_courses->second) {
            string semester = semester_courses.first;
            cout << "\nAssigning courses for " << branch << " - " << semester << "...\n";

            // Reset timetable for each semester
            reset_timetable();

            // Assign courses to the timetable
            for (auto& course_load : semester_courses.second) {
                string course = course_load.first;
                int lectures = course_load.second["lectures"];
                int tutorials = course_load.second["tutorials"];
                int labs = course_load.second["labs"];

                assign_course_to_timetable(course, lectures, tutorials, labs);
            }

            // Save the timetable for this semester to a file
            save_timetable_to_file(semester);
        }
    }
}

// Main function to start the program
int main() {
    string branch = "A7";  // For computer science branch
    generate_timetable(branch);
    return 0;
}
