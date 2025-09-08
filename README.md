This project is an AI-powered scheduling system that generates conflict-free university timetables using a Genetic Algorithm (GA) and a modern Tkinter + ttkbootstrap GUI.

Main features:
Automatic allocation of professors, rooms, and time slots; handling of hard constraints such as room conflicts, professor load limits, and year clashes; optimization of soft constraints like minimizing gaps and balancing workloads; support for two distribution modes (by students or by sections); and Excel input/output for courses, professors, and rooms.

Usage:
Prepare three Excel files (courses, professors, rooms), run the main program file "Main Program GUI.py", choose semester and distribution mode, upload the files, and click Generate. The system will create and export the final timetable to Excel.

Project structure:
Data Generation GUI.py → for creating and editing Excel input files.
Main Program GUI.py → for running the GA and generating the timetable.
Example Excel files (courses, professors, rooms).

Requirements: Python 3.9+ with ttkbootstrap, pandas, and openpyxl installed.
