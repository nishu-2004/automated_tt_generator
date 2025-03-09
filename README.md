# Automated Timetable Generator

![Hackathon Badge](https://img.shields.io/badge/Hackathon-Project-blue)  
![Python](https://img.shields.io/badge/Python-3.x-green)  
![License](https://img.shields.io/badge/License-MIT-orange)

This project is an **Automated Timetable Generator** developed during a hackathon. It is designed to automate the process of creating timetables for educational institutions, organizations, or personal use. The tool uses Python to generate timetables efficiently and avoids scheduling conflicts.


## Features

- **Automated Timetable Generation**: Generates timetables based on input constraints.
- **Conflict-Free Scheduling**: Ensures no overlapping schedules for resources (teachers, rooms, etc.).
- **Customizable Inputs**: Allows users to define subjects, teachers, rooms, and time slots.
- **Export Options**: Exports the generated timetable in a user-friendly format (e.g., CSV, Excel).
- **Easy to Use**: Simple command-line interface for quick setup and execution.

---

## How It Works

The script uses a combination of algorithms to allocate subjects, teachers, and rooms to available time slots while avoiding conflicts. It takes inputs such as:
- List of subjects
- List of teachers
- List of rooms
- Time slots
- Constraints (e.g., teacher availability, room capacity)

Based on these inputs, the script generates a timetable that satisfies all constraints.

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/nishu-2004/automated_tt_generator.git
   ```
2. Navigate to the project directory:
   ```bash
   cd automated_tt_generator
   ```
3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

1. Update the input files (e.g., `subjects.csv`, `teachers.csv`, `rooms.csv`) with your data.
2. Run the script:
   ```bash
   python TT_Automated.py
   ```
3. The generated timetable will be saved in the `output` folder.

---

## Contributing

Contributions are welcome! If you'd like to contribute, please follow these steps:
1. Fork the repository.
2. Create a new branch for your feature or bugfix.
3. Commit your changes.
4. Submit a pull request.

---

## Acknowledgments

- Thanks to the hackathon organizers for providing the platform.
- Inspiration from existing timetable generation tools.
- Contributors and teammates who helped in building this project.



---

This README provides a comprehensive overview of your project and makes it easy for others to understand and use your work. You can customize it further based on your specific requirements or additional features you’ve implemented. Good luck with your hackathon project! 🚀
