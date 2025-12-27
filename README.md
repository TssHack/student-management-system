<div align="center">

  # 🎓 Student Management System

  ![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
  ![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-green?logo=qt)
  ![License](https://img.shields.io/badge/License-MIT-yellow)

  A comprehensive, modern desktop application for managing student records, grades, and statistics. Built with Python and PyQt5, featuring a fully localized Persian (RTL) interface.

  [![Open Source Love](https://badges.frapsoft.com/os/v2/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badges)
  
</div>

---

## 📸 Screenshots

*(Add your screenshots here)*

| Dashboard | Student List | Statistics |
| :---: | :---: | :---: |
| ![Dashboard](https://raw.githubusercontent.com/TssHack/student-management-system/refs/heads/main/1.png) | ![List](https://raw.githubusercontent.com/TssHack/student-management-system/refs/heads/main/2.png) | ![ab](https://raw.githubusercontent.com/TssHack/student-management-system/refs/heads/main/abute.png) |

---

## ✨ Features

- **👤 Student Management:** Full CRUD (Create, Read, Update, Delete) operations for student data.
- **📊 Grade Calculation:** Automatic calculation of weighted averages (30% Midterm, 70% Final) with manual override.
- **🏆 Ranking System:** One-click ranking of students based on their average scores.
- **🔍 Advanced Search:** Filter students by name, ID, or grade range.
- **📈 Visual Statistics:** Interactive charts displaying grade distribution.
- **📥 Excel Integration:** Import student data from Excel and export reports.
- **💾 Backup & Restore:** Built-in tools to backup and restore the SQLite database.
- **🌐 RTL Support:** Native Right-to-Left layout support optimized for the Persian language.
- **📸 Photo Management:** Attach student photos to profiles.

---

## 🚀 Installation

To run this project locally, follow these steps:

### Prerequisites

- Python 3.8 or higher installed on your machine.

### Setup

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/your-username/student-management-system.git
    cd student-management-system
    ```

2.  **Install required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the application:**
    ```bash
    python main.py
    ```

*Note: The application will automatically create a `students.db` SQLite file upon the first run.*

---

## 🛠️ Tech Stack

This project is built using the following technologies:

*   **Language:** Python 3
*   **GUI Framework:** PyQt5
*   **Database:** SQLite3
*   **Data Processing:** Pandas, OpenPyXL
*   **Visualization:** Matplotlib

---

## 📖 Usage Guide

1.  **Dashboard:** View quick statistics like total students, class average, and top performers upon launching.
2.  **Adding Students:** Click the "➕ Add Student" button. Fill in the details. Grades are automatically calculated, but you can edit the final average manually if needed.
3.  **Editing/Deleting:** Select a student from the table and use the buttons in the toolbar or the Edit menu.
4.  **Ranking:** Click the "🏆 Rank" button to calculate and assign ranks based on the current averages.
5.  **Excel Export:** Go to `File > Export Excel` to save the current list for reporting.

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!

1.  Fork the Project
2.  Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3.  Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4.  Push to the Branch (`git push origin feature/AmazingFeature`)
5.  Open a Pull Request

---

## 👨‍💻 Author

**Ehsan Fazli**

- GitHub: [tsshack](https://github.com/tsshack)
- Project Link: [https://github.com/tsshack/student-management-system](https://github.com/tsshack/student-management-system)

---

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information.

---

<div align="center">
  <b>⭐ If you like this project, please give it a star! ⭐</b>
</div>
