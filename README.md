# ⏳ Time Sheet Analyzer

### 🚀 Automate Your Project Time Tracking Across Multiple Excel Files

This **Time Sheet Analyzer** is a Python script that **automates** the process of summing time entries for specific projects across multiple Excel time sheets. It scans through **multiple folders and subfolders**, extracts project hours, and provides a **total breakdown** of the time spent on a project.

---

## 📌 Features
✅ **Scans multiple folders** (e.g., `2024`, `2025`) to find all relevant time sheets  
✅ **Extracts project-specific hours** from Excel files  
✅ **Summarizes total time spent** on a project across all sheets  
✅ **Lists all files where the project was logged**  
✅ **Fast and efficient** — automates a tedious manual process  

---

## 🛠️ Installation

### **1. Clone the Repository**
```sh
git clone https://github.com/a-a-ronc/time-sheet-analyzer.git
cd time-sheet-analyzer

2. Install Dependencies
Ensure you have Python installed (≥3.7). Then, install the required libraries:
pip install pandas openpyxl

🎯 Usage

Run the script and follow the prompts:
python sum_project_time.py

🔹 Input:
Enter the path to the main folder containing all year folders (e.g., C:\Users\YourName\Time Sheet\).
Enter the project name you want to track.

🔹 Output:
The total hours spent on the project across all time sheets.
The number of files where the project was found.
A list of file names where the project appeared.

📂 Folder Structure

Time Sheet/
│
├── 2024/
│   ├── January_Time_Sheet.xlsx
│   ├── February_Time_Sheet.xlsx
│
├── 2025/
│   ├── March_Time_Sheet.xlsx
│   ├── April_Time_Sheet.xlsx
│
└── sum_project_time.py
👨‍💻 Example Output

Enter the main folder path containing time sheets: C:\Users\YourName\Time Sheet
Enter the project name: Project X

Total time allocated to project 'Project X': 45.5 hours
The project appeared in 3 files.
Files where the project appeared:
 - January_Time_Sheet.xlsx
 - March_Time_Sheet.xlsx
 - April_Time_Sheet.xlsx
🏆 Why Use This?
🔹 Saves time—no more manual data entry!
🔹 Improves accuracy—automates project time tracking.
🔹 Easy to use—just run and get insights instantly!

📜 License
This project is open-source under the MIT License.

🤝 Contributing
If you find bugs or have suggestions for improvements, feel free to fork the repo and submit a pull request! 🚀

🌟 Show Some Love
If this project helps you, give it a star ⭐ on GitHub!
Happy coding! 🎉

