import os
import pandas as pd

def sum_project_time(folder_path, project_name):
    total_time = 0
    file_count = 0
    file_list = []

    # Iterate through all subfolders (years) in the main folder
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)

                # Load the Excel file
                try:
                    sheet_data = pd.read_excel(file_path, header=None)

                    # Extract the project names and times
                    projects = sheet_data.loc[47:56, 2]  # C48:C57
                    times = sheet_data.loc[47:56, 9]    # J48:J57

                    # Check if the project appears in this file
                    found = False
                    for project, time in zip(projects, times):
                        if str(project).strip().lower() == project_name.strip().lower():
                            total_time += pd.to_numeric(time, errors='coerce')
                            found = True
                    
                    if found:
                        file_count += 1
                        file_list.append(file)

                except Exception as e:
                    print(f"Error processing file {file}: {e}")

    return total_time, file_count, file_list

if __name__ == "__main__":
    folder = input("Enter the main folder path containing time sheets: ").strip()
    project = input("Enter the project name: ").strip()

    total_time, file_count, file_list = sum_project_time(folder, project)
    
    print(f"Total time allocated to project '{project}': {total_time}")
    print(f"The project appeared in {file_count} files.")
    print("Files where the project appeared:")
    for f in file_list:
        print(f" - {f}")