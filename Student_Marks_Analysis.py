import os
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import psutil
import time

class StudentDataSystem:
    def __init__(self):
        self.folder = "Student_Data"
        self.file_path = os.path.join(self.folder, "Student_Data.xlsx")
        self.subjects = ["Maths", "Physics", "Chemistry", "English", "Computer"]
        os.makedirs(self.folder, exist_ok=True)

    # Excel close
    def close_excel_if_open(self):
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
                    print("Closing Excel process...")
                    proc.kill()
                    print("Excel closed successfully.")
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                pass

    # Wait for file availability
    def wait_for_file(self, timeout=5):
        start = time.time()
        while True:
            try:
                if os.path.exists(self.file_path):
                    with open(self.file_path, 'rb'):
                        break  # file is accessible
                else:
                    break  # file does not exist yet
            except PermissionError:
                if time.time() - start > timeout:
                    raise PermissionError(f"File still locked after {timeout} seconds.")
                time.sleep(0.5)  # wait before retry

    # Insert Data
    def insert_data(self):
        self.close_excel_if_open()
        self.wait_for_file()

        if os.path.exists(self.file_path):
            old_df = pd.read_excel(self.file_path)
            if not old_df.empty:
                print("Existing Data in Excel:")
                print(old_df)
                while True:
                    clear = input("Do you want to clear the existing data? (yes/no): ").strip().lower()
                    if clear in ["yes", "no"]:
                        break
                    print("Invalid input. Please type 'yes' or 'no'.")
                if clear == "yes":
                    old_df = pd.DataFrame(columns=["Name"] + self.subjects)
                    self.close_excel_if_open()
                    self.wait_for_file()
                    old_df.to_excel(self.file_path, index=False)
                    print("Existing data cleared. You can enter new data now.")
            else:
                old_df = pd.DataFrame(columns=["Name"] + self.subjects)
        else:
            old_df = pd.DataFrame(columns=["Name"] + self.subjects)

        # Ensure numeric columns
        for sub in self.subjects:
            if sub not in old_df.columns:
                old_df[sub] = pd.Series(dtype=float)
            else:
                old_df[sub] = pd.to_numeric(old_df[sub], errors='coerce')

        # Input number of students
        while True:
            try:
                n = int(input("Insert number of students: "))
                if n <= 0:
                    print("Please enter a number greater than 0.")
                    continue
                break
            except ValueError:
                print("Invalid input! Please enter a numeric value.")

        data = []
        for i in range(n):
            print(f"Insert student {i + 1}")
            name = input("Insert student name: ")
            marks = []
            for sub in self.subjects:
                while True:
                    try:
                        mark = float(input(f"Insert marks for {sub}: "))
                        if mark < 0 or mark > 100:
                            print("Marks should be between 0 and 100.")
                            continue
                        marks.append(mark)
                        break
                    except ValueError:
                        print("Invalid input! Please enter a numeric value.")
            data.append([name] + marks)

        new_df = pd.DataFrame(data, columns=["Name"] + self.subjects)
        final_df = pd.concat([old_df, new_df], ignore_index=True)

        self.close_excel_if_open()
        self.wait_for_file()
        final_df.to_excel(self.file_path, index=False)
        print("Data saved successfully")
        print(f"File location: {self.file_path}")

    # Calculate Statistics
    def calculate_statistics(self):
        self.close_excel_if_open()
        self.wait_for_file()

        if not os.path.exists(self.file_path):
            print("File does not exist")
            return

        df = pd.read_excel(self.file_path)
        if df.empty:
            print("Excel file empty")
            return

        students = df["Name"].tolist()
        marks = df[self.subjects].to_numpy()

        total_marks = np.sum(marks, axis=1)
        df['Total Marks'] = total_marks
        average_marks = np.mean(marks, axis=1)
        df['Average Marks'] = average_marks
        percentage = (total_marks / (len(self.subjects) * 100)) * 100
        df['Percentage'] = percentage

        grades = []
        for p in percentage:
            if p >= 90:
                grades.append("A+")
            elif p >= 80:
                grades.append("A")
            elif p >= 70:
                grades.append("B")
            elif p >= 60:
                grades.append("C")
            else:
                grades.append("Fail")
        df['Grade'] = grades

        rank_order = np.argsort(-total_marks)
        ranks = np.empty_like(rank_order)
        ranks[rank_order] = np.arange(1, len(students) + 1)
        df['Rank'] = ranks

        topper_index = np.argmax(total_marks)
        topper_name = df.loc[topper_index, "Name"]
        topper_marks = df.loc[topper_index, "Total Marks"]

        print("===== STUDENT MARKS ANALYSIS =====")
        for i, s in enumerate(students):
            print(f"{s} - Total: {total_marks[i]}, Avg: {round(average_marks[i], 2)}, "
                  f"Percentage: {round(percentage[i], 2)}%, Grade: {grades[i]}, Rank: {ranks[i]}")

        print(f"\nTopper of the Class: {topper_name}, Total Marks: {topper_marks}")

        subject_avg = np.mean(marks, axis=0)
        print("\nSubject-wise Average Marks:")
        for i, sub in enumerate(self.subjects):
            print(f"{sub}: {round(subject_avg[i], 2)}")

        print("\nSubject-wise Highest & Lowest Marks:")
        for i, sub in enumerate(self.subjects):
            print(f"{sub} Highest: {np.max(marks[:, i])}, Lowest: {np.min(marks[:, i])}")

        stats_df = pd.DataFrame({
            "Name": ["Subject-wise Average", "Subject-wise Highest", "Subject-wise Lowest"]
        })
        for i, sub in enumerate(self.subjects):
            stats_df[sub] = [round(subject_avg[i], 2), np.max(marks[:, i]), np.min(marks[:, i])]
        final_df = pd.concat([df, stats_df], ignore_index=True)

        self.close_excel_if_open()
        self.wait_for_file()
        final_df.to_excel(self.file_path, index=False)

        # Auto-adjust column widths
        wb = openpyxl.load_workbook(self.file_path)
        ws = wb.active
        for column_cells in ws.columns:
            max_length = 0
            column = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            ws.column_dimensions[column].width = max_length + 2
        wb.save(self.file_path)
        print(f"All statistics saved to Excel with auto-adjusted column widths: {self.file_path}")

    # Visualize Data
    def visualize_data(self):
        self.close_excel_if_open()
        self.wait_for_file()

        if not os.path.exists(self.file_path):
            print("File does not exist")
            return

        df = pd.read_excel(self.file_path)
        if df.empty:
            print("Excel file empty")
            return

        students = df["Name"]
        marks = df.iloc[:, 1:len(self.subjects) + 1]

        #Bar Graph
        plt.figure()
        plt.bar(students, marks.sum(axis=1))
        plt.xlabel("Students")
        plt.ylabel("Total Marks")
        plt.title("Total Marks of Students")
        plt.xticks(rotation=30)
        plt.tight_layout()
        plt.show()

        #Bar Graph
        plt.figure(figsize=(10, 6))
        x = np.arange(len(marks.columns))
        width = 0.8 / len(students)
        for i, student in enumerate(students):
            plt.bar(x + i * width, marks.iloc[i], width, label=student)
        plt.xlabel("Subjects")
        plt.ylabel("Marks")
        plt.title("Subject-wise Marks Comparison")
        plt.xticks(x + width * (len(students) - 1) / 2, marks.columns)
        plt.legend()
        plt.tight_layout()
        plt.show()

        plt.figure()
        for i in range(len(students)):
            plt.plot(marks.columns, marks.iloc[i], marker='o', label=students[i])
        plt.xlabel("Subjects")
        plt.ylabel("Marks")
        plt.title("Student-wise Performance")
        plt.legend()
        plt.tight_layout()
        plt.show()

    # Menu
    def menu(self):
        while True:
            print("\n===== STUDENT MARKS ANALYSIS SYSTEM =====")
            print("1. Insert Student Data")
            print("2. Read Excel & Visualize Data")
            print("3. Calculate Statistics")
            print("4. Exit")

            choice = input("Enter your choice (1-4): ")

            if choice == "1":
                self.insert_data()
            elif choice == "2":
                self.visualize_data()
            elif choice == "3":
                self.calculate_statistics()
            elif choice == "4":
                print("Program exited successfully.")
                # Open the Excel file at the end
                if os.path.exists(self.file_path):
                    print("Opening Excel file...")
                    os.startfile(self.file_path)
                break
            else:
                print("Invalid choice. Try again.")


system = StudentDataSystem()
system.menu()