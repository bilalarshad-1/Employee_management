import pandas as pd
import os
from datetime import datetime

EMPLOYEE_FILE = "employees.xlsx"


def load_employees():
    if os.path.exists(EMPLOYEE_FILE):
        df = pd.read_excel(EMPLOYEE_FILE)
        df.columns = df.columns.str.strip() 
        return df.to_dict(orient='records')
    return []


def save_employees(employees):
    df = pd.DataFrame(employees)
    df.to_excel(EMPLOYEE_FILE, index=False)
    print("Employees saved successfully!")


def get_employee_attendance_file(emp_num):
    return f"employee_attendance_{emp_num}.xlsx"


def get_daily_attendance_file():
    today = datetime.today().strftime("%Y-%m-%d")
    return f"daily_attendance_{today}.xlsx"


def load_employee_attendance(emp_num):
    file_path = get_employee_attendance_file(emp_num)
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        return df.to_dict(orient='records')
    return []


def save_employee_attendance(emp_num, attendance):
    file_path = get_employee_attendance_file(emp_num)
    df = pd.DataFrame(attendance)
    df.to_excel(file_path, index=False)
    print(f"Employee {emp_num} attendance saved!")


def load_daily_attendance():
    file_path = get_daily_attendance_file()
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        return df.to_dict(orient='records')
    return []


def save_daily_attendance(attendance):
    file_path = get_daily_attendance_file()
    df = pd.DataFrame(attendance)
    df.to_excel(file_path, index=False)
    print(f"Daily attendance saved in {file_path}!")


def list_employees(employees):
    if not employees:
        print(" No employees found! Please add employees first.")
        return
    df = pd.DataFrame(employees)
    print("\nEmployee List:\n", df)


def add_employee(employees):
    name = input("Enter Employee Name: ")
    department = input("Enter Department: ")
    role = input("Enter Role: ")
    employee_number = len(employees) + 1 

    employees.append({
        'Employee Number': employee_number,
        'Name': name,
        'Department': department,
        'Role': role
    })

    save_employees(employees)
    print(f" Employee {name} added with Employee Number: {employee_number}")

def mark_attendance(employees):
    if not employees:
        print(" No employees found! Add employees first.")
        return

    list_employees(employees)

    try:
        emp_num = int(input("Enter Employee Number to mark attendance: "))
        employee = next((emp for emp in employees if int(emp['Employee Number']) == emp_num), None)
    except ValueError:
        print(" Invalid input! Enter a valid employee number.")
        return

    if not employee:
        print(" Employee not found!")
        return

    date = datetime.today().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M")
    status = input("Enter Status (P for Present, A for Absent): ").upper()

    if status not in ["P", "A"]:
        print(" Invalid status! Use 'P' for Present or 'A' for Absent.")
        return

    late = "Late ()" if int(current_time.split(':')[0]) >= 9 else "On Time (✅)"

    record = {
        'Date': date,
        'Time': current_time,
        'Status': status,
        'Late Status': late
    }

   
    emp_attendance = load_employee_attendance(emp_num)
    emp_attendance.append(record)
    save_employee_attendance(emp_num, emp_attendance)

    
    daily_attendance = load_daily_attendance()
    daily_attendance.append({
        'Employee Number': emp_num,
        'Name': employee['Name'],
        'Date': date,
        'Time': current_time,
        'Status': status,
        'Late Status': late
    })
    save_daily_attendance(daily_attendance)

    print(f"Attendance marked for {employee['Name']} at {current_time}.")


def apply_leave(employees):
    if not employees:
        print("No employees found! Add employees first.")
        return

    list_employees(employees)

    try:
        emp_num = int(input("Enter Employee Number to apply for leave: "))
        employee = next((emp for emp in employees if int(emp['Employee Number']) == emp_num), None)
    except ValueError:
        print(" Invalid input! Enter a valid employee number.")
        return

    if not employee:
        print(" Employee not found!")
        return

    start_date = input("Enter Leave Start Date (YYYY-MM-DD): ")
    end_date = input("Enter Leave End Date (YYYY-MM-DD): ")

    record = {
        'Date': 'N/A',
        'Time': 'N/A',
        'Status': 'On Leave',
        'Late Status': 'N/A',
        'Leave Period': f"{start_date} to {end_date}"
    }

   
    emp_attendance = load_employee_attendance(emp_num)
    emp_attendance.append(record)
    save_employee_attendance(emp_num, emp_attendance)

    print(f" Leave applied for {employee['Name']} from {start_date} to {end_date}")


def view_employee_attendance():
    try:
        emp_num = int(input("Enter Employee Number to view attendance: "))
    except ValueError:
        print("Invalid input! Enter a valid employee number.")
        return

    emp_attendance = load_employee_attendance(emp_num)

    if not emp_attendance:
        print(" No attendance records found for this employee!")
        return

    df = pd.DataFrame(emp_attendance)
    print(f"\n Attendance for Employee {emp_num}:\n", df)


def view_daily_attendance():
    daily_attendance = load_daily_attendance()

    if not daily_attendance:
        print("No attendance records found for today!")
        return

    df = pd.DataFrame(daily_attendance)
    print("\nToday's Attendance:\n", df)


def main():
    employees = load_employees()

    while True:
        print("\nEmployee Attendance System")
        print("1 Add Employee")
        print("2 Mark Attendance")
        print("3 Apply for Leave")
        print("4 View Employees")
        print("5 View Employee Attendance")
        print("6 View Daily Attendance")
        print("7 Exit")

        choice = input("Enter your choice: ")

        match choice:
            case '1':
                add_employee(employees)
            case '2':
                mark_attendance(employees)
            case '3':
                apply_leave(employees)
            case '4':
                list_employees(employees)
            case '5':
                view_employee_attendance()
            case '6':
                view_daily_attendance()
            case '7':
                break
            case _:
                print(" Invalid choice! Try again.")

if __name__ == "__main__":
    main()
