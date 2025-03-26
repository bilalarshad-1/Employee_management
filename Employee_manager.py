import pandas as pd
import os


FILE_PATH = "Employees.xlsx"


def load_data():
    if os.path.exists(FILE_PATH):
        try:
            df = pd.read_excel(FILE_PATH)
            return df.to_dict(orient="records")  
        except Exception as e:
            print("Error loading data:", e)
            return []
    return []


def save_data(emp):
    df = pd.DataFrame(emp)
    df.to_excel(FILE_PATH, index=False) 


def list_all_emp(emp):
    if not emp:
        print("\nNo employees found!\n")
        return

    print("\n" + "*" * 70)
    for index, employee in enumerate(emp, start=1):
        print(f"{index}. {employee['Name']}, Department: {employee['Department']}, Role: {employee['Role']}, "
              f"Joining Date: {employee['Joining Date']}, Monthly Salary: ${employee['Monthly Salary']}")
    print("*" * 70 + "\n")


def add_employee(emp):
    name = input("Enter Employee Name: ")
    department = input("Enter Employee Department: ") or "Not Assigned"
    role = input("Enter Employee Role: ") or "Not Specified"
    joining_date = input("Enter Employee Joining Date (YYYY-MM-DD): ") or "Unknown"
    monthly_salary = input("Enter Employee Monthly Salary: $") or "0"

    emp.append({
        'Name': name,
        'Department': department,
        'Role': role,
        'Joining Date': joining_date,
        'Monthly Salary': monthly_salary
    })
    
    save_data(emp)
    print("\nEmployee added successfully!\n")


def update_employee(emp):
    list_all_emp(emp)
    
    if not emp:
        return

    try:
        index = int(input("Enter the Employee number to update: ")) - 1
        if 0 <= index < len(emp):
            emp[index]['Name'] = input("Enter New Employee Name: ") or emp[index]['Name']
            emp[index]['Department'] = input("Enter New Employee Department: ") or emp[index]['Department']
            emp[index]['Role'] = input("Enter New Employee Role: ") or emp[index]['Role']
            emp[index]['Joining Date'] = input("Enter New Employee Joining Date (YYYY-MM-DD): ") or emp[index]['Joining Date']
            emp[index]['Monthly Salary'] = input("Enter New Employee Monthly Salary: $") or emp[index]['Monthly Salary']
            
            save_data(emp)
            print("\nEmployee details updated successfully!\n")
        else:
            print("Invalid index selected!")
    except ValueError:
        print("Please enter a valid number!")


def delete_employee(emp):
    list_all_emp(emp)
    
    if not emp:
        return

    try:
        index = int(input("Enter the Employee number to delete: ")) - 1
        if 0 <= index < len(emp):
            del emp[index]
            save_data(emp)
            print("\nEmployee deleted successfully!\n")
        else:
            print("Invalid index selected!")
    except ValueError:
        print("Please enter a valid number!")


def main():
    emp = load_data()

    while True:
        print("\nEmployee Management System | Choose an option:")
        print("1. List All Employees")
        print("2. Add New Employee")
        print("3. Update Employee Details")
        print("4. Delete Employee")
        print("5. Exit")

        choice = input("Enter your choice: ")

        match choice:
            case '1':
                list_all_emp(emp)
            case '2':
                add_employee(emp)
            case '3':
                update_employee(emp)
            case '4':
                delete_employee(emp)
            case '5':
                print("Exiting the application. Goodbye!")
                break
            case _:
                print("Invalid choice! Please enter a number between 1 and 5.")


if __name__ == "__main__":
    main()
