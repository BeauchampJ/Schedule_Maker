import openpyxl
import random
import csv
import os
import sys

def check_availability(filename):
    """
    This function imports an Excel sheet and creates a dictionary containing information about available employees and their team lead status.

    Args:
        filename: The path to the Excel file.
    """
    if not os.path.isfile(filename):
      print(f"Warning: File '{filename}' does not exist in the current working directory.")
      sys.exit()
    # Open the workbook
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    # Get header row (assuming headers are in the first row)
    header = [cell.value for cell in sheet[1]]

    # Initialize dictionary for each day (skip first 3 columns for headers and employee names)
    data = {}
    for day in header[3::]:  # Slice header to get elements from the fourth column onward
        data[day] = {}

    # Iterate through each row in the sheet (skip first row header)
    for row in sheet.iter_rows(min_row=2):
        # Get the employee name (assuming employee names are in the second column)
        employee = row[1].value
        # Get the team lead status (assuming team lead status is in the third column)
        team_lead_value = row[2].value.lower()  # Convert team lead value to lowercase for case-insensitive check

        # Check for team lead status (considering "Yes" or "No" values)
        team_lead = team_lead_value in ("yes", "true")  # Check if value is "yes" or "true" (case-insensitive)

        # Iterate through each day (using index offset for skipping columns)
        for i, day in enumerate(data.keys()):
            # Check if the employee is available on that day
            availability_cell = row[i + 3]  # Adjust offset based on column skip (start from 4th)
            if availability_cell.value == "Yes":
                # Store employee information with availability and team lead status
                data[day][employee] = {'available': True, 'team_lead': team_lead}

    # Return the dictionary with day, employee information (availability and team lead), and team lead information
    return data


# Example usage
filename = "availability.xlsx"
days_available = check_availability(filename)


# Write availability to CSV file
with open("Final_Schedule.csv", 'w', newline='') as csvfile:
  writer = csv.writer(csvfile)

  # Write header row
  writer.writerow(["Shift", "Employee Name", "Team Lead Status"])

  # Iterate through days and employees
  for day, employees in days_available.items():
    if employees:
        # Ensure at least one team lead is scheduled
        team_leads = [emp for emp, info in employees.items() if info['team_lead']]
        if not team_leads:
            # Select a random employee and make them team lead if none exist
            random_employee = random.choice(list(employees.keys()))
            employees[random_employee]['team_lead'] = True
            team_leads.append(random_employee)

        # Check for sufficient employees (at least 2)
        available_count = len(employees)
        if available_count < 2:
          writer.writerow([day, "Emily", "Manager"])
          continue  # Skip to next day to avoid writing unavailable employees

        # Select two random employees (or all available if less than 2)
        selected_employees = random.sample(list(employees.keys()), min(2, available_count))

        # Write employee information to CSV
        for employee in selected_employees:
            team_lead = employees[employee]['team_lead']
            writer.writerow([day, employee, "Team Lead" if team_lead else "Team Member"])
    else:
        # Write "No Employees Available" for unavailable days
        writer.writerow([day, "No Employees Available", ""])

# Print the availability with desired format
print("Output Schedule saved to Final_Schedule.csv")
for day, employees in days_available.items():
    if employees:
        # Ensure at least one team lead is scheduled
        team_leads = [emp for emp, info in employees.items() if info['team_lead']]
        if not team_leads:
            # Select a random employee and make them team lead if none exist
            random_employee = random.choice(list(employees.keys()))
            employees[random_employee]['team_lead'] = True
            team_leads.append(random_employee)


