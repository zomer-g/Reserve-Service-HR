# Manpower Control Script

## Purpose
This script is designed to allow control of manpower data based on a central Gantt chart in which the names of the soldiers are entered. It processes data from multiple sheets and generates comprehensive reports to track the presence and activities of personnel.

## Structure

### Sheets

#### Gantt chart
- **Description**: The main source of data. It includes the schedule and assignment of personnel.
- **Range**: `A3:AT100`
- **Columns**:
  - **A**: Category (Type of activity)
  - **B**: Sub-category (Detailed description of the activity)
  - **C-AT**: Dates from 22/05/2024 to 04/07/2024. Cells contain names of people assigned to activities on each day.

#### ID
- **Description**: Contains the names and ID numbers of all personnel.
- **Range**: `A:D`
- **Columns**:
  - **A**: Name
  - **D**: ID

#### Summary
- **Description**: Summarizes the names, IDs, and their assigned dates.
- **Columns**:
  - **A**: Name
  - **B**: ID
  - **C**: Dates (Grouped by continuous periods)

#### Kishur
- **Description**: Lists all unique dates in chronological order and indicates the start and end of personnel assignments.
- **Columns**:
  - **A**: Dates
  - **B**: Start Date (Names that appear on this date but not on the previous date)
  - **C**: End Date (Names that appear on this date but not on the next date)

#### Report1
- **Description**: Generates a report for specific categories ("קו" and "מפלג") for the current date.
- **Columns**:
  - **A**: Category
  - **B-D**: Sub-categories (Splitting the content of column B in the Gantt chart)
  - **E**: Name
  - **F**: ID

#### Counter
- **Description**: Provides a count of the number of days each personnel appears in each category and in total.
- **Columns**:
  - **A**: Name
  - **B**: ID
  - **C**: Columns for each category (Number of days each name appears in the category)
  - **Last Column**: "All" (Total number of days across all categories)

## Script Functionality

1. **Extract and Log Unique Names and Details**:
   - Fetch data from the "Gantt chart" sheet.
   - Extract unique names from the date columns.
   - Create a vocabulary for each name, including their ID and dates of assignment.
   - Group continuous dates.

2. **Write to Summary Sheet**:
   - Clear existing data.
   - Write names, IDs, and grouped dates.

3. **Write to Kishur Sheet**:
   - Collect all unique dates.
   - Identify and log start and end dates for each personnel.

4. **Generate Report1**:
   - Generate a report for "קו" and "מפלג" categories.
   - Write today's assignments and IDs.

5. **Write to Counter Sheet**:
   - List names and IDs sorted by name.
   - Count the number of days each name appears in each category.
   - Calculate the total number of days across all categories for each name.

## Usage Instructions

1. **Open your Google Sheet**.
2. **Go to Extensions > Apps Script**.
3. **Delete any existing code in the script editor and replace it with the provided script**.
4. **Save the script**.
5. **Run the function `extractAndLogUniqueNamesAndDetails`**.

This script is designed to be run as a Google Apps Script. It processes data from the specified sheets and generates comprehensive reports to facilitate manpower control and management. Make sure to structure your sheets as described for the script to work correctly.
