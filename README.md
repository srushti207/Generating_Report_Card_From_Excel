# Generating_Report_Card_From_Excel
Generating Report Cards from Excel (PDF form)

## Task: Generating Report Cards from Excel (PDF form)
You are given an Excel file named student_scores.xlsx containing the following columns:

Student ID | Name | Subject Score
Instructions:
Please write a Python script that does the following:

Reads the data from the Excel file using the pandas library. Groups the data by student and calculates their total and average scores across all subjects. Generates a personalized PDF report card for each student using the ReportLab library. Each report card should include: Student Name / Total Score / Average Score/ A table showing subject-wise scores. Save the PDF files with the naming format report_card_<StudentID>.pdf. Ensure that your script includes proper error handling for missing or invalid data in the Excel file.


## Explanation
1. Reading the Data:
Use the pandas library to load the data from an Excel file (student_scores.xlsx).
Validate that the necessary columns (Student ID, Name, Subject, Score) are present. Handle missing or invalid data by filling or discarding incomplete rows.

2. Aggregating Data:
Group the data by Student ID and Name to compute each student's total and average scores.
Extract subject-wise scores for individual students for detailed reporting.

3. Generating PDF Report Cards:
Use the ReportLab library to create PDF files.
Each report includes:
  The student’s name.
  Their total score and average score.
  A table displaying their subject-wise scores.
Apply styling to the table for better readability.

4. Saving PDF Files:
Name each file using the format report_card_<StudentID>.pdf and save it in the working directory.

5. Error Handling:
Include error handling for file-related issues (e.g., missing file), missing columns, or unexpected data formats.
