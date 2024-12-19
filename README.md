# PubRanker
## Introduction
The committee’s decision on whether a PhD student will receive a scholarship is often influenced by the number of publications the student has. Applying for a scholarship can be a time-consuming process, involving obtaining a recommendation from the principal investigator (PI), writing a letter of motivation, and preparing a resume. This competition for scholarships often occurs within the departments of the Weizmann Institute of Science. **PubRanker** is designed to help students estimate their chances of winning the fellowship relative to other PhD students in their department.

## How does it work?
**PubRanker** collects data from the website of the selected Weizmann department, retrieving the names of the PhD students and their corresponding PIs. Using this data, the app performs two checks:

1. **Department-specific publications**: The app queries the PubMed database to find publications coauthored by each student and their PI, estimating how many publications each student has within their Weizmann Institute department.
2. **Total publications**: The app also checks PubMed for publications associated with each student's name to determine the student's total number of publications, regardless of the Weizmann Institute of Science. To reduce the chances of including publications from individuals with similar names, the data is limited to the last 10 years.

**PubRanker** then generates two separate plots in an Excel file:
- One plot shows the number of department-specific publications for each student, with the x-axis representing the students' names and the y-axis showing the number of publications within the department.
- The second plot shows the total number of publications for each student, with the x-axis representing the students' names and the y-axis showing their total publications.

## Conclusion
With the generated data and plots, PhD students using **PubRanker** will be able to estimate their standing in the publication ranking within their department as well as in total. This will help them determine whether they have a competitive chance of winning the scholarship and whether applying for it is worthwhile.
