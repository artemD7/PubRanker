import tkinter as tk
from tkinter import ttk
from bs4 import BeautifulSoup
import requests
from pymed import PubMed
import xlsxwriter
import time
# Function for filtering PI list from Prof., Dr., and PI:
def prof_dr_mr_filter(item):
    no_need = ['PI:', 'Prof.', 'Dr.', 'Mr.', 'Ms.']
    return False if item in no_need else True

# General function to fetch and rank students based on department URL and HTML of the website 
def fetch_and_rank_students(department):
    if department == 'Chemical and Structural Biology':
        url = 'http://www.weizmann.ac.il/CSB/people'
        phd_id = 'quicktabs-tabpage-people-2'
        stud_class = 'member_info'
        pi_class = 'pi-name'
    elif department == 'Molecular Genetics':
        url = 'https://www.weizmann.ac.il/molgen/people'
        phd_id = 'quicktabs-tabpage-view__people__page-7'
        stud_class = 'quicktabs-views-group'
        pi_class = 'h5 website'
    elif department == 'Systems Immunology':
        url = 'https://www.weizmann.ac.il/immunology/people'
        phd_id = 'panel-graduate-students'
        stud_class = 'person-wrapper rounded-corners graduate-students'
        pi_class = None  # Assuming no PI info is needed for Systems Immunology
    else:
        return []

    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    PhD_page = soup.find('div', id=phd_id)
    stud_info = PhD_page.findAll('div', class_=stud_class)

    student_name_list = []
    stud_pi_couple = []
    
    for stud in stud_info:
        # Extract student name based on department
        if department == 'Systems Immunology':
            name = stud.find('h2').text
        else:
            name = stud.find('div', class_='people_title').text if department == 'Chemical and Structural Biology' else stud.find('h3').text

        name = list(filter(prof_dr_mr_filter, name.split()))
        if '|website' in name:
            name[2:] = []
        name = ' '.join(name)
        name = name.strip()
        student_name_list.append(name)

        # Getting the professor name
        try:
            if department == 'Chemical and Structural Biology':
                pi_name = stud.find('p', pi_class).text
            elif department == 'Molecular Genetics':
                pi_name = stud.find('div', class_= pi_class).text
            else:
                pi_name = None
                
            if pi_name != None:
                pi_name = pi_name.split()
                # Filtering the professor name from PI, Prof. and Dr
                pi_name = list(filter(prof_dr_mr_filter, pi_name))
                pi_name = ' '.join(pi_name)
                # Removing "lab" word in the name 
                pi_name = pi_name.split("'")[0]
                # Removing '()' from the names
                if '(' in pi_name:
                    pi_name = pi_name.split(' (')[0] + ' ' + pi_name.split(') ')[1]
            
            # Print student name and PI name
            print(f"Student: {name}, PI: {pi_name}")
            
            # Coupling of the student name and professor name
            couple = name, pi_name
            stud_pi_couple.append(couple)
        except:
            pi_name = 'None'
            print(f"Student: {name}, PI: None")

    return student_name_list, stud_pi_couple

# Function to search for publications by student name
def name_search_count(search_name):
    time.sleep(1)
    pubmed = PubMed(tool="PubMedSearcher", email="dubovetskiy@i.ua")
    results = pubmed.query(search_name, max_results=30)
    count = 0
    for article in results:
        count += 1
    return count
# Function to remove dublicates from student names
def remove_duplicate(name):
    try:
        # Split the string by commas
        name_parts = name.split(', ')
        
        # Split the first part into words
        first_part = name_parts[0].split()
        
        # Check if the second part contains any of the words in the first part
        second_part = name_parts[1].split()
        
        # Remove the duplicate words from the second part
        second_part_cleaned = [word for word in second_part if word not in first_part]
        
        # Combine the first part with the cleaned second part
        return ' '.join(first_part + second_part_cleaned)
    except: 
        return name

# Function to generate Excel with rankings and chart
def generate_excel_with_ranking(search_list):
    data_students_counts = []
    for search_name in search_list:
        count = name_search_count(search_name)

        if count == 0 and len(search_name.split()) == 3:
            search_name = search_name.split()
            del search_name[1]
            search_name = ' '.join(search_name)
            count = name_search_count(search_name)

        if count == 0 and len(search_name.split()) > 5:
            search_name = remove_duplicate(search_name)
            
        if count > 20:
            count = 0

        data_students_counts.append([search_name, count])
        print(f'search term: {search_name}, number of publications: {count}')
    
    sorted_data_students_counts = sorted(data_students_counts, key=lambda x: x[1], reverse=True)
    try:
        workbook = xlsxwriter.Workbook('students_rating.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write_column('A1', [row[0] for row in sorted_data_students_counts])
        worksheet.write_column('B1', [row[1] for row in sorted_data_students_counts])

        chart = workbook.add_chart({'type': 'column'})
        data_len = len(sorted_data_students_counts)
        chart.add_series({
            'name': "Publication Ranking",
            'categories': f'=Sheet1!$A$1:$A${data_len}',
            'values': f'=Sheet1!$B$1:$B${data_len}',
        })

        chart.set_title({'name': "Student's Publication Ranking"})
        chart.set_x_axis({'name': 'Students Names'})
        chart.set_y_axis({'name': 'Number of Publications'})
        worksheet.insert_chart('D2', chart)
        workbook.close()
        print('"students_rating.xlsx" file was generated in the current directory')
    except:
        print("!Please close the Excel file. The new file can't be generated when the old one is open!")
        quit()
# Callback function to update combobox based on selected department
def update_combobox(*args):
    department = clicked.get()
    if department == 'Systems Immunology':
        search_term['values'] = ['Total publications']
    else:
        search_term['values'] = ["Total publications", "Department publications"]
    
    if search_term['values']:
        search_term.current(0)

# Function for Ok button action
def on_ok_button_clicked():
    department = clicked.get()
    search_type = search_term.get()

    # Fetch student name and PI information
    student_name_list, stud_pi_couple = fetch_and_rank_students(department)
    
    # Prepare the search query based on selected search condition
    if student_name_list:
        if search_type == 'Total publications':
            # Search by student name only
            search_list = student_name_list
        elif search_type == 'Department publications':
            # Search by both student name and PI name
            search_list = [f"{student} {pi}" for student, pi in stud_pi_couple]
            

        # Generate the Excel with rankings and chart based on the search query
        generate_excel_with_ranking(search_list)
        result_label.config(text=f"{department}: Rankings and chart generated!")
    else:
        result_label.config(text="Please select a valid department.")

# GUI cancel button function
def btnCancel_clicked():
    window.destroy()

# GUI setup
window = tk.Tk()
window.title('PubRanker')
window.geometry('350x240')
window.tk.call('tk', 'scaling', 2.0)

database_frame = tk.LabelFrame(window, text='Choose the department', padx=5, pady=5)
database_frame.pack(padx=10, pady=10)

# Select department menu
clicked = tk.StringVar()
clicked.set('Chemical and Structural Biology')
drop_database = tk.OptionMenu(database_frame, clicked, 'Chemical and Structural Biology', 'Molecular Genetics', 'Systems Immunology')
drop_database.grid(column=0, row=0)

clicked.trace('w', update_combobox)

# Search term entry
search_lbl = tk.Label(window, text="Conditions for search")
search_lbl.pack()
search_term = ttk.Combobox(window)
search_term.pack()

update_combobox()

# Ok and cancel buttons
buttons_frame = tk.LabelFrame(window, padx=5, pady=5)
buttons_frame.pack(padx=10, pady=10)

btnOk = tk.Button(buttons_frame, text='Ok', command=on_ok_button_clicked)  
btnOk.grid(column=0, row=0)
btnCancel = tk.Button(buttons_frame, text='Cancel', command=btnCancel_clicked)
btnCancel.grid(column=2, row=0)

# Label to show result
result_label = tk.Label(window, text="")
result_label.pack(pady=10)

window.mainloop()
