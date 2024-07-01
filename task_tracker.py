import sqlite3
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pandas as pd

def connection():
    conn = sqlite3.connect('tasks_db.sqlite')
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_tasks (
            TaskID INTEGER PRIMARY KEY AUTOINCREMENT,
            TaskName TEXT NOT NULL,
            TaskDescription TEXT,
            TaskStatus TEXT,
            TaskDate DATE
        )
    ''')

    conn.commit()
    return conn

# Function to add a new task
def add_task():
    task_name = task_name_entry.get()
    task_description = task_description_entry.get()
    task_status = task_status_combobox.get()
    task_date = datetime.now().date()  # Current date as task date

    if not task_name.strip():
        messagebox.showinfo("Error", "Task Name cannot be empty.")
        return

    try:
        conn = connection()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO daily_tasks (TaskName, TaskDescription, TaskStatus, TaskDate) VALUES (?,?,?,?)",
                       (task_name, task_description, task_status, task_date))
        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Task added successfully")
        reset_fields()
        refresh_tasks()
    except sqlite3.Error as e:
        messagebox.showinfo("Error", f"Error occurred: {e}")

# Function to delete selected task
def delete_task():
    try:
        selected_item = tasks_treeview.selection()[0]
        task_id = tasks_treeview.item(selected_item)['values'][0]

        conn = connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM daily_tasks WHERE TaskID=?", (task_id,))
        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Task deleted successfully")
        refresh_tasks()
    except IndexError:
        messagebox.showinfo("Error", "Please select a task from the list.")
    except sqlite3.Error as e:
        messagebox.showinfo("Error", f"Error occurred: {e}")

# Function to update selected task
def update_task():
    try:
        selected_item = tasks_treeview.selection()[0]
        task_id = tasks_treeview.item(selected_item)['values'][0]

        task_name = task_name_entry.get()
        task_description = task_description_entry.get()
        task_status = task_status_combobox.get()

        conn = connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE daily_tasks SET TaskName=?, TaskDescription=?, TaskStatus=? WHERE TaskID=?",
                       (task_name, task_description, task_status, task_id))
        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Task updated successfully")
        reset_fields()
        refresh_tasks()
    except IndexError:
        messagebox.showinfo("Error", "Please select a task from the list.")
    except sqlite3.Error as e:
        messagebox.showinfo("Error", f"Error occurred: {e}")

# Function to search tasks based on keyword
def search_tasks():
    keyword = search_entry.get().strip()
    if not keyword:
        messagebox.showinfo("Error", "Please enter a keyword to search.")
        return

    try:
        conn = connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM daily_tasks WHERE TaskName LIKE ? OR TaskDescription LIKE ?", ('%'+keyword+'%', '%'+keyword+'%'))
        tasks = cursor.fetchall()
        conn.close()

        # Clear existing items in Treeview
        for data in tasks_treeview.get_children():
            tasks_treeview.delete(data)

        # Insert searched tasks into Treeview
        for task in tasks:
            tasks_treeview.insert(parent='', index='end', values=task)

    except sqlite3.Error as e:
        messagebox.showinfo("Error", f"Error occurred: {e}")

# Function to refresh the task list
def refresh_tasks():
    for data in tasks_treeview.get_children():
        tasks_treeview.delete(data)

    conn = connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM daily_tasks")
    tasks = cursor.fetchall()
    conn.close()

    for task in tasks:
        tasks_treeview.insert(parent='', index='end', values=task)

# Function to reset input fields
def reset_fields():
    task_name_entry.delete(0, 'end')
    task_description_entry.delete(0, 'end')
    task_status_combobox.set('Pending')  # Default status

# Function to reset all fields and search
def reset_all():
    reset_fields()
    refresh_tasks()
    search_entry.delete(0, 'end')  # Clear search entry

# Function to sort treeview columns
def treeview_sort_column(tv, col, reverse):
    data = [(tv.set(child, col), child) for child in tv.get_children('')]
    data.sort(reverse=reverse)

    for index, (val, child) in enumerate(data):
        tv.move(child, '', index)

    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

# Function to handle task selection from Treeview
def on_task_select(event):
    try:
        selected_item = tasks_treeview.selection()[0]
        task_details = tasks_treeview.item(selected_item, 'values')
        task_name_entry.delete(0, 'end')
        task_name_entry.insert(0, task_details[1])
        task_description_entry.delete(0, 'end')
        task_description_entry.insert(0, task_details[2])
        task_status_combobox.set(task_details[3])
    except IndexError:
        pass  # Ignore if no item is selected

# Function to export tasks to Excel
def export_to_excel():
    try:
        conn = connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM daily_tasks")
        tasks = cursor.fetchall()
        conn.close()

        # Convert fetched data to DataFrame
        df_tasks = pd.DataFrame(tasks, columns=["TaskID", "TaskName", "TaskDescription", "TaskStatus", "TaskDate"])

        # Ask user for filename and location to save the Excel file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        
        if file_path:
            df_tasks.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Data exported to Excel successfully")
    except sqlite3.Error as e:
        messagebox.showinfo("Error", f"Error occurred: {e}")

# Create main Tkinter window
root = Tk()
root.title("Daily Task Tracker")

# Fix window size
root.resizable(False, False)  # Fixed width and height

# Frame for task details
task_frame = Frame(root)
task_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

# Labels and Entries for task details
task_name_label = Label(task_frame, text="Task Name")
task_name_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')
task_name_entry = Entry(task_frame, width=40)
task_name_entry.grid(row=0, column=1, padx=10, pady=10, sticky='ew')

task_description_label = Label(task_frame, text="Task Description")
task_description_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')
task_description_entry = Entry(task_frame, width=40)
task_description_entry.grid(row=1, column=1, padx=10, pady=10, sticky='ew')

task_status_label = Label(task_frame, text="Task Status")
task_status_label.grid(row=2, column=0, padx=10, pady=10, sticky='w')
task_status_combobox = ttk.Combobox(task_frame, values=["Pending", "In Progress", "Completed"])
task_status_combobox.set("Pending")
task_status_combobox.grid(row=2, column=1, padx=10, pady=10, sticky='ew')

# Search Entry and Button
search_frame = Frame(root)
search_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=10)

search_label = Label(search_frame, text="Search:")
search_label.grid(row=0, column=0, padx=5, pady=5)

search_entry = Entry(search_frame, width=30)
search_entry.grid(row=0, column=1, padx=5, pady=5)

search_button = Button(search_frame, text="Search", command=search_tasks)
search_button.grid(row=0, column=2, padx=5, pady=5)

# Buttons for task operations
button_frame = Frame(root)
button_frame.grid(row=2, column=0, sticky='nsew', padx=10, pady=10)

add_task_button = Button(button_frame, text="Add Task", command=add_task)
add_task_button.grid(row=0, column=0, padx=5, pady=10)

update_task_button = Button(button_frame, text="Update Task", command=update_task)
update_task_button.grid(row=0, column=1, padx=5, pady=10)

delete_task_button = Button(button_frame, text="Delete Task", command=delete_task)
delete_task_button.grid(row=0, column=2, padx=5, pady=10)

reset_all_button = Button(button_frame, text="Reset All", command=reset_all)
reset_all_button.grid(row=0, column=3, padx=5, pady=10)

export_button = Button(button_frame, text="Export to Excel", command=export_to_excel)
export_button.grid(row=0, column=4, padx=5, pady=10)

# Treeview to display tasks
tree_frame = Frame(root)
tree_frame.grid(row=3, column=0, sticky='nsew', padx=10, pady=10)

tasks_treeview = ttk.Treeview(tree_frame, columns=("TaskID", "TaskName", "TaskDescription", "TaskStatus", "TaskDate"), show="headings")
tasks_treeview.heading("TaskID", text="TaskID", command=lambda: treeview_sort_column(tasks_treeview, "TaskID", False))
tasks_treeview.heading("TaskName", text="Task Name", command=lambda: treeview_sort_column(tasks_treeview, "TaskName", False))
tasks_treeview.heading("TaskDescription", text="Task Description", command=lambda: treeview_sort_column(tasks_treeview, "TaskDescription", False))
tasks_treeview.heading("TaskStatus", text="Task Status", command=lambda: treeview_sort_column(tasks_treeview, "TaskStatus", False))
tasks_treeview.heading("TaskDate", text="Task Date", command=lambda: treeview_sort_column(tasks_treeview, "TaskDate", False))

tasks_treeview.column("TaskID", width=50, anchor='center')
tasks_treeview.column("TaskName", width=150, anchor='center')
tasks_treeview.column("TaskDescription", width=200, anchor='center')
tasks_treeview.column("TaskStatus", width=100, anchor='center')
tasks_treeview.column("TaskDate", width=100, anchor='center')

tasks_treeview.grid(row=0, column=0, sticky='nsew')

tree_scroll = Scrollbar(tree_frame, orient="vertical", command=tasks_treeview.yview)
tree_scroll.grid(row=0, column=1, sticky='ns')
tasks_treeview.configure(yscrollcommand=tree_scroll.set)

# Bind the selection event of Treeview
tasks_treeview.bind("<ButtonRelease-1>", on_task_select)

# Initial task list display
refresh_tasks()

# Configure weight for resizing
root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(0, weight=1)

# Run the application
root.mainloop()
