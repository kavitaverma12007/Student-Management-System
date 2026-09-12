# Student-Management-System
import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl
import os

EXCEL_FILE = "Student_results.xlsx"

window = tk.Tk()
window.title("Student Details")
window.geometry("650x650")

if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append([
        "Name",
        "Roll No",
        "Class",
        "Subject 1",
        "Subject 2",
        "Subject 3",
        "Subject 4",
        "Subject 5",
        "Total Marks",
        "Percentage",
        "Result"
    ])

    wb.save(EXCEL_FILE)

def add_student():

    for widget in window.winfo_children():
        widget.destroy()

    tk.Label( window,text="Student Management System",font=("Arial", 24, "bold")
    ).grid(row=0, column=0, columnspan=4, pady=20)

    tk.Label(window,text="Name"
    ).grid(row=1, column=0, padx=10, pady=10)

    name = tk.Entry(window, width=30)
    name.grid(row=1, column=1)

    tk.Label(
        window,
        text="Roll No"
    ).grid(row=1, column=2, padx=10, pady=10)

    roll_no = tk.Entry(window, width=30)
    roll_no.grid(row=1, column=3)

    tk.Label(window,text="Class").grid(row=2, column=0, padx=10, pady=10)

    class_ = tk.Entry(window, width=30)
    class_.grid(row=2, column=1)


    marks = []

    for i in range(5):

        tk.Label( window, text=f"Subject {i + 1} Marks").grid(row=3 + i,
            column=0, padx=10, pady=8 )

        mark_entry = tk.Entry(window, width=30)
        mark_entry.grid(row=3 + i,
            column=1,padx=10,pady=8)

        marks.append(mark_entry)

    def save():

        Name = name.get()
        Roll_no = roll_no.get()
        Class = class_.get()

        
        if Name == "" or Roll_no == "" or Class == "":
            messagebox.showerror(
                "Error",
                "Please enter Name, Roll No and Class."
            )
            return

        try:
            Marks = [int(entry.get()) for entry in marks]

        except ValueError:
            messagebox.showerror(
                "Error",
                "Please enter valid marks."
            )
            return


        if any(mark < 0 or mark > 100 for mark in Marks):
            messagebox.showerror(
                "Error",
                "Enter Valid Marks (0-100)"
            )
            return

        total = sum(Marks)
        percentage = total / 5
        result = "Pass" if all(
            mark >= 35 for mark in Marks
        ) else "Fail"

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active

        ws.append([
            Name,
            Roll_no,
            Class,
            *Marks,
            total,
            percentage,
            result
        ])

        wb.save(EXCEL_FILE)

        messagebox.showinfo(
            "Successful",
            "Record Saved!"
        )

    tk.Button( window,text="Save",command=save).grid( row=8,
        column=0, padx=10, pady=15)

    tk.Button(window, text="Back", command=main_menu ).grid(row=8,
        column=1, padx=10, pady=15)

def get_result():

    for widget in window.winfo_children():
        widget.destroy()

    tk.Label( window, text="Enter Roll No" ).grid(row=0,column=0,padx=10,
        pady=10 )

    roll_entry = tk.Entry(window, width=30)
    roll_entry.grid(
        row=0,
        column=1,
        padx=10,
        pady=10
    )

    result_columns = (
        "Name",
        "Roll No",
        "Class",
        "Total Marks",
        "Percentage",
        "Result"
    )

    result_tree = ttk.Treeview(
        window,
        columns=result_columns,
        show="headings",
        height=8
    )

    for column in result_columns:

        result_tree.heading(
            column,
            text=column
        )

        result_tree.column(
            column,
            width=100,
            anchor="center"
        )

    result_tree.grid(
        row=2,
        column=0,
        columnspan=3,
        padx=10,
        pady=20
    )


    def search_result():

        enter_roll_no = roll_entry.get().strip()

        if enter_roll_no == "":
            messagebox.showerror(
                "Error",
                "Enter Roll No"
            )
            return

        for item in result_tree.get_children():
            result_tree.delete(item)

        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active

        found = False

        for row in ws.iter_rows(
            min_row=2,
            values_only=True
        ):

            if str(row[1]) == enter_roll_no:

                result_tree.insert(
                    "",
                    "end",
                    values=(
                        row[0],
                        row[1],
                        row[2],
                        row[8],
                        row[9],
                        row[10]
                    )
                )

                found = True
                break

        if not found:
            messagebox.showinfo(
                "Result",
                "Student record not found."
            )

    tk.Button(window, text="Get Result", command=search_result
    ).grid( row=0, column=2, padx=10, pady=10 )

    tk.Button( window,text="Back",command=main_menu).grid(row=3,
        column=0, pady=15 )

def show_all_results():
    for widget in window.winfo_children():
        widget.destroy()

    result_columns = (
        "Name",
        "Roll No",
        "Class",
        "Total Marks",
        "Percentage",
        "Result"
    )

    tk.Label(
        window,
        text="All Student Results",
        font=("Arial", 22, "bold")
    ).grid(
        row=0,
        column=0,
        columnspan=6,
        pady=20
    )

    all_result_tree = ttk.Treeview(
        window,
        columns=result_columns,
        show="headings",
        height=18
    )

    for column in result_columns:

        all_result_tree.heading(
            column,
            text=column
        )

        all_result_tree.column(
            column,
            width=100,
            anchor="center"
        )

    all_result_tree.grid(
        row=1,
        column=0,
        columnspan=6,
        padx=10,
        pady=10
    )

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active


    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        all_result_tree.insert(
            "",
            "end",
            values=(
                row[0],
                row[1],
                row[2],
                row[8],
                row[9],
                row[10]
            )
        )

    tk.Button( window, text="Back", command=main_menu ).grid(
        row=2,column=0, pady=15 )


def exit_app():

    if messagebox.askyesno("Exit",
        "Are you sure you want to exit?"):
        window.destroy()


def main_menu():

    for widget in window.winfo_children():
        widget.destroy()

    tk.Label(window, text="Student Management System",
     font=("Arial", 24, "bold") ).pack(pady=40)

    tk.Button(window,text="Add Student",command=add_student,
              width=25 ).pack(pady=10)

    tk.Button(window,text="Get Result",command=get_result,
        width=25).pack(pady=10)

    tk.Button(window, text="Show All Results", command=show_all_results,
     width=25).pack(pady=10)

    tk.Button( window,text="Exit", command=exit_app,
        width=25).pack(pady=10)


main_menu()

window.mainloop()
