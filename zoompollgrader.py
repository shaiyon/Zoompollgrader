import tkinter as tk
from tkinter.filedialog import askopenfilename, askdirectory
import pandas as pd
import numpy as np
from tooltip import CreateToolTip
import os
import sys

# Global variables
questions = []
widgets_onscreen = []
answers = None
answer_dicts = []
question_num = 0
students = None
totals = ["Total points"]

root = tk.Tk(className="zoom poll simplifier")
# Window settings
root.eval('tk::PlaceWindow . center')
if getattr(sys, 'frozen', False):
    icon = tk.PhotoImage(file=os.path.join(sys._MEIPASS, "icons8-poll-32.png"))
else:
    icon = tk.PhotoImage(file="icons8-poll-32.png")
root.wm_iconphoto(True, icon)
root.geometry("800x500")

def on_startup(root=root):
    for widget in widgets_onscreen:
        widget.destroy()
    widgets_onscreen.clear()
    # Hello label (in 104 languages!)
    if getattr(sys, 'frozen', False):
        hello = pd.read_csv(os.path.join(sys._MEIPASS, "hello.txt"), delimiter="\t", header=0).sample()
    else:
        hello = pd.read_csv("hello.txt", delimiter="\t", header=0).sample()
    hello_label = tk.Label(root, text=hello["hello"].item(), font=("Helvetica", 26))
    CreateToolTip(hello_label, text="hello in "+hello["language"].item())
    hello_label.pack(side="top", pady=80)
    widgets_onscreen.append(hello_label)

    # Choose file button
    def choose_file():
        filename = ""
        filename = askopenfilename()
        # Only proceed if file was selected
        if filename!="":
            on_file_selected(root, filename)
    choose_file_button = tk.Button(root, text="select poll report file", font=("Helvetica", 22),  command=choose_file)
    choose_file_button.pack()
    widgets_onscreen.append(choose_file_button)

    # Author label
    author_label = tk.Label(root, text="developed by shaiyon", font=("Helvetica", 8))
    author_label.pack(side="bottom", pady=10)
    widgets_onscreen.append(author_label)
    root.mainloop()

def on_file_selected(root, filename):
    try:
        try:
            df = pd.read_excel(filename, index_col=0, header=5)
        except:
            df = pd.read_csv(filename, index_col=0, header=5)
    except Exception as excpt:
        for widget in widgets_onscreen:
            widget.destroy()
        exception_label = tk.Label(root, text="Error loading file:\n"+str(excpt), font=("Helvetica", 20), wraplength=800, justify="center")
        exception_label.pack(pady=30)
        widgets_onscreen.append(exception_label)
        return_button = tk.Button(root, text="return to file selection", font=("Helvetica", 22),  command=on_startup)
        return_button.pack()
        widgets_onscreen.append(return_button)

    global students
    students = df.iloc[:,0].unique()
    students.sort()
    students = pd.DataFrame({'students':students})

    # Separate into list of dataframes, one for each question.
    split_index = [0]
    for row_index in range(df.shape[0]-1):
        if df.iloc[row_index, 3] != df.iloc[row_index+1, 3]:
            split_index.append(row_index+1)
    split_index.append(len(df))
                
    for i in range(1, len(split_index)):
        questions.append(df.iloc[split_index[i-1]:split_index[i], :])

    evaluate_question()

def evaluate_question(root=root, questions=questions):
    global question_num
    global widgets_onscreen
    global answers
    question_num = question_num + 1
    for widget in widgets_onscreen:
        widget.destroy()
    widgets_onscreen.clear()
    # When past final question, evaluate results.
    try:
        question = questions[question_num-1]
    except:
        on_output()

    answers = question.iloc[:,4].unique()
    answers.sort()

    title_label = tk.Label(root, text="Question "+str(question_num)+" of "+str(len(questions)), font=("Helvetica", 20), wraplength=400, justify="center")
    title_label.grid(row=0, column=0, padx=10)
    widgets_onscreen.append(title_label)

    question_label = tk.Label(root, text=question.iloc[0, 3], font=("Helvetica", 16), wraplength=400, justify="left")
    question_label.grid(row=0, column=1, pady=15)
    widgets_onscreen.append(question_label)

    ans_title_label = tk.Label(root, text="Answers:", font=("Helvetica", 14), justify="left")
    ans_title_label.grid(row=1, column=0, pady=7.5)
    ans_points_label = tk.Label(root, text="Points awarded for answer:", font=("Helvetica", 14), justify="right")
    ans_points_label.grid(row=1, column=1, pady=7.5)
    widgets_onscreen.append(ans_title_label)
    widgets_onscreen.append(ans_points_label)

    for i, answer in enumerate(answers):
        ans_label = tk.Label(root, text=answer, font=("Helvetica", 14), wraplength=400, justify="left")
        ans_label.grid(row=i+2, column=0)
        ans_entry = tk.Spinbox(root, from_=0, to=100, justify="center", font=("Helvetica", 10)) 
        ans_entry.grid(row=i+2, column=1)
        widgets_onscreen.append(ans_label)
        widgets_onscreen.append(ans_entry)

    total_points_label = tk.Label(root, text="Total question points:", font=("Helvetica", 14), justify="left")
    total_points_label.grid(row=len(answers)+2, column=0, pady=10)
    total_points_entry = tk.Spinbox(root, from_=1, to=100, justify="center", font=("Helvetica", 10)) 
    total_points_entry.grid(row=len(answers)+2, column=1)
    widgets_onscreen.append(total_points_label)
    widgets_onscreen.append(total_points_entry)

    next_button = tk.Button(root, text="next question", font=("Helvetica", 9), command=evaluate_next_question)
    next_button.grid(row=len(answers)+3, column=1)
    widgets_onscreen.append(next_button)

    root.mainloop()

def evaluate_next_question(root=root, questions=questions):
    global question_num
    global widgets_onscreen
    global answers
    global answer_dicts
    global totals
    counter = 0
    answer_dicts.append({})
    for widget in widgets_onscreen:
        if isinstance(widget, tk.Spinbox):
            try:
                answer_dicts[question_num-1][answers[counter]] = widget.get()
            except:
                totals.append(widget.get())
            counter = counter + 1
    evaluate_question()

def on_output():
        global answer_dicts
        global answers
        global students
        global questions
        global totals

        for widget in widgets_onscreen:
            widget.destroy()

        for i, question in enumerate(questions):
            questions[i][questions[i].iloc[0,3]] = question.iloc[:,4].map(answer_dicts[i])

        for question in questions:
            students = students.merge(question[["User Name", question.iloc[0,3]]], how="left", left_on="students", right_on="User Name")
        students = students.loc[:,~students.columns.str.startswith('User Name')]
        students.loc[len(students.index),:] = totals
        students = students.set_index("students")
        students = students.astype(float)

        students["Total"] = students.sum(axis=1, skipna=True)
        students["Percentage"] = students["Total"] / students["Total"].iloc[[-1]].item()

        def choose_directory():
            outputdir = ""
            outputdir = askdirectory()
            # Only proceed if file was selected
            if outputdir!="":
                outputname = outputname_entry.get()
                students.to_csv(outputdir+"/"+outputname+".csv", index=True)
                exit()

        outputname_label = tk.Label(root, text="type output file name (no file extension)", font=("Helvetica", 18))
        outputname_label.pack()
        outputname_entry = tk.Entry(root, font=("Helvetica", 18))
        outputname_entry.pack()

        choose_output_button = tk.Button(root, text="select output directory", font=("Helvetica", 22),  command=choose_directory)
        choose_output_button.pack()
        widgets_onscreen.append(choose_output_button)
    
        root.mainloop()

if __name__=="__main__":
    on_startup()