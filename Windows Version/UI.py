import tkinter as tk
from tkinter import ttk
import subprocess
import threading


def run_script_one():
    console.insert(tk.END, "Running Script 1...\n")
    console.see(tk.END)
    subprocess.run(["python", "Filter.py"])
    console.insert(tk.END, "Script 1 completed.\n")
    console.see(tk.END)

def run_script_two():
    console.insert(tk.END, "Running Script 2...\n")
    console.see(tk.END)
    subprocess.run(["python", "Populate.py"])
    console.insert(tk.END, "Script 2 completed.\n")
    console.see(tk.END)

def run_script_three():
    console.insert(tk.END, "Running Script 3...\n")
    console.see(tk.END)
    process = subprocess.Popen(["python", "NeonCRM_Tool.py"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, bufsize=1, universal_newlines=True)
    while True:
        output = process.stdout.readline()
        if output == '' and process.poll() is not None:
            break
        if output:
            console.insert(tk.END, output)
            console.see(tk.END)
    console.insert(tk.END, "Script 3 completed.\n")
    console.see(tk.END)

def thread_function(script_function):
    threading.Thread(target=script_function).start()

root = tk.Tk()
root.title('Script Runner')

console = tk.Text(root, wrap=tk.WORD)
console.pack(padx=10, pady=10)

frame = tk.Frame(root)
frame.pack(pady=0)

button1 = ttk.Button(frame, text='Run Script 1: Filtering Tool', command=lambda: thread_function(run_script_one))
button1.grid(row=0, column=0, padx=30, pady=10)
button1.config(width=20)

button2 = ttk.Button(frame, text='Run Script 2: Populating Tool', command=lambda: thread_function(run_script_two))
button2.grid(row=0, column=1, padx=30, pady=10)
button2.config(width=20)

button3 = ttk.Button(frame, text='Run Script 3: NeonCRM Tool', command=lambda: thread_function(run_script_three))
button3.grid(row=0, column=2, padx=30, pady=10)
button3.config(width=20)

root.mainloop()
