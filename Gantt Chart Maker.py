import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import webbrowser
import plotly.express as px
from tkcalendar import DateEntry

def generate_chart():
    tasks = []
    tooltips = []
    for row in tree.get_children():
        values = tree.item(row)['values']
        if len(values) == 5:
            task_name, start_date, end_date, status, remark = values
        else:
            task_name, start_date, end_date, status = values
            remark = ""
        start_dt = pd.to_datetime(start_date)
        end_dt = pd.to_datetime(end_date)
        if end_dt == start_dt:
            display_end_dt = start_dt + pd.Timedelta(days=2)
        else:
            display_end_dt = end_dt
        # Format remark tooltip as bold and red HTML
        formatted_remark = f"<span style='color:#d32f2f;font-weight:bold'>{remark}</span>" if remark else ""
        tasks.append({
            'Task': task_name,
            'Start Date': start_dt,
            'End Date': end_dt,
            'Display End Date': display_end_dt,
            'Status': status,
            'Remark': formatted_remark
        })
        tooltips.append(formatted_remark)

    df = pd.DataFrame(tasks)
    if df.empty:
        messagebox.showwarning("No Data", "No tasks to display.")
        return

    chart_title = entry_title.get()  

    color_map = {
        "In Progress": "#F8E960",
        "Completed": "#05CF56",
        "To Do": "#D8D1D1",
        "Target Go Live": "#F39405"  
    }

    df['Label'] = df.apply(
        lambda row: f"<b>{row['Task']}</b> | {row['Start Date'].strftime('%Y-%m-%d')} - {row['End Date'].strftime('%Y-%m-%d')}",
        axis=1
    )

    df = df.iloc[::-1].reset_index(drop=True)

    fig = px.timeline(
        df,
        x_start="Start Date",
        x_end="Display End Date",
        y="Label",
        color="Status",
        color_discrete_map=color_map,
        title=chart_title,  
        labels={"Label": "Task", "Start Date": "Start Date", "End Date": "End Date", "Status": "Status"},
        hover_data={"Remark": True}
    )

    fig.update_yaxes(
        categoryorder="array",
        categoryarray=df['Label'].tolist(),
        title_text="Task | Start Date ~ End Date",
        showgrid=True
    )
    min_date = df['Start Date'].min()
    first_month = pd.Timestamp(year=min_date.year, month=min_date.month, day=1)
    max_date = df['Display End Date'].max()
    last_month = pd.Timestamp(year=max_date.year, month=max_date.month, day=1) + pd.offsets.MonthEnd(1)
    fig.update_xaxes(
        title_text="Date",
        tickformat="%b-%Y",
        showgrid=True,
        dtick="M1",
        ticklabelmode="period",
        tick0=first_month.strftime('%Y-%m-%d'),
        range=[first_month, last_month]
    )
    fig.update_layout(
        xaxis=dict(
            tickangle=0,
            tickfont=dict(size=12),
            type="date"
        ),
        legend_title_text="Task Status",
        bargap=0.5,
        height=600,
        margin=dict(l=250, r=40, t=60, b=60),
        plot_bgcolor="#fff"
    )

    # Remove text inside bars and enlarge "Target Go Live" bar (no border)
    for i, d in enumerate(fig.data):
        d.text = ""
        d.textposition = "none"
        if d.name == "Target Go Live":
            d.width = 1.2  # Make the bar thicker
            d.marker.line.width = 0  # Remove border

    fig.update_layout(
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=14)
        )
    )

    fig.write_html("gantt_chart.html")
    webbrowser.open("gantt_chart.html")

#Treeview
def add_task():
    task_name = entry_task.get()
    start_date = entry_start.get()
    end_date = entry_end.get()
    status = status_var.get()
    remark = entry_remark.get()

    if not task_name or not start_date or not end_date or not status:
        messagebox.showerror("Input Error", "Please fill in all fields.")
        return

    tree.insert("", "end", values=(task_name, start_date, end_date, status, remark))
    entry_task.delete(0, tk.END)
    entry_start.delete(0, tk.END)
    entry_end.delete(0, tk.END)
    status_var.set("To Do")
    entry_remark.delete(0, tk.END)

# Function to remove selected task from the treeview
def remove_task():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a task to remove.")
        return
    for item in selected_item:
        tree.delete(item)

def import_from_csv():
    file_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    if not file_path:
        return
    try:
        df = pd.read_csv(file_path)
        required_cols = ["Task", "Start_Date", "End_Date", "Status", "Remark"]
        for col in required_cols:
            if col not in df.columns:
                messagebox.showerror("Import Error", f"Missing column: {col}")
                return
        tree.delete(*tree.get_children())
        for idx, row in df.iterrows():
            task = str(row["Task"])
            start = str(row["Start_Date"]) if pd.notnull(row["Start_Date"]) else ""
            end = str(row["End_Date"]) if pd.notnull(row["End_Date"]) else ""
            status = str(row["Status"])
            remark = str(row["Remark"]) if pd.notnull(row["Remark"]) else ""
            if task and start and end and status:
                tree.insert("", "end", values=(task, start, end, status, remark))
        messagebox.showinfo("Import Successful", "Tasks imported from CSV.")
    except Exception as e:
        messagebox.showerror("Import Error", f"Failed to import: {e}")

def edit_task():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a task to edit.")
        return
    item = selected[0]
    values = tree.item(item, "values")
    entry_task.delete(0, tk.END)
    entry_task.insert(0, values[0])
    entry_start.set_date(values[1])
    entry_end.set_date(values[2])
    status_var.set(values[3])
    if len(values) > 4:
        entry_remark.delete(0, tk.END)
        entry_remark.insert(0, values[4])
    else:
        entry_remark.delete(0, tk.END)
    tree.delete(item)

def move_task_up():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a task to move.")
        return
    item = selected[0]
    index = tree.index(item)
    if index == 0:
        return  
    above_item = tree.get_children()[index - 1]
    tree.move(item, '', index - 1)
    tree.selection_set(item)

def move_task_down():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a task to move.")
        return
    item = selected[0]
    index = tree.index(item)
    children = tree.get_children()
    if index == len(children) - 1:
        return  
    tree.move(item, '', index + 1)
    tree.selection_set(item)

def export_to_csv():
    file_path = filedialog.asksaveasfilename(
        title="Save CSV File",
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv")]
    )
    if not file_path:
        return
    tasks = []
    for row in tree.get_children():
        values = tree.item(row)['values']
        if len(values) == 5:
            task_name, start_date, end_date, status, remark = values
        else:
            task_name, start_date, end_date, status = values
            remark = ""
        tasks.append({
            "Task": task_name,
            "Start_Date": start_date,
            "End_Date": end_date,
            "Status": status,
            "Remark": remark
        })
    df = pd.DataFrame(tasks)
    df.to_csv(file_path, index=False)
    messagebox.showinfo("Export Successful", "Tasks exported to CSV.")

root = tk.Tk()
root.title("Gantt Chart Generator")


frame_input = tk.Frame(root)
frame_input.pack(pady=10, anchor="w")  

tk.Label(frame_input, text="Chart Title").grid(row=0, column=0, sticky="w")
entry_title = tk.Entry(frame_input)
entry_title.insert(0, "Project Gantt Chart")  
entry_title.grid(row=0, column=1, sticky="w")

tk.Label(frame_input, text="Task Name").grid(row=1, column=0, sticky="w")
entry_task = tk.Entry(frame_input)
entry_task.grid(row=1, column=1, sticky="w")

tk.Label(frame_input, text="Start Date").grid(row=2, column=0, sticky="w")
entry_start = DateEntry(frame_input, date_pattern='yyyy-mm-dd')
entry_start.grid(row=2, column=1, sticky="w")

tk.Label(frame_input, text="End Date").grid(row=3, column=0, sticky="w")
entry_end = DateEntry(frame_input, date_pattern='yyyy-mm-dd')
entry_end.grid(row=3, column=1, sticky="w")

tk.Label(frame_input, text="Status").grid(row=4, column=0, sticky="w")
status_var = tk.StringVar(value="To Do")
status_options = ttk.Combobox(
    frame_input,
    textvariable=status_var,
    values=["To Do", "In Progress", "Completed", "Target Go Live"],
    state="readonly"
)
status_options.grid(row=4, column=1, sticky="w")

tk.Label(frame_input, text="Remark").grid(row=5, column=0, sticky="w")
entry_remark = tk.Entry(frame_input)
entry_remark.grid(row=5, column=1, sticky="w")

btn_add = tk.Button(frame_input, text="Add Task", command=add_task)
btn_add.grid(row=6, column=0, pady=5, sticky="w")

btn_edit = tk.Button(frame_input, text="Edit Task", command=edit_task)
btn_edit.grid(row=6, column=1, pady=5, sticky="w")

# Treeview to display tasks
tree = ttk.Treeview(root, columns=("Task", "Start", "End", "Status", "Remark"), show="headings")
tree.heading("Task", text="Task")
tree.heading("Start", text="Start Date")
tree.heading("End", text="End Date")
tree.heading("Status", text="Status")
tree.heading("Remark", text="Remark")
tree.pack(pady=10, anchor="w")

# Button frame for chart generation and other actions
frame_buttons = tk.Frame(root)
frame_buttons.pack(fill="x", padx=20, pady=15)  

btn_generate = tk.Button(frame_buttons, text="Generate Gantt Chart", command=generate_chart)
btn_generate.pack(side="right", padx=15, pady=8)  

btn_remove = tk.Button(frame_buttons, text="Remove Selected Task", command=remove_task)
btn_remove.pack(side="left", padx=15, pady=8)  

btn_import_csv = tk.Button(frame_buttons, text="Import from CSV", command=import_from_csv)
btn_import_csv.pack(side="left", padx=15, pady=8)  

btn_export_csv = tk.Button(frame_buttons, text="Export to CSV", command=export_to_csv)
btn_export_csv.pack(side="left", padx=15, pady=8)

def on_treeview_button_press(event):
    tree._drag_data = tree.identify_row(event.y)
    tree._drag_target = None

def on_treeview_motion(event):
    row_under_cursor = tree.identify_row(event.y)
    drag_row = getattr(tree, "_drag_data", None)
    if drag_row and row_under_cursor and drag_row != row_under_cursor:
        tree._drag_target = row_under_cursor
        tree.selection_set(row_under_cursor)

def on_treeview_button_release(event):
    drag_row = getattr(tree, "_drag_data", None)
    target_row = getattr(tree, "_drag_target", None)
    if drag_row and target_row and drag_row != target_row:
        index = tree.index(target_row)
        tree.move(drag_row, '', index)
        tree.selection_set(drag_row)
    tree._drag_data = None
    tree._drag_target = None

tree.bind("<ButtonPress-1>", on_treeview_button_press)
tree.bind("<B1-Motion>", on_treeview_motion)
tree.bind("<ButtonRelease-1>", on_treeview_button_release)

root.mainloop()
