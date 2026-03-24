import json
import tkinter as tk
import os
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import datetime


# =========================
# ШАБЛОНЫ

templates = {
    "Типовой(А/Ц)": "template.docx",
    "Вн. работы": "template2.docx"
}


# =========================
# Тут загрузка и сохранение

def load_last():
    if os.path.exists("last_data.json"):
        with open("last_data.json", "r", encoding="utf8") as f:
            return json.load(f)
    return {}


def save_last(data):
    with open("last_data.json", "w", encoding="utf8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def load_json(file):
    with open(file, "r", encoding="utf8") as f:
        return json.load(f)


works = load_json("works.json")
employees = load_json("employees.json")["employees"]

employee_names = [e["name"] for e in employees]


# =========================
# допы

def get_post(name):
    for e in employees:
        if e["name"] == name:
            return e["post"]
    return ""


def refresh_employee_lists():
    global employee_names
    employee_names = [e["name"] for e in employees]

    for cb in [executor_cb, leader_cb, rb_cb, chief_cb, issued_cb, accepted_cb]:
        cb["values"] = employee_names


def auto_save():
    save_last({
        "template": template_var.get(),

        "executor": executor_cb.get(),
        "leader": leader_cb.get(),
        "rb": rb_cb.get(),
        "chief": chief_cb.get(),
        "issued": issued_cb.get(),
        "accepted": accepted_cb.get(),
        "team": team_cb.get(),
        "date": date_entry.get(),
        "jobs": jobs_cb.get(),
        "material": material_cb.get(),
        "protect": protect_cb.get(),
        "rboblast": rboblast_cb.get(),
        "controls": controls_cb.get(),
        "rbRules": rbRules_cb.get(),
        "workers": [w.get() for w in worker_rows if w.get()]
    })

def update_template_ui():
    selected = template_var.get()

    if selected == "Вн. работы":
        rboblast_cb.config(state="disabled")
    else:
        rboblast_cb.config(state="normal")
def safe_set(cb, value, options):
    if value in options:
        cb.set(value)
    elif options:
        cb.set(options[0])


# =========================
# все что связано с ворд

def generate():
    if not executor_cb.get():
        tk.messagebox.showerror("Ошибка", "Выберите исполнителя")
        return

    if not jobs_cb.get():
        tk.messagebox.showerror("Ошибка", "Выберите работу")
        return

    date = date_entry.get()
    dateDDMM = date[:5]
    dateYYYY = "20" + date[-2:]

    workers = []

    for w in worker_rows:
        name = w.get()
        if name:
            workers.append({
                "name": name,
                "post": get_post(name),
                "date": date
            })

    data = {
        "executor": executor_cb.get(),
        "team": team_cb.get(),

        "leader": leader_cb.get(),
        "leadePost": get_post(leader_cb.get()),

        "rb": rb_cb.get(),
        "rbPost": get_post(rb_cb.get()),

        "chief": chief_cb.get(),
        "chiefPost": get_post(chief_cb.get()),

        "issued": issued_cb.get(),
        "issuedPost": get_post(issued_cb.get()),

        "accepted": accepted_cb.get(),
        "acceptedPost": get_post(accepted_cb.get()),

        "jobs": jobs_cb.get(),
        "material": material_cb.get(),
        "protect": protect_cb.get(),
        "rboblast": rboblast_cb.get(),
        "controls": controls_cb.get(),
        "rbRules": rbRules_cb.get(),

        "data": date,
        "dataDDMM": dateDDMM,
        "dataYYYY": dateYYYY,

        "workers": workers
    }

    # ✅ ВЫБОР ШАБЛОНА
    template_name = template_var.get()
    template_path = templates.get(template_name, "template.docx")

    doc = DocxTemplate(template_path)
    doc.render(data)

    path = os.path.join(os.path.expanduser("~"), "Documents", "narad_ready.docx")
    doc.save(path)
    os.startfile(path)


# =========================
# данные (НЕ ТРОГАЛ)

def manage_data():

    win = tk.Toplevel(root)
    win.title("Управление данными")

    notebook = ttk.Notebook(win)
    notebook.pack(fill="both", expand=True)

    # ---------- Сотрудники ----------
    emp_frame = tk.Frame(notebook)
    notebook.add(emp_frame, text="Сотрудники")

    emp_list = tk.Listbox(emp_frame, width=50)
    emp_list.pack()

    def refresh_emp():
        emp_list.delete(0, tk.END)
        for e in employees:
            emp_list.insert(tk.END, f"{e['name']} — {e['post']}")

    refresh_emp()

    def delete_employee():
        sel = emp_list.curselection()
        if not sel:
            return

        index = sel[0]
        employees.pop(index)

        with open("employees.json", "w", encoding="utf8") as f:
            json.dump({"employees": employees}, f, ensure_ascii=False, indent=4)

        refresh_emp()
        refresh_employee_lists()
        auto_save()

    def edit_employee():
        sel = emp_list.curselection()
        if not sel:
            return

        index = sel[0]
        e = employees[index]

        edit = tk.Toplevel(win)

        tk.Label(edit, text="ФИО").grid(row=0)
        name = tk.Entry(edit)
        name.insert(0, e["name"])
        name.grid(row=0, column=1)

        tk.Label(edit, text="Должность").grid(row=1)
        post = tk.Entry(edit)
        post.insert(0, e["post"])
        post.grid(row=1, column=1)

        def save():
            employees[index]["name"] = name.get()
            employees[index]["post"] = post.get()

            with open("employees.json", "w", encoding="utf8") as f:
                json.dump({"employees": employees}, f, ensure_ascii=False, indent=4)

            refresh_emp()
            refresh_employee_lists()
            auto_save()
            edit.destroy()

        tk.Button(edit, text="Сохранить", command=save).grid(row=2, column=1)

    tk.Button(emp_frame, text="Изменить", command=edit_employee).pack()
    tk.Button(emp_frame, text="Удалить", command=delete_employee).pack()

    # ---------- Работы ----------
    work_frame = tk.Frame(notebook)
    notebook.add(work_frame, text="Работы")

    work_list = tk.Listbox(work_frame, width=70)
    work_list.pack()

    def refresh_work():
        work_list.delete(0, tk.END)
        for k, vals in works.items():
            for v in vals:
                work_list.insert(tk.END, f"{k} : {v}")

    refresh_work()

    def delete_work():
        sel = work_list.curselection()
        if not sel:
            return

        line = work_list.get(sel)
        key, text = line.split(" : ", 1)

        works[key].remove(text)

        with open("works.json", "w", encoding="utf8") as f:
            json.dump(works, f, ensure_ascii=False, indent=4)

        refresh_work()

        jobs_cb["values"] = works["jobs"]
        material_cb["values"] = works["material"]
        protect_cb["values"] = works["protect"]
        rboblast_cb["values"] = works["rboblast"]
        controls_cb["values"] = works["controls"]
        rbRules_cb["values"] = works["rbRules"]

        auto_save()

    tk.Button(work_frame, text="Удалить", command=delete_work).pack()


# =========================
# добавление данных (НЕ ТРОГАЛ)

def add_employee_window():

    win = tk.Toplevel(root)
    win.title("Добавить сотрудника")

    tk.Label(win, text="ФИО").grid(row=0)
    name = tk.Entry(win)
    name.grid(row=0, column=1)

    tk.Label(win, text="Должность").grid(row=1)
    post = tk.Entry(win)
    post.grid(row=1, column=1)

    def save():
        employees.append({"name": name.get(), "post": post.get()})

        with open("employees.json", "w", encoding="utf8") as f:
            json.dump({"employees": employees}, f, ensure_ascii=False, indent=4)

        refresh_employee_lists()
        auto_save()
        win.destroy()

    tk.Button(win, text="Сохранить", command=save).grid(row=2, column=1)


def add_work_window():

    win = tk.Toplevel(root)
    win.title("Добавить вариант")

    tk.Label(win, text="Тип").grid(row=0)
    type_cb = ttk.Combobox(win, values=list(works.keys()))
    type_cb.grid(row=0, column=1)

    tk.Label(win, text="Текст").grid(row=1)
    text = tk.Entry(win, width=50)
    text.grid(row=1, column=1)

    def save():
        t = type_cb.get()
        txt = text.get()

        if not t or not txt:
            return

        works[t].append(txt)

        with open("works.json", "w", encoding="utf8") as f:
            json.dump(works, f, ensure_ascii=False, indent=4)

        jobs_cb["values"] = works["jobs"]
        material_cb["values"] = works["material"]
        protect_cb["values"] = works["protect"]
        rboblast_cb["values"] = works["rboblast"]
        controls_cb["values"] = works["controls"]
        rbRules_cb["values"] = works["rbRules"]

        auto_save()
        win.destroy()

    tk.Button(win, text="Сохранить", command=save).grid(row=2, column=1)


# =========================
# UI

root = tk.Tk()


def paste(event):
    try:
        text = root.clipboard_get()
        widget = event.widget

        # вставка в позицию курсора
        if isinstance(widget, tk.Entry) or isinstance(widget, ttk.Combobox):
            widget.delete(0, tk.END)
            widget.insert(0, text)

    except:
        pass


def copy(event):
    try:
        widget = event.widget
        selection = widget.selection_get()
        root.clipboard_clear()
        root.clipboard_append(selection)
    except:
        pass


def cut(event):
    copy(event)
    try:
        event.widget.delete("sel.first", "sel.last")
    except:
        pass


# бинды
root.bind_all("<Control-v>", paste)
root.bind_all("<Control-c>", copy)
root.bind_all("<Control-x>", cut)
def paste(event):
    event.widget.event_generate("<<Paste>>")

root.bind_class("Entry", "<Button-3>", paste)
root.bind_class("TCombobox", "<Button-3>", paste)
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")

root.title("Наряд-допуск")

last = load_last()

frame = tk.Frame(root)
frame.pack(padx=20, pady=20)
# ===== ВЫБОР ШАБЛОНА (ВВЕРХУ И ВЫДЕЛЕН)

template_var = tk.StringVar()

template_frame = tk.LabelFrame(
    frame,
    text="Выбор шаблона",
    padx=10,
    pady=5,
    font=("Arial", 9, "bold")
)

template_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

for name in templates.keys():
    rb = tk.Radiobutton(
        template_frame,
        text=name,
        variable=template_var,
        value=name,
        command=lambda: [auto_save(), update_template_ui()]
    )
    rb.pack(side="left", padx=10)

def add_combo(label, values, row):
    tk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
    cb = ttk.Combobox(frame, values=values, width=35)
    cb.grid(row=row, column=1, padx=5, pady=3)
    return cb


def add_entry(label, row):
    tk.Label(frame, text=label).grid(row=row, column=0, sticky="w")
    entry = tk.Entry(frame, width=38)
    entry.grid(row=row, column=1)
    return entry


executor_cb = add_combo("Ответственный исполнитель", employee_names, 1)
leader_cb = add_combo("Руководитель работ", employee_names, 2)
rb_cb = add_combo("Ответственный РБ", employee_names, 3)
chief_cb = add_combo("Ответственный за работы", employee_names, 4)

issued_cb = add_combo("Наряд выдал", employee_names, 5)
accepted_cb = add_combo("Наряд принял", employee_names, 6)

tk.Label(frame, text="Бригада").grid(row=8)
team_cb = ttk.Combobox(frame, values=list(range(1, 11)))
team_cb.grid(row=8, column=1)

today = datetime.now().strftime("%d.%m.%y")
date_entry = add_entry("Дата", 9)
date_entry.insert(0, today)



jobs_cb = add_combo("Работы", works["jobs"], 10)
material_cb = add_combo("Материалы", works["material"], 11)
protect_cb = add_combo("СИЗ", works["protect"], 12)
rboblast_cb = add_combo("Рад. обстановка", works["rboblast"], 13)
controls_cb = add_combo("Контроль", works["controls"], 14)
rbRules_cb = add_combo("Инструкции РБ", works["rbRules"], 15)


template_var.set(last.get("template", list(templates.keys())[0]))
update_template_ui()
executor_cb.set(last.get("executor", ""))
leader_cb.set(last.get("leader", ""))
rb_cb.set(last.get("rb", ""))
chief_cb.set(last.get("chief", ""))
issued_cb.set(last.get("issued", ""))
accepted_cb.set(last.get("accepted", ""))
team_cb.set(last.get("team", 1))


safe_set(jobs_cb, last.get("jobs"), works["jobs"])
safe_set(material_cb, last.get("material"), works["material"])
safe_set(protect_cb, last.get("protect"), works["protect"])
safe_set(rboblast_cb, last.get("rboblast"), works["rboblast"])
safe_set(controls_cb, last.get("controls"), works["controls"])
safe_set(rbRules_cb, last.get("rbRules"), works["rbRules"])


worker_rows = []


def add_worker(name=""):
    r = len(worker_rows) + 16
    cb = ttk.Combobox(frame, values=employee_names)
    cb.grid(row=r, column=0, columnspan=2, pady=2)
    if name:
        cb.set(name)
    worker_rows.append(cb)


def remove_worker():
    if worker_rows:
        worker_rows.pop().destroy()
        auto_save()


for w in last.get("workers", []):
    add_worker(w)


tk.Button(frame, text="Добавить исполнителя", command=add_worker).grid(row=16, column=2)
tk.Button(frame, text="Удалить исполнителя", command=remove_worker).grid(row=17, column=2)

tk.Button(frame, text="Добавить сотрудника", command=add_employee_window).grid(row=30, column=0)
tk.Button(frame, text="Управление данными", command=manage_data).grid(row=30, column=1)
tk.Button(frame, text="Добавить вариант", command=add_work_window).grid(row=30, column=2)

tk.Button(frame, text="Создать наряд", command=generate, bg="green", fg="white").grid(row=31, column=1)


widgets = [
    executor_cb, leader_cb, rb_cb, chief_cb,
    issued_cb, accepted_cb,
    team_cb, date_entry,
    jobs_cb, material_cb, protect_cb,
    rboblast_cb, controls_cb,
    rbRules_cb
]

for w in widgets:
    w.bind("<<ComboboxSelected>>", lambda e: auto_save())
    w.bind("<KeyRelease>", lambda e: auto_save())


# ===== подпись (как ты хотел)

footer = tk.Label(
    root,
    text="Разработал: Миндалев П.Ю.",
    fg="#222222",
    font=("Arial", 10, "italic"),
    anchor="e"
)
footer.pack(side="bottom", anchor="e", padx=10, pady=5)


root.mainloop()