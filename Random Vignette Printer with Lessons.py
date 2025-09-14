import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from escpos.printer import Usb
import textwrap
import random
import datetime
import os

# Printer config
VENDOR_ID = 0x04b8
PRODUCT_ID = 0x0e28

def wrap_text(text, width=48):
    return "\n".join(textwrap.wrap(text, width=width))

class VignetteOptionsDialog(tk.Toplevel):
    def __init__(self, parent, counts):
        super().__init__(parent)
        self.title("Vignette Options")
        self.result = None

        self.length_var = tk.IntVar()
        self.include_reflection = tk.BooleanVar(value=True)
        self.auto_print = tk.BooleanVar(value=True)

        frame = tk.Frame(self)
        frame.pack(padx=10, pady=10)

        tk.Label(frame, text="Select vignette length:").pack(anchor='w')

        self.enabled_options = 0
        for value, label in [(1, "Short [fewer than 200 words]"), (2, "Medium [between 200 & 300 words]"), (3, "Long [more than 300 words]")]:
            count = counts.get(value, 0)
            label_text = f"{label} ({count} remaining)"
            state = 'normal' if count > 0 else 'disabled'
            btn = tk.Radiobutton(frame, text=label_text, variable=self.length_var, value=value, state=state)
            btn.pack(anchor='w', padx=10)
            if count > 0 and self.enabled_options == 0:
                self.length_var.set(value)
            if count > 0:
                self.enabled_options += 1

        tk.Checkbutton(frame, text="Include a reflection question", variable=self.include_reflection).pack(anchor='w', pady=(10, 5))
        tk.Checkbutton(frame, text="Skip screen display and print directly", variable=self.auto_print).pack(anchor='w', pady=(0, 10))

        button_frame = tk.Frame(frame)
        button_frame.pack()

        self.ok_button = tk.Button(button_frame, text="OK", command=self.on_ok, state='normal' if self.enabled_options > 0 else 'disabled')
        self.ok_button.pack(side='left', padx=10)

        tk.Button(button_frame, text="Exit", command=self.on_exit).pack(side='left', padx=10)

        self.grab_set()
        self.wait_window()

    def on_ok(self):
        self.result = {
            'length': self.length_var.get(),
            'include_reflection': self.include_reflection.get(),
            'auto_print': self.auto_print.get()
        }
        self.destroy()

    def on_exit(self):
        self.result = None
        self.destroy()

# Initialize app
root = tk.Tk()
root.withdraw()

messagebox.showinfo("Select File", "Please select the spreadsheet where the vignette data is stored.")
file_path = filedialog.askopenfilename(
    title="Select the Excel file containing vignettes",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)
if not file_path:
    messagebox.showerror("Error", "No file selected. Exiting.")
    root.destroy()
    exit()

try:
    df = pd.read_excel(file_path)
except Exception as e:
    messagebox.showerror("Error", f"Failed to load the file:\n{e}")
    root.destroy()
    exit()

used_ids = set()
log_path = "vignette_print_log.csv"
log_columns = ["ID", "Timestamp", "Reflection_Included", "Reflection_Text"]
if not os.path.exists(log_path):
    pd.DataFrame(columns=log_columns).to_csv(log_path, index=False)

# Main loop
while True:
    available_df = df[~df['ID'].isin(used_ids)]
    if available_df.empty:
        messagebox.showinfo("Pool Reset", "All vignettes have been used in this session. The pool has been reset.")
        used_ids = set()
        available_df = df

    available_counts = available_df.groupby('Length').size().to_dict()
    root.update()
    dialog = VignetteOptionsDialog(root, available_counts)
    if not dialog.result:
        break

    length = dialog.result['length']
    include_reflection = dialog.result['include_reflection']
    auto_print = dialog.result['auto_print']

    filtered_df = df[(df['Length'] == length) & (~df['ID'].isin(used_ids))]
    if filtered_df.empty:
        messagebox.showinfo("No Results", "No more vignettes available for the selected length.")
        continue

    vignette = filtered_df.sample(n=1).iloc[0]
    used_ids.add(vignette['ID'])

    lines = []
    if pd.notna(vignette['Warning']) and str(vignette['Warning']).strip():
        lines.append(f"Content Warning: {vignette['Warning']}\n")
    lines.append(vignette['Content'])

    citation_line = f"Excerpted from page/s: {vignette['Page_No']} in {vignette['Author_last']}, {vignette['Author_first']}. {vignette['Publication_date']}. {vignette['Title']}. {vignette['Publisher_Journal_Website']}."
    lines.append("")
    lines.append(citation_line)

    reflection_line = ""
    reflection_selected = False
    if include_reflection:
        q_options = [vignette[col] for col in ['Q1', 'Q2', 'Q3', 'Q4'] if pd.notna(vignette.get(col)) and str(vignette[col]).strip()]
        if q_options:
            selected_question = random.choice(q_options)
            reflection_line = f"Reflection: {selected_question}"
            lines.append("")
            lines.append(reflection_line)
            reflection_selected = True

    lesson_title = vignette.get('Lesson_title')
    lesson_link = vignette.get('Lesson_link')
    lesson_text = ""
    if pd.notna(lesson_title) and pd.notna(lesson_link):
        lesson_text = f"Curious about how to explore \"{lesson_title}\" in your ethnographic writing? Check out this short lesson:"
        lines.append("")
        lines.append(lesson_text)
        lines.append(lesson_link)

    if not auto_print:
        messagebox.showinfo("Random Vignette", "\n".join(lines))

    if auto_print or messagebox.askyesno("Print Vignette", "Would you like to print this vignette?"):
        try:
            printer = Usb(VENDOR_ID, PRODUCT_ID)
            printer.set(align='left', font='a', width=1, height=1)

            if pd.notna(vignette['Warning']) and str(vignette['Warning']).strip():
                printer.text(wrap_text(f"Content Warning: {vignette['Warning']}") + "\n\n")
            printer.text(wrap_text(vignette['Content']) + "\n")
            printer.text("-" * 48 + "\n")
            printer.text(wrap_text(citation_line) + "\n")
            printer.text("-" * 48 + "\n")
            if reflection_line:
                printer.text(wrap_text(reflection_line) + "\n")
                printer.text("-" * 48 + "\n")
            if lesson_text:
                printer.text(wrap_text(lesson_text) + "\n")
                printer.qr(lesson_link, size=6)
            printer.cut()

            log_entry = {
                "ID": vignette['ID'],
                "Timestamp": datetime.datetime.now().isoformat(),
                "Reflection_Included": reflection_selected,
                "Reflection_Text": selected_question if reflection_selected else ""
            }
            pd.DataFrame([log_entry]).to_csv(log_path, mode='a', header=False, index=False)

        except Exception as e:
            messagebox.showerror("Printing Error", f"Could not print to thermal printer:\n{e}")

root.destroy()
