import tkinter as tk
from tkinter import filedialog
import pandas as pd

def srt_to_dataframe(file_path):
    encodings = ['utf-8', 'latin-1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue

    times = []
    subtitles = []
    current_sub = ""
    current_time = ""

    for line in lines:
        line = line.strip()
        if '-->' in line:
            current_time = line
        elif line.isdigit() or line == '':
            if current_time and current_sub:
                times.append(current_time)
                subtitles.append(current_sub.strip())
                current_time = ""
                current_sub = ""
        else:
            current_sub += " " + line

    return pd.DataFrame({'Time': times, 'Subtitle': subtitles})

def save_to_excel(dataframe, output_path):
    dataframe.to_excel(output_path, index=False)

def convert_to_xlsx():
    file_path = filedialog.askopenfilename(filetypes=[("SRT files", "*.srt")])
    if file_path:
        output_file = file_path.replace('.srt', '.xlsx')
        df = srt_to_dataframe(file_path)
        save_to_excel(df, output_file)
        status_label.config(text=f'Arquivo criado: {output_file}')

root = tk.Tk()
root.title("Legendas XLSX")

select_button = tk.Button(root, text='Selecionar Arquivo SRT', command=convert_to_xlsx)
select_button.pack(pady=30, padx=100)

status_label = tk.Label(root, text='')
status_label.pack()

root.mainloop()