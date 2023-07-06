import pandas as pd
from tkinter import Tk, Button, filedialog, Label, Entry, Toplevel
from tkinter.messagebox import showinfo
import analyze_file_hash
import analyze_ip_addresses
import analyze_domain

def browse_input_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    input_file_entry.delete(0, 'end')
    input_file_entry.insert('end', file_path)

def browse_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])
    output_file_entry.delete(0, 'end')
    output_file_entry.insert('end', file_path)

def start_analysis():
    input_file = input_file_entry.get()
    output_file = output_file_entry.get()
    df = pd.read_excel(input_file)
    first_column_name = df.columns[0]
    if first_column_name == 'ip':
        if input_file and output_file:
            analyze_ip_addresses.analyze_ip_addresses(input_file, output_file)
        else:
            showinfo("Lỗi", "Vui lòng chọn tệp đầu vào và tệp đầu ra!")
    elif first_column_name == 'hash':
        if input_file and output_file:
            analyze_file_hash.analyze_file_hash(input_file, output_file)
        else:
            showinfo("Lỗi", "Vui lòng chọn tệp đầu vào và tệp đầu ra!")
    elif first_column_name == 'domain':
        if input_file and output_file:
            analyze_domain.analyze_domain(input_file, output_file)
        else:
            showinfo("Lỗi", "Vui lòng chọn tệp đầu vào và tệp đầu ra!")
    else:
        showinfo("Lỗi", "Vui lòng kiểm tra tên cột trong file!")

# Tạo giao diện Windows Forms
root = Tk()
root.iconbitmap("icon.ico")
root.title("Đỗ Văn Quang")


input_file_label = Label(root, text="Tệp đầu vào:")
input_file_label.grid(row=0, column=0, padx=5, pady=5)

input_file_entry = Entry(root)
input_file_entry.grid(row=0, column=1, padx=5, pady=5)

input_file_button = Button(root, text="Chọn tệp", command=browse_input_file)
input_file_button.grid(row=0, column=2, padx=5, pady=5)

output_file_label = Label(root, text="Tệp đầu ra:")
output_file_label.grid(row=1, column=0, padx=5, pady=5)

output_file_entry = Entry(root)
output_file_entry.grid(row=1, column=1, padx=5, pady=5)

output_file_button = Button(root, text="Chọn tệp", command=browse_output_file)
output_file_button.grid(row=1, column=2, padx=5, pady=5)

start_button = Button(root, text="Bắt đầu phân tích", command=start_analysis)
start_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()