import time
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from tkinter import messagebox
import threading
from process import process


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ứng dụng Sanji by AHT-ASIA")
        self.root.geometry("1000x800")  
        self.sheet = None
        # Button Import Excel
        import_button = tk.Button(root, text="Import Excel", height=2, width=20, command=self.load_excel)
        import_button.pack(pady=10)
        import_button.pack()

        # Tạo Combobox
        self.combo_box = ttk.Combobox(root, values=list([]), width=50, state="readonly")
        self.combo_box.bind("<<ComboboxSelected>>", self.on_combobox_select)
        self.combo_box.place(relx=0.5, rely=0.5, anchor="center")
        
        # Tạo Treeview ban đầu
        self.frame = ttk.LabelFrame(root, text="Dữ liệu từ Excel")
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree = None 

        self.urlExcel = None
        self.y_scrollbar = None
        self.x_scrollbar = None

        self.center_window(1200, 800)
                
    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                # Đọc tệp Excel
                xls = pd.ExcelFile(file_path)
                self.urlExcel = xls
                self.combo_box.pack(pady=10)
                self.df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
                sheet_names = pd.ExcelFile(file_path).sheet_names
                selected_option = sheet_names[0]
                self.sheet = selected_option
                self.combo_box['values'] = sheet_names
                self.combo_box.set(xls.sheet_names[0])
                self.update_treeview(self.df)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi đọc tệp Excel: {str(e)}")

    def update_treeview(self, df):
        if self.tree:
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.show_data(df)
        else:
            self.show_data(df)

    def on_combobox_select(self, event):
        selected_option = self.combo_box.get()
        self.df = pd.read_excel(self.urlExcel, sheet_name=selected_option)
        sheet_names = pd.ExcelFile(self.urlExcel).sheet_names
        self.combo_box['values'] = sheet_names
        self.update_treeview(self.df)
        self.sheet = selected_option

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")


    def show_data(self, df):
        if self.y_scrollbar:
            self.y_scrollbar.destroy()
        if self.x_scrollbar:
            self.x_scrollbar.destroy()
        if self.tree:
            self.tree.destroy()
        self.y_scrollbar = ttk.Scrollbar(self.frame, orient="vertical")
        self.x_scrollbar = ttk.Scrollbar(self.frame, orient="horizontal")
        style = ttk.Style()
        style.configure('Treeview.Heading', foreground='black', background='#66FFFF', font=('Arial',14),)
        self.tree = ttk.Treeview(self.frame, columns=list(df.columns), show="headings",
                                 xscrollcommand=self.x_scrollbar.set)
        self.tree.bind("<Double-Button-1>", self.on_treeview_select)
        for column in df.columns:
            self.tree.heading(column, text=column)

        for row in df.itertuples(index=False):
            formatted_row2 = tuple('Rỗng' if pd.isna(value) else value for value in row)
            row = formatted_row2
            if row[0] != 'Rỗng':
                formatted_row = (int(row[0]),) + row[1:]
                self.tree.insert("", "end", values=formatted_row)

        self.x_scrollbar.pack(side="bottom", fill="x")
        self.y_scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.y_scrollbar.config(command=self.tree.yview)
        self.x_scrollbar.config(command=self.tree.xview)

    def on_treeview_select(self, event):
        selected_item = self.tree.selection()[0]
        data = self.tree.item(selected_item, "values")
        thread = threading.Thread(target=process, args=(self.sheet, data))
        time.sleep(0.2)
        thread.start()

    def exit_tool(self):
        exit(0)

if __name__ == "__main__":
    # Lấy phiên bản của Pandas
    pandas_version = pd.__version__

    # In phiên bản Pandas
    print("Phiên bản Pandas:", pandas_version)
    root = tk.Tk()

    app = App(root)
    root.mainloop()
