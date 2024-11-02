import tkinter as tk
from tkinter import messagebox, Label
from tkinter.filedialog import askopenfilename
from Scripts import comment_script, comment_script_xml


def About():
   messagebox.showinfo("О программе","Это простая программа для работы с документами ворд, в которых есть комментарии. Если вы хотите подсказать идеей или нашли какие-либо баги, пожалуйста, напишите мне.\nserpantyne03@gmail.com")

def open_file():
   file = askopenfilename(filetypes =[('Word Files', '*.docx')])
   if file is not None:
      comment_script(file)


def open_file_xml():
   file = askopenfilename(filetypes =[('Word Files', '*.docx')])
   if file is not None:
      comment_script_xml(file)

class App(tk.Tk):
   def __init__(self):
      super().__init__()
      menu = tk.Menu(self)
      file_menu = tk.Menu(menu, tearoff=0)

      file_menu.add_command(label="Выбрать файл", command=lambda: open_file())
      file_menu.add_command(label="Открыть с помощью XML", command=lambda: open_file_xml())
      file_menu.add_separator()
      file_menu.add_command(label="Сохранить как...")
      file_menu.add_command(label="Сохранить как")


      menu.add_cascade(label="Файл", menu=file_menu)
      menu.add_cascade(label="О программе", command=About)

      menu.add_cascade(label="Выход", command=self.destroy)
      self.config(menu=menu)

if __name__ == "__main__":
    app = App()
    app.iconbitmap(default="icon.ico") 
    app.label = Label(app, text="Привет!\nДля начала работы нажми на кнопку \n в контексном меню 'Файл'\n после 'Выбрать файл'", fg="green", font=("Arial", 12, "bold"))
    app.label.pack()
    app.label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
    app.title("Комметарии Word")
    app.geometry("350x350+700+300")
    app.mainloop()
