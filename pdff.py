import os
import ctypes
from pypdf import PdfReader, PdfWriter
import tkinter as tk
import subprocess

from tkinter import ttk
from tkinter import filedialog

#PDFを生成するボタンが押されたとき
def generatePDF() :
    targetFile = txt1.get()
    targetPage = int(txt2.get()) - 1
    reader = PdfReader(targetFile)
    filename = filedialog.asksaveasfilename(
        title = "名前を付けて保存",
        filetypes = [("PDF", ".pdf")], # ファイルフィルタ
        initialdir = "",
        defaultextension = ".pdf"
    )
    # 書き込み用のオブジェクトを作成
    writer = PdfWriter()

    # targetPath番目の要素(1ページ目のPDF)を抜き出す
    pdf = reader.pages[targetPage]

    # 書き込み用オブジェクトに追加
    writer.add_page(pdf)

    # ファイルに書き出し
    with open(filename, "wb") as fp:
        writer.write(fp)

    #エクスプローラーを開く
    subprocess.run(['start', '', filename], shell=True)

#参照ボタンがおされたとき
def dirdialog_clicked() :
    idir = ''
    filetype = [("PDF","*.pdf"), ("すべて","*")]
    file_path = tk.filedialog.askopenfilename(filetypes = filetype, initialdir = idir)
    txt1.insert(tk.END, file_path)

#画面に表示するコンテンツを作成
ctypes.windll.shcore.SetProcessDpiAwareness(1)
root = tk.Tk()

root.title(u"PDF")
root.geometry("300x200")

btn1 = ttk.Button(root, text="Generate PDF", width=15, command=generatePDF)
label1 = ttk.Label(text="File Path")
label2 = ttk.Label(text="Page Number")
IDirButton = ttk.Button(text="参照", command=dirdialog_clicked)

txt1 = ttk.Entry(width=20)
txt2 = ttk.Entry(width=20)

# 画面に表示
label1.pack()
txt1.pack()
IDirButton.pack()
label2.pack()
txt2.pack()
btn1.pack()

root.mainloop()