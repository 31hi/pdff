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
root.geometry("500x250")
root.minsize(500, 250)

frameOfInputAndOutPut = ttk.Frame(root)
frameOfInput = ttk.Frame(frameOfInputAndOutPut)
frameOfOutput = ttk.Frame(frameOfInputAndOutPut)
btn1 = ttk.Button(root, text="PDFファイルの保存", width=20, command=generatePDF)
label1 = ttk.Label(frameOfInput, text="ファイルのパスを指定")
label2 = ttk.Label(frameOfOutput, text="抽出するページ数")
IDirButton = ttk.Button(frameOfInput, text="参照", width=5, command=dirdialog_clicked)
txt1 = ttk.Entry(frameOfInput, width=20)
txt2 = ttk.Entry(frameOfOutput, width=20)

# 画面に表示
label1.pack(side=tk.LEFT)
txt1.pack(side=tk.LEFT, ipadx=5, ipady=5)
IDirButton.pack(side=tk.LEFT, ipadx=5, ipady=5)
frameOfInput.pack(side=tk.TOP, anchor=tk.NW)
label2.pack(side=tk.LEFT)
txt2.pack(side=tk.LEFT, ipadx=5, ipady=5)
frameOfOutput.pack(side=tk.TOP, anchor=tk.NW)
frameOfInputAndOutPut.pack(side=tk.TOP, anchor=tk.CENTER, pady=30)
btn1.pack(side=tk.TOP, anchor=tk.CENTER, padx="20", pady="20", ipadx=5, ipady=5)

root.mainloop()