import os
import ctypes
from pypdf import PdfReader, PdfWriter
import tkinter as tk
import subprocess

from tkinter import ttk
from tkinter import filedialog

#PDFを生成するボタンが押されたとき
def generatePDF() :

    # txt2から取得したデータをページ番号に振り分け配列をつくる
    pages = txt2.get()
    pageArr = []
    if ',' in pages:
        splitByComma = pages.split(',')
        for page in splitByComma:
            page = page.strip()
            if '-' in page:
                pageSpace = page.split('-')
                pageArr.extend(range(int(pageSpace[0]), int(pageSpace[1]) + 1))
            else:
                pageArr.append(int(page))
    else:
        if '-' in pages:
            pageSpace = pages.split('-')
            pageArr.extend(range(int(pageSpace[0]), int(pageSpace[1]) + 1))
        else:
            pageArr = int(page)

    # 分割するファイルを指定
    targetFile = txt1.get()
    pageArr = list(map(lambda x: x-1, pageArr))

    # 保存先のPDFファイル
    reader = PdfReader(targetFile)
    filename = filedialog.asksaveasfilename(
        title = "名前を付けて保存",
        filetypes = [("PDF", ".pdf")], # ファイルフィルタ
        initialdir = "",
        defaultextension = ".pdf"
    )
    # 書き込み用のオブジェクトを作成
    writer = PdfWriter()

    # PDFを抜き出す
    reader = PdfReader(targetFile)
    for i in pageArr:
        writer.add_page(reader.pages[i])

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

root.title(u"PDFF")
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