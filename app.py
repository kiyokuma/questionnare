# coding: utf-8
import tkinter as tk
from questionnaire import startExcelCreate

class mainFrame():
    def excelCreate(self):
        startExcelCreate()

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"アンケート集計")
        self.root.geometry("800x500")
        self.root.grid_columnconfigure(0, weight=1)

        self.frame0 = tk.Frame(self.root,relief=tk.GROOVE,padx=20)
        self.frame1 = tk.Frame(self.root, padx=20)
        self.frame4 = tk.Frame(self.root,relief=tk.GROOVE,bg='white', bd=2, padx=20)

        
        self.Title = tk.Label(self.frame0, text="アンケートデータ集計", font=("meiryo", "10" ))
        self.label_frame1_text_1 = tk.Label(self.frame1, text='ファイルを取り込んでください。', font=("meiryo", "10" ))

        self.button_all = tk.Button(self.frame1,
                                text="ファイル取込み",
                                command=self.excelCreate,
                                font=("meiryo", "10" ))

        self.frame0.grid(row=0, column=0,sticky=tk.EW,columnspan = 1)
        self.frame1.grid(row=1, column=0,sticky=tk.EW + tk.NS,columnspan = 2)
        self.Title.pack(fill=tk.X, pady=5)
        self.button_all.grid(row=1, column=1,pady=10,padx=10)
        self.frame4.grid(row=4, column=0,sticky=tk.EW + tk.NS,columnspan = 2)
        
        self.label_frame4_text_1 = tk.Label(self.frame4, text='使い方',bg='white', font=("meiryo", "10" ) )
        self.label_frame4_text_2 = tk.Label(self.frame4, text='1.T-net「カスタマーフィードバックデータ」から【契約】【損害】それぞれの回答一覧をExcelデータでダウンロードします。',bg='white', font=("meiryo", "8" ))
        self.label_frame4_text_3 = tk.Label(self.frame4, text='2.それぞれのファイルごとにファイル取込みボタンから集計ファイルを作成します。',bg='white', font=("meiryo", "8" ))
        self.label_frame4_text_4 = tk.Label(self.frame4, text='3.ファイルが作成されたら、【契約】【損害】がわかるようにファイルの名前の変更をお願いします。',bg='white',font=("meiryo", "8" ) )

        self.label_frame1_text_1.grid(row=0, column=0,pady=10,padx=10)
        self.label_frame4_text_1.pack(padx=5,pady=5,anchor=tk.W)
        self.label_frame4_text_2.pack(padx=5,anchor=tk.W)
        self.label_frame4_text_3.pack(padx=5,anchor=tk.W)
        self.label_frame4_text_4.pack(padx=5,anchor=tk.W)

        self.root.mainloop()

def open_app():
    mainFrame()
        
if __name__ == '__main__':
    open_app()