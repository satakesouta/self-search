#自動検索ツール（mac用）
#検索1.xlsxに乗った単語を任意のサイトで自動検索し表示してくれる（最大30語）
#検索1.xlsxの単語は任意で
#ブラウザアプリのパス（ブラウザアプリのUNIX実行ファイルのパス）及び検索したいサイトのURLは任意で指定

import subprocess
from time import sleep
import openpyxl
import tkinter


class NotebookSample(tkinter.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.launch()
        self.design()
    
    def launch(self):
        # ブラウザ起動
        N = "調べたいサイトのURLを記入"
        subprocess.Popen(['ブラウザアプリのUNIX実行ファイルのパス',N])
        # エクセル起動
        subprocess.Popen(['open','検索1.xlsxのパスを記入'])
    
    def launchEXCEL(self):
        # エクセル起動、ウィンドウデザイン変更
        subprocess.Popen(['open','検索1.xlsxのパスを記入'])
        label1.pack_forget()
        label2.pack_forget()
        Button1.pack_forget()
        root.configure(bg = 'white')
        self.design()

    def FinishWindow(self):
        # 検索終了のお知らせ&再度検索選択ウィンドウ表示
        root.lift()
        label.pack_forget()
        Button.pack_forget()
        global label1,label2,Button1
        label1 = tkinter.Label(root,text = "自動検索完了！",font = ("",30,"bold"),fg = "lime",bg = 'white')
        label1.pack()
        label2 = tkinter.Label(root,text = "まだ検索する？",font = ("",20),fg = "black",bg = 'white')
        label2.pack()
        Button1 = tkinter.Button(text='Yes!',command = self.launchEXCEL,font=("",20),bg = "aqua",fg = "black")
        Button1.pack()

    def Search(self):
        # ブラウザで繰り返し自動検索する関数
        # 調べたいサイトのURLをs1,s2,s3に分ける（Google仕様）
        # s1とs2、s2とs3の間に検索ワードを入れる
        s1 = "https://www.google.com/search?q="
        s2 = "&tbm=isch&ved=2ahUKEwjDjLGQuOnuAhWYAKYKHXCaB6wQ2-cCegQIABAA&oq="
        s3 = "&gs_lcp=CgNpbWcQAzIECCMQJ1BUWI8PYIESaAFwAHgAgAHRAogB6giSAQcwLjQuMS4xmAEAoAEBqgELZ3dzLXdpei1pbWfAAQE&sclient=img&ei=1h8pYMPJFJiBmAXwtJ7gCg&rlz=1C1EJFC_enJP888JP888"
        for i in range(int(sheet['G8'].value)):
            V[i] = s1 + S[i] + s2 + S[i] + s3
            subprocess.Popen(['ブラウザアプリのUNIX実行ファイルのパス',V[i+1]])
            sleep(5)

        # V[]が検索URL
        self.FinishWindow()


    def Urldeta(self):
        # URL
        # 調べたいサイトのURLをs1,s2,s3に分ける
        # s1とs2、s2とs3の間に検索ワードを入れる
        # 下ではgoogleで検索
        s1 = "https://www.google.com/search?q="
        s2 = "&tbm=isch&ved=2ahUKEwjDjLGQuOnuAhWYAKYKHXCaB6wQ2-cCegQIABAA&oq="
        s3 = "&gs_lcp=CgNpbWcQAzIECCMQJ1BUWI8PYIESaAFwAHgAgAHRAogB6giSAQcwLjQuMS4xmAEAoAEBqgELZ3dzLXdpei1pbWfAAQE&sclient=img&ei=1h8pYMPJFJiBmAXwtJ7gCg&rlz=1C1EJFC_enJP888JP888"
        # エクセルからデータ読み出し
        wb = openpyxl.load_workbook('検索1.xlsxのパスを記入')
        global sheet
        sheet = wb['Sheet1']
        def get_value_list(t_2d):
            return([[cell.value for cell in row] for row in t_2d])
        cell = get_value_list(sheet['A2:C11'])

        # エクセルデータを辞書型変数Sに書き込み
        global S
        S = []
        for i in range(10):
            for m in range(3):
                S.append(str(cell[i][m]))

        # エクセルデータをURL化
        global V
    
        V ={}
        V[0] = s1 + S[0] + s2 + S[0] + s3
        self.Search(sheet)


    def design(self):
        # ウィンドウのデザイン
        global label,Button
        label = tkinter.Label(root,text = "検索始める？",font = ("",25),fg = "red",bg = 'white')
        label.pack()
        Button = tkinter.Button(text='Yes!',command = self.Urldeta,font=("",25),bg = "aqua",fg = "black")
        Button.pack(pady=15)




if __name__ == '__main__':
    # ウィンドウの表示
    root = tkinter.Tk()
    root.title(u"自動検索")
    root.geometry("200x120")
    root.configure(bg = 'white')
    # 最大化ボタン非表示
    root.resizable(0,0)
    NotebookSample(master = root)
    #Button.place(x=110, y=80,width=80,height=50)
    root.mainloop()


