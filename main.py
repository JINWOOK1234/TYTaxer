from tkinter import *
from tkinter import filedialog, ttk, messagebox, simpledialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import os
import calendar
from datetime import datetime
from openpyxl import load_workbook
from tkinter import Toplevel, Label, Text, Scrollbar, RIGHT, Y, BOTH, END
import csv
from CardPaymentList import CardPaymentList


class ExcelComparerApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 거래처 비교 & 자동 양식 작성기")
        self.root.geometry("1100x750")
        self.root.configure(bg="#f8f8f8")

        self.file1_path = StringVar()
        self.file2_path = StringVar()
        self.template_path = StringVar()
        self.selected_month = StringVar()
        self.selected_month.set(f"{datetime.today().month}월")
        self.month_var = self.selected_month
        self.save_option = IntVar(value=0)
        self.template_option = IntVar(value=0)

        self.df_result = None
        self.df2 = None

            # 카드 결제 손님 목록 모듈화
        self.card_payment_list = CardPaymentList(self.root)
        
        self.setup_ui()

    def setup_ui(self):
        from ExcelHandler import on_fill_template, on_compare_file
        Label(self.root, text="엑셀 거래처 비교 및 세금계산서 양식 자동 작성기", font=("맑은 고딕", 16, "bold"), bg="#f8f8f8").pack(pady=10)
        frame = Frame(self.root, bg="#f8f8f8")
        frame.pack(pady=10)

        self.drop_label1 = Frame(frame, bg="#5b9bd5", relief="solid", bd=1, width=400, height=250)
        self.drop_label1.pack_propagate(False)
        Label(self.drop_label1, text="① 매출현황 파일 드래그", bg="#5b9bd5", fg="white", font=("Arial", 12)).pack()
        self.preview1 = ttk.Treeview(self.drop_label1)
        self.preview1.pack(expand=True, fill=BOTH, padx=5, pady=5)
        self.drop_label1.pack(side=LEFT, padx=20)
        self.drop_label1.drop_target_register(DND_FILES)
        self.drop_label1.dnd_bind('<<Drop>>', self.on_drop_1)

        self.drop_label2 = Frame(frame, bg="#5b9bd5", relief="solid", bd=1, width=400, height=250)
        self.drop_label2.pack_propagate(False)
        Label(self.drop_label2, text="② 거래처 목록 파일 드래그", bg="#5b9bd5", fg="white", font=("Arial", 12)).pack()
        self.preview2 = ttk.Treeview(self.drop_label2)
        self.preview2.pack(expand=True, fill=BOTH, padx=5, pady=5)
        self.drop_label2.pack(side=RIGHT, padx=20)
        self.drop_label2.drop_target_register(DND_FILES)
        self.drop_label2.dnd_bind('<<Drop>>', self.on_drop_2)

        Label(self.root, textvariable=self.file1_path, bg="#f8f8f8", fg="gray").pack()
        Label(self.root, textvariable=self.file2_path, bg="#f8f8f8", fg="gray").pack()

        save_frame = Frame(self.root, bg="#f8f8f8")
        save_frame.pack()
        Label(save_frame, text="비교 결과 저장 여부:", bg="#f8f8f8").pack(side=LEFT, padx=5)
        Radiobutton(save_frame, text="저장 안 함", variable=self.save_option, value=0, bg="#f8f8f8").pack(side=LEFT)
        Radiobutton(save_frame, text="저장 함", variable=self.save_option, value=1, bg="#f8f8f8").pack(side=LEFT)

        Button(self.root, text="📊 거래처 비교하기", command=lambda:on_compare_file(self), bg="#4caf50", fg="white", font=("맑은 고딕", 12)).pack(pady=10)

        frame2 = Frame(self.root, bg="#f8f8f8")
        frame2.pack(pady=10)
         # 양식 옵션
        self.template_option_a = Radiobutton(frame2, text="a. 계산서등록양식(일반)_대량", variable=self.template_option, value=0, bg="#f8f8f8", command=self.toggle_template_button)
        self.template_option_a.pack(side=LEFT)

        self.template_option_b = Radiobutton(frame2, text="b. 새로운 파일 열기", variable=self.template_option, value=1, bg="#f8f8f8", command=self.toggle_template_button)
        self.template_option_b.pack(side=LEFT)

        # 양식 선택 버튼
        self.template_button = Button(frame2, text="📁 양식 선택", command=self.load_template, state=DISABLED)
        self.template_button.pack(side=LEFT, padx=5)

        month_options = [f"{i}월" for i in range(1, 13)]
        OptionMenu(frame2, self.month_var, *month_options).pack(side=LEFT, padx=10)
        Button(frame2, text="📝 양식에 입력하기", command=lambda:on_fill_template(self), bg="#2196f3", fg="white").pack(side=LEFT, padx=10)

        Button(self.root, text="초기화", command=self.reset_all, bg="#e91e63", fg="white").pack(pady=5)
        Button(self.root, text="❓ 도움말 보기", command=self.show_help, bg="#9c27b0", fg="white").pack(pady=5)

   
        self.set_default_template()
        
            # 카드 결제 손님 목록 보기 버튼
        view_button = Button(self.root, text="카드 결제 손님 목록 보기", command=self.view_card_payment_list)
        view_button.pack(pady=10)

    def toggle_template_button(self):
        # "b. 새로운 파일 열기"를 선택하면 양식 선택 버튼을 활성화
        if self.template_option.get() == 1:
            self.template_button.config(state=NORMAL)
        else:
            self.template_button.config(state=DISABLED)

    def view_card_payment_list(self):
        # 카드 결제 손님 목록 보기 모달 창
        modal_window = Toplevel(self.root)
        modal_window.title("카드 결제 손님 목록")
        modal_window.geometry("600x400")

        # 카드 결제 손님 목록을 Treeview로 표시
        tree = ttk.Treeview(modal_window, columns=("Card Name", "Discount Amount"), show="headings")
        tree.pack(expand=True, fill=BOTH)

        tree.heading("Card Name", text="거래처명")
        tree.heading("Discount Amount", text="차감 금액")

          # 카드 결제 손님 목록을 Treeview에 추가
        self.treeview = tree
        self.update_treeview()


        # 마우스 오른쪽 버튼 메뉴
        def on_right_click(event):
            item = tree.identify('item', event.x, event.y)
        
            if not item:
                return  # 항목을 선택하지 않으면 함수 종료
            
            # 선택된 항목의 데이터를 가져오고, 수정 또는 삭제 메뉴를 보여줄 수 있습니다.
            selected_values = tree.item(item)["values"]
            card_name = selected_values[0]
            discount_amount = selected_values[1]
    
            context_menu = Menu(modal_window, tearoff=0)
            context_menu.add_command(label="수정", command=lambda: self.modify_card_payment(item, card_name, discount_amount, tree))
            context_menu.add_command(label="삭제", command=lambda: self.delete_card_payment(item,tree))
            context_menu.post(event.x_root, event.y_root)

        tree.bind("<Button-3>", on_right_click)
        
        # 새로 추가하는 '...' 버튼 추가 (동적 추가)
        add_button = Button(modal_window, text="새로 추가 (+)", command=self.open_card_payment_modal)
        add_button.pack(pady=10)

   
    def modify_card_payment(self, item, card_name, discount_amount, tree):
        # 수정 로직: 선택된 항목의 값을 수정하는 작업
        selected_values = tree.item(item)["values"]
        card_name = selected_values[0]
        discount_amount = selected_values[1]

        """ 카드 결제 손님 수정 """
        modal_window = Toplevel(self.root)
        modal_window.title("카드 결제 손님 수정")
        modal_window.geometry("400x200")

        # 수정할 거래처명 입력
        Label(modal_window, text="거래처명:").pack(pady=5)
        card_name_entry = Entry(modal_window, width=30)
        card_name_entry.insert(0, card_name)
        card_name_entry.pack(pady=5)

        # 수정할 차감 금액 입력
        Label(modal_window, text="차감 금액:").pack(pady=5)
        discount_amount_entry = Entry(modal_window, width=30)
        discount_amount_entry.insert(0, str(discount_amount))
        discount_amount_entry.pack(pady=5)
        
         # 수정 버튼
        def save_changes():
            new_card_name = card_name_entry.get()
            try:
                new_discount_amount = float(discount_amount_entry.get())
                if new_card_name and new_discount_amount >= 0:
                    index = tree.index(item)
                    self.card_payment_list.update_entry(index, new_card_name, new_discount_amount)
                    tree.item(item, values=(new_card_name, new_discount_amount))
                    self.update_treeview()
                    #messagebox.showinfo("수정 완료", f"{new_card_name}의 정보가 수정되었습니다.")
                    modal_window.destroy()
                else:
                    messagebox.showerror("입력 오류", "거래처명과 차감 금액을 정확히 입력해주세요.")
            except ValueError:
                messagebox.showerror("입력 오류", "차감 금액은 숫자여야 합니다.")

        save_button = Button(modal_window, text="수정", command=save_changes)
        save_button.pack(pady=10)

    def delete_card_payment(self, item, tree):
        # 삭제 로직: 선택된 항목을 삭제하는 작업
        selected_values = tree.item(item)["values"]
        card_name = selected_values[0]
        discount_amount = selected_values[1]

        # 삭제 확인
        confirm = messagebox.askyesno("삭제 확인", f"정말로 {card_name}을 삭제하시겠습니까?")
        if confirm:
            index = tree.index(item)
            self.card_payment_list.delete_entry(index)
            tree.delete(item)  # Treeview에서 해당 항목을 삭제
            self.update_treeview()
            #messagebox.showinfo("삭제 완료", "카드 결제 손님 정보가 삭제되었습니다.")

    def update_treeview(self):
        """ Treeview 갱신 """
        for row in self.treeview.get_children():
            self.treeview.delete(row)

        for card_name, discount_amount in self.card_payment_list.get_entries():
            self.treeview.insert("", "end", values=(card_name, discount_amount))

    #def save_card_payment_list(self):
     #   self.card_payment_list.save_card_payment_list()
 
   
    def show_help(self):
        help_window = Toplevel(self.root)
        help_window.title("도움말 - 사용 설명서")
        help_window.geometry("700x600")
        help_window.configure(bg="white")

        Label(help_window, text="📘 엑셀 거래처 비교 & 세금계산서 양식 자동입력기 사용법", font=("맑은 고딕", 14, "bold"), bg="white", fg="#333").pack(pady=10)

        text = Text(help_window, wrap="word", font=("맑은 고딕", 11), bg="white", fg="#222")
        text.pack(expand=True, fill=BOTH, padx=15, pady=10)

        scrollbar = Scrollbar(help_window, command=text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        text.config(yscrollcommand=scrollbar.set)

        help_content = """
    📁 1. 파일 드래그
    - 왼쪽: 매출현황 파일 (B열 아래에 매출처 목록이 위치)
    - 오른쪽: 거래처 목록 파일 (B열=상호, R열=별명)

    📊 2. 거래처 비교
    - 상호가 일치하면 '구분' = 1
    - 없으면 별명 비교, 둘 다 없으면 '구분' = 0
    - '일치 인덱스' = 거래처 목록 파일의 행번호 (양식 입력에 사용됨)

    💾 3. 결과 저장 옵션
    - 저장 안 함 (기본)
    - 저장 함 → 결과를 엑셀로 저장 가능

    📑 4. 양식 선택 방법
    - a. 기본 양식 사용 ('계산서등록양식(일반)_대량.xls') – 자동 로딩
    - b. 직접 선택 – 수동으로 양식 파일 선택 가능

    📆 5. 월 선택 & 양식에 자동입력
    - 선택한 월의 마지막 날짜를 작성일자로 사용
    - 일치 인덱스를 통해 거래처 정보 자동 입력
    - 매출금액 = (E열 - G열) 계산 후 입력됨

    💾 저장 시 기본 파일명
    → '[선택한월]_거래처등록양식(일반)_대량.xls' 또는 .xlsx
    예: 3월_거래처등록양식(일반)_대량.xls

    🧼 6. 초기화 버튼
    → 파일 경로, 미리보기, 내부 상태 모두 초기화

    ⚙️ 7. 필수 설치 모듈
    - pandas
    - openpyxl
    - tkinterdnd2
    - pywin32

    📎 참고
    - Excel이 설치된 환경에서 실행되어야 `.xls` 파일 저장 가능
    - 기본 양식 파일은 본 프로그램과 동일한 폴더에 위치해야 합니다
    """
        text.insert(END, help_content)
        text.config(state="disabled")


    def on_drop_1(self, event):
        path = event.data.strip("{}")
        self.file1_path.set(path)
        self.show_preview(path, self.preview1)

    def on_drop_2(self, event):
        path = event.data.strip("{}")
        self.file2_path.set(path)
        self.show_preview(path, self.preview2)
   
    def show_preview(self, path, tree):
        try:
            df = pd.read_excel(path).head(5)
            tree.delete(*tree.get_children())
            tree["columns"] = list(df.columns)
            tree["show"] = "headings"
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)
            for _, row in df.iterrows():
                tree.insert("", "end", values=list(row))
        except Exception as e:
            tree.delete(*tree.get_children())
            tree.insert("", "end", values=["미리보기 실패", str(e)])

    def reset_all(self):
        self.file1_path.set("")
        self.file2_path.set("")
        self.template_path.set("")
        self.df_result = None
        self.df2 = None
        for tree in [self.preview1, self.preview2]:
            tree.delete(*tree.get_children())
        messagebox.showinfo("초기화", "모든 정보가 초기화되었습니다.")
        
    def open_card_payment_modal(self):
        # 모달창을 생성
        modal_window = Toplevel(self.root)
        modal_window.title("카드 결제 손님 추가")
        modal_window.geometry("400x200")

        # 거래처명 입력
        Label(modal_window, text="거래처명:").pack(pady=5)
        card_name_entry = Entry(modal_window, width=30)
        card_name_entry.pack(pady=5)

        # 차감 금액 입력
        Label(modal_window, text="차감 금액:").pack(pady=5)
        discount_amount_entry = Entry(modal_window, width=30)
        discount_amount_entry.pack(pady=5)

        # 추가 버튼 (입력값을 목록에 추가)
        def add_entry():
            card_name = card_name_entry.get()
            try:
                discount_amount = float(discount_amount_entry.get())
                if card_name and discount_amount >= 0:
                    self.card_payment_list.add_card_payment_entry(card_name, discount_amount)
                    self.update_treeview()
                 #   messagebox.showinfo("추가 완료", f"{card_name}의 카드 결제 정보가 추가되었습니다.")
                    modal_window.destroy()  # 모달창 닫기
                else:
                    messagebox.showerror("입력 오류", "거래처명과 차감 금액을 정확히 입력해주세요.")
            except ValueError:
                messagebox.showerror("입력 오류", "차감 금액은 숫자여야 합니다.")

        add_button = Button(modal_window, text="추가", command=add_entry)
        add_button.pack(pady=10)        
    

    def set_default_template(self):
        if self.template_option.get() == 0:
            # 수동으로 변환된 .xlsx 파일 경로 설정
            script_dir = os.path.dirname(os.path.abspath(__file__))  # 현재 스크립트가 위치한 폴더
            default_path = os.path.join(script_dir, "계산서등록양식(일반)_대량.xlsx")  # .xlsx 형식으로 설정

            if os.path.exists(default_path):
                self.template_path.set(default_path)
                messagebox.showinfo("기본 양식 설정", f"기본 양식이 설정되었습니다:\n{default_path}")
            else:
                messagebox.showwarning("파일 없음", "기본 양식 파일이 존재하지 않습니다.")

    def load_template(self):
        if self.template_option.get() == 1:
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if path:
                self.template_path.set(path)
                messagebox.showinfo("양식 선택 완료", f"선택된 양식: {os.path.basename(path)}")


if __name__ == '__main__':
    root = TkinterDnD.Tk()
    app = ExcelComparerApp(root)
    root.mainloop()

def on_close(self):
    self.save_card_payment_list()  # 프로그램 종료 전에 목록 저장
    self.root.destroy()  # 종료 창을 닫음