import os
import csv
from tkinter import *
from tkinter import simpledialog

class CardPaymentList:
    def __init__(self, root):
        self.root = root
        self.card_payment_entries = {}  # 카드 결제 손님 목록을 딕셔너리로 변경
        self.load_card_payment_list()  # 앱 시작 시 카드 결제 손님 목록 불러오기

    def add_card_payment_entry(self, card_name, discount_amount):
        # 카드 결제 손님 추가 (딕셔너리 형태로 추가)
        self.card_payment_entries[card_name] = discount_amount
        self.save_card_payment_list()

    def save_card_payment_list(self):
        # 카드 결제 손님 목록을 파일로 저장 (CSV 파일)
        with open("card_payment_list.csv", mode="w", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["거래처명", "차감 금액"])  # 헤더 작성
            for card_name, discount_amount in self.card_payment_entries.items():
                writer.writerow([card_name, discount_amount])

    def load_card_payment_list(self):
        # 카드 결제 손님 목록을 파일에서 불러오기 (CSV 파일)
        if os.path.exists("card_payment_list.csv"):
            with open("card_payment_list.csv", mode="r") as file:
                reader = csv.reader(file)
                try:
                    next(reader)  # 헤더를 건너뛰기
                except StopIteration:
                    return  # 파일이 비어있으면 종료
                for row in reader:
                    if len(row) == 2:
                        card_name, discount_amount = row
                        self.card_payment_entries[card_name] = float(discount_amount)

    def get_entries(self):
        return self.card_payment_entries  # 딕셔너리 형태로 반환
    
    def to_dict(self):
        """거래처명을 키, 차감 금액을 값으로 하는 dict 반환"""
        return self.card_payment_entries.copy()


    def delete_entry(self, card_name):
        # 카드 결제 손님 삭제 (딕셔너리에서 삭제)
        if card_name in self.card_payment_entries:
            del self.card_payment_entries[card_name]
            self.save_card_payment_list()

    def update_entry(self, card_name, new_discount_amount):
        # 카드 결제 손님 수정 (딕셔너리에서 수정)
        if card_name in self.card_payment_entries:
            self.card_payment_entries[card_name] = new_discount_amount
            self.save_card_payment_list()
