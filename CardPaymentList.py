import os
import csv
from tkinter import *
from tkinter import simpledialog

class CardPaymentList:
    def __init__(self, root):
        self.root = root
        self.card_payment_entries = []  # 카드 결제 손님 목록
        self.load_card_payment_list()  # 앱 시작 시 카드 결제 손님 목록 불러오기

    def add_card_payment_entry(self, card_name, discount_amount):
        # 카드 결제 손님 추가
        self.card_payment_entries.append((card_name, discount_amount))
        self.save_card_payment_list()

    def save_card_payment_list(self):
        # 카드 결제 손님 목록을 파일로 저장 (CSV 파일)
        with open("card_payment_list.csv", mode="w", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["거래처명", "차감 금액"])  # 헤더 작성
            for card_name, discount_amount in self.card_payment_entries:
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
                        self.card_payment_entries.append((card_name, float(discount_amount)))

    def get_entries(self):
        return self.card_payment_entries

    def delete_entry(self, index):
        # 카드 결제 손님 삭제
        del self.card_payment_entries[index]
        self.save_card_payment_list()

    def update_entry(self, index, new_card_name, new_discount_amount):
        # 카드 결제 손님 수정
        self.card_payment_entries[index] = (new_card_name, new_discount_amount)
        self.save_card_payment_list()

