from openpyxl import load_workbook

from collections import defaultdict

book_f = open('books.txt', 'r', encoding='utf8')
books = book_f.read().splitlines()

def get_data(wb):
    names = [[] for i in range(len(books))]
    wb = wb['설문지 응답 시트1']
    wb = [[str(cell.value) for cell in row] for row in wb]
    name_idx = dict()
    for i in range(1, len(wb)):
        name_idx[wb[i][1]] = i
    for name in name_idx:
        idx = name_idx[name]
        all_books = ', '.join(wb[idx][2:11])
        for i in range(len(books)):
            if books[i] in all_books:
                names[i].append(name)
    return names

BUY, SELL = 0, 1
buy = load_workbook('구매 신청서(응답).xlsx')
buy = get_data(buy)
sell = load_workbook('판매 신청서(응답).xlsx')
sell = get_data(sell)
people = defaultdict(list)
for i in range(len(books)):
    print(books[i])
    print(f'구매: {buy[i]}')
    print(f'판매: {sell[i]}')
for i in range(len(books)):
    amount = min(len(buy[i]), len(sell[i]))
    for j in range(len(buy[i])):
        people[buy[i][j]].append([BUY, j<amount, i])
    for j in range(len(sell[i])):
        people[sell[i][j]].append([SELL, j<amount, i])
for p in people:
    print(p)
    for book in people[p]:
        if book[0] == BUY:
            if book[1]:
                print(f'구매: {books[book[2]]}')
            else:
                print(f'구매 실패: {books[book[2]]}')
        else:
            if book[1]:
                print(f'판매: {books[book[2]]}')
            else:
                print(f'판매 실패: {books[book[2]]}')
