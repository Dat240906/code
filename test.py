import os

os.system('pip install openpyxl')
try: 
    import pandas as pd
    import math
except:
    os.system('pip install openpyxl && pip install pandas')



def get_data( name_col, point_col):
    excel_file_path = r'test.xlsx'

    data = pd.read_excel(excel_file_path, usecols=[name_col, point_col], engine='openpyxl')

    ho_ten_list = []
    diem_list = []

    for index, row in data.iterrows():
        ho_ten = row[name_col]
        diem = row[point_col]
        
        if math.isnan(diem):
            continue
        
        ho_ten_list.append(ho_ten)
        diem_list.append(diem)

    # Tạo từ điển từ danh sách họ tên và điểm
    student_scores = {ho_ten: diem for ho_ten, diem in zip(ho_ten_list, diem_list)}

    # In danh sách họ tên và điểm dưới dạng từ điển
    return student_scores



def handle(dictt, list_n):

    def cut_n(lst, num_groups):
        sorted_lst = sorted(lst.items(), key=lambda x: x[1], reverse=True)
        groups = [[] for _ in range(num_groups)]
        
        while sorted_lst:
            min_sum = min(sum(score for _, score in group) for group in groups)
            index = [i for i, group in enumerate(groups) if sum(score for _, score in group) == min_sum][0]

            groups[index].append(sorted_lst.pop(0))

        return groups
    num_groups_list = list_n
    
    inputt = dictt
    with open('result.txt', 'w') as f:
        for num_groups in num_groups_list:
            result_groups = cut_n(inputt, num_groups)
            
            f.write(f"$$$$ Chia thành {num_groups} nhóm: $$$$\n")
            for i, group in enumerate(result_groups):
                total = 0
                f.write(f"- Nhóm {i + 1}:\n")
                for student_name, score in group:
                    f.write(f"\t+ {student_name}: {score}\n")
                    total += score
                f.write(f"\t TONG DIEM = {total}")
                f.write("\n")



def main():
    inputt = get_data(input('ten cot ten hoc sinh; '), input('ten cot diem so: '))
    list_n =  [int(x) for x in input("Nhập số lượng nhóm cần chia, cách nhau bằng dấu cách: ").split()]
    handle(inputt,list_n)

main()