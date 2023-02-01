from tkinter import *
from tkinter import filedialog, messagebox, ttk

import tkinter.font as tkFont
import csv, openpyxl

# Tkinter GUI 기본 설정 ----------------------------------------------------------
root = Tk() # Tk 클래스 객체 생성

root.title("Weight CG")
root.geometry("300x350+200+200") # [너비] x [높이] + [x좌표] + [y좌표]
root.resizable(False, False) # 창 크기 변경 불가
# -------------------------------------------------------------------------------


# Progressbar 기본 설정 -------------------------------------------------------------------------------------------
p_var2 = StringVar() #Progressbar 관련 변수
progressbar = ttk.Progressbar(root, maximum=100, mode="determinate", length="280", variable=p_var2) # 상태 진행바
progressbar.place(x=10,y=310) # GUI상 상태 진행바의 위치 지정

fontExample = tkFont.Font(family="arial", size=12, weight="bold") # 폰트 설정
# ----------------------------------------------------------------------------------------------------------------


# ------- 초기값 설정 ---------------------------
security_data_1 = ""
security_data_2 = ""
security_data_3 = ""
# ----------------------------------------------


# OPEN 함수 START -------------------------------------------------------------------------------------------------
def openfile(): # open 버튼 눌러서 파일 선택할 수 있는 함수
    global file # 파일 경로를 해당 py 내에서 공유하기 위한 global 함수
    file = filedialog.askopenfilename(initialdir='./python', title='파일 선택', filetypes=(('csv files', '*.csv'),('all files', '*.*'))) # csv파일 불러올 수 있도록 하는 경로 함수
    file_name_msg.configure(text=" " + file) # 파일 경로를 GUI 상에 띄워주기 위한 변수 file_name_msg에 file 경로 저장.
    file_format = file.split("/") # 파일 format을 확인하기 위해서 경로로 가져온 file directory의 마지막 부분을 잘라냄
    
    if file_format[-1][-3:] == "csv": # 파일 확장자가 csv일 경우
        pass
    else:
        messagebox.showerror("Error", "해당 파일은 csv 파일이 아닙니다.")

# OPEN 함수 END -------------------------------------------------------------------------------------------------


# INS 함수(Input 데이터 입력 받는 함수) START -------------------------------------------------------------------------------------------------
def ins(): # 입력으로 받은 input_A,B,C 데이터 형식을 str -> int로 변경
    global security_data_1, security_data_2, security_data_3 # 사용자 입력 데이터 공유를 위한 global 함수

    # try-except 문
    try: # Input 총 3개 받음.
        security_data_1 = float(input_a.get()) # 업무 보안을 위해 내용 삭제(깃 전용)
        security_data_2 = float(input_b_ua.get()) # 업무 보안을 위해 내용 삭제(깃 전용)
        security_data_3 = float(input_b_pa.get()) # 업무 보안을 위해 내용 삭제(깃 전용)

        messagebox.showinfo("Complete", "Input Data 적용 완료.") # 해당 input 데이터 3개가 모두 정상적으로 입력되었을 때, 완료 메세지 창 
    except:
        messagebox.showerror("Error", "Input Data에 지원되지 않는 형식의 문자나 공백이 있습니다.")
# INS 함수 END -------------------------------------------------------------------------------------------------------------------------------

# CAL 함수 START -------------------------------------------------------------------------------------------------
def cal(): # open 함수로 연 파일 내부 값과 입력으로 넣은 값을 계산하는 함수
    try:
        global security_data_1, security_data_2, security_data_3

        f = open(file, 'r')
        rdr = csv.reader(f)
        list_rows = list(rdr)
        f.close()

        wb = openpyxl.Workbook() # file 변수에 있는 엑셀 파일을 wb 변수에 열기
        wb.active.title = "Blank"
        ws = wb[wb.sheetnames[0]] # Sheet[0] 이름을 가진 시트를 ws 변수에 활성화 하기
        

        wb.create_sheet("업무 보안을 위해 내용 삭제(깃 전용)") # 새로운 시트 만들기
        ws_result = wb["업무 보안을 위해 내용 삭제(깃 전용)"] # 최종 결과물이 나올 시트

        
        # 업무 보안을 위해 내용 삭제(깃 전용)

        if list_rows[0][0] == "DATE": 

            ws_result.cell(row=1, column=1).value = "DATE"
            ws_result.cell(row=1, column=2).value = list_rows[0][1]
            ws_result.cell(row=1, column=3).value = list_rows[0][2]
            ws_result.cell(row=1, column=4).value = list_rows[0][3]
            ws_result.cell(row=1, column=5).value = "업무 보안을 위해 내용 삭제(깃 전용)"
            ws_result.cell(row=1, column=6).value = "업무 보안을 위해 내용 삭제(깃 전용)"
            ws_result.cell(row=1, column=7).value = "업무 보안을 위해 내용 삭제(깃 전용)"

        else:
            messagebox.showerror("Error", "First Parameter(Column B) is not 업무 보안을 위해 내용 삭제(깃 전용).")
            root.destroy()

        max_row = len(list_rows)
        security_data_4 = 0

        for x in range(3, max_row +1): # 1열과 2열에 있는 데이터 입력 데이터와 결합하여 계산 반복
            
            # 업무 보안을 위해 내용 삭제(깃 전용)
            # 업무 보안을 위해 내용 삭제(깃 전용)
            # 업무 보안을 위해 내용 삭제(깃 전용)
            # 업무 보안을 위해 내용 삭제(깃 전용)

            if list_rows[x-1][1] == '' or list_rows[x-1][2] == '' or list_rows[x-1][3] == '':
                pass
            else:
                ws_result.cell(row=x, column=1).value = list_rows[x-1][0]
                ws_result.cell(row=x, column=2).value = list_rows[x-1][1]
                ws_result.cell(row=x, column=3).value = list_rows[x-1][2]
                ws_result.cell(row=x, column=4).value = list_rows[x-1][3]
            
            if list_rows[x-1][1] == '' or list_rows[x-1][2] == '' or list_rows[x-1][3] == '':
                security_data_4 += 1
                
            elif list_rows[x-1][3] == 1 or list_rows[x-1][3] == "1" or list_rows[x-1][3] == "1.000" or list_rows[x-1][3] == 1.000:
                ws_result.cell(row=x, column=5).value = security_data_1 + float(list_rows[x-1][2])
                #  업무 보안을 위해 내용 삭제(깃 전용)

                ws_result.cell(row=x, column=6).value = ( ((security_data_2 * 188.9 / 100 + 334.3) / 12 * security_data_1) + ((float(list_rows[x-1][1]) * float(list_rows[x-1][2])))) / ws_result.cell(row=x, column=5).value
                # 업무 보안을 위해 내용 삭제(깃 전용)
                
                ws_result.cell(row=x, column=7).value = ((ws_result.cell(row=x, column=6).value) * 12 - 334.3) / 188.9 * 100
                # 업무 보안을 위해 내용 삭제(깃 전용)

            elif list_rows[x-1][3] == 0 or list_rows[x-1][3] == "0" or list_rows[x-1][3] == "0.000" or list_rows[x-1][3] == 0.000:
                ws_result.cell(row=x, column=5).value = security_data_1 + float(list_rows[x-1][2])
                #  업무 보안을 위해 내용 삭제(깃 전용)

                ws_result.cell(row=x, column=6).value = ( ((security_data_3 * 188.9 / 100 + 334.3) / 12 * security_data_1) + ((float(list_rows[x-1][1]) * float(list_rows[x-1][2])))) / ws_result.cell(row=x, column=5).value
                # 업무 보안을 위해 내용 삭제(깃 전용)
                
                ws_result.cell(row=x, column=7).value = ((ws_result.cell(row=x, column=6).value) * 12 - 334.3) / 188.9 * 100
                # 업무 보안을 위해 내용 삭제(깃 전용)

            else:
                security_data_4 += 1
                # if security_data_4 == 50000:
                #     messagebox.showerror("Error", "업무 보안을 위해 내용 삭제(깃 전용) Parameter(Column D) is not 0 or 1.")
                #     return
            
            p_var2.set((x/(max_row-1))*100) # 해당 for문 반복 %에 따라서 상태 진행바 채우기
            progressbar.update() # for %에 따라 데이터 업데이트

        ws_result.delete_rows(3,security_data_4)
        wb.remove_sheet(wb[wb.sheetnames[0]])
        wb.save("save_data.csv") # 파일 이름 저장
        wb.close() # 열messagebox.showinfo("Complete", "Merge Complete")려있는 excel 닫기
        messagebox.showinfo("Complete", "Merge Complete") # 완료 메세지 출력
        root.destroy()
    except TypeError:
        messagebox.showerror("Error", "Input Data 적용 버튼을 누르지 않았습니다.")
# CAL 함수 END -------------------------------------------------------------------------------------------------

# Directory 이름 넣기
file_dir = Label(root, text="Directory")
file_dir.place(x=9, y=9)

# Open을 열린 파일 경로를 나타내 줄 하얀색 입력창 생성
file_name_msg = Label(root, text=" ", width=33,bg='white', relief="solid", borderwidth=1, anchor="w")
file_name_msg.place(x=10, y=34)

# open이라 쓰여있는 버튼 클릭 시 위에서 정의한 open 함수 실행
open_btn = Button(root, text='open', command=openfile)
open_btn.place(x=250, y=30)

# 구분선
sp1 = ttk.Separator(root, orient="horizontal")
sp1.place(relx=0.01, rely=0.2, relheight=0.1, relwidth=0.95)

# input Data 이름 넣기
input_text = Label(root, text="Input Data", )
input_text.place(x=100, y=80)
input_text.configure(font=fontExample)

# Input a,b,c 이름과 데이터 입력 창 삽입 -----------------------------
input_a_text = Label(root, text="Float Data 1 : ")
input_a_text.place(x=19, y=120)

input_a = Entry(root, width=22)
input_a.place(x=120, y=120)

input_b_text = Label(root, text="Float Data 2  :                UA -  ")
input_b_text.place(x=19, y=150)
input_b2_text = Label(root, text="PA -")
input_b2_text.place(x=162, y=175)

input_b_ua = Entry(root, width=11)
input_b_ua.place(x=197, y=150)

input_b_pa = Entry(root, width=11)
input_b_pa.place(x=197, y=175)

# input_c_text = Label(root, text="C Data : ")
# input_c_text.place(x=9, y=180)
# input_c = Entry(root, width=15)
# input_c.place(x=69, y=180)
# ------------------------------------------------------------------

# 구분선
sp1 = ttk.Separator(root, orient="horizontal")
sp1.place(relx=0.01, rely=0.72, relheight=0.1, relwidth=0.95)

# 입력한 input data를 위에서 정의한 ins 함수를 통해 int값으로 변환 후 데이터 저장
ins_btn = Button(root, width=35, height=1, text=' Input Data 적용 ', command=ins)
ins_btn.place(x=20, y=210)

# 파일 생성 버튼을 클릭하면 위에서 정의한 cal 함수 실행
add_btn = Button(root,width=20, height=1, text=' 파일 생성 ', command=cal)
add_btn.place(x=70, y=270)

version_number = Label(root, text="Version 1.1")
version_number.place(x=230, y=332)

# 끝
root.mainloop()