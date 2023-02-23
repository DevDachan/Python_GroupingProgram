import random
import time
import math
import copy
import pandas as pd
from tkinter import *

pd.set_option('mode.chained_assignment',  None) # 경고 off

tk = Tk()
tk.title("조 편성")
tk.geometry("400x400")

filename = "전체 명단.xlsx"
df = pd.read_excel(filename, engine="openpyxl", sheet_name=["명단"])["명단"]
df2 = pd.read_excel(filename, engine="openpyxl", sheet_name=["신임"])["신임"]



def event():
    ## N == 전체 팀 개수
    N = int(entry.get())
    print(N)
    ## total_len == 전체 사람 수
    total_len = len(df) + len(df2)


    ## sort_data = 전임 교수 명단
    sort_data = df.sort_values(by=["학부","성별","나이"])
    sort_data = sort_data.reset_index(drop=True)
    sort_data["팀 번호"] = 0

    ## sort_new = 신임 교수 명단
    sort_new = df2.sort_values(by=["학부","성별","나이"], ascending=False)
    sort_data = sort_data.reset_index(drop=True)
    sort_new["팀 번호"] = 0


    ## team_arr는 최종적으로 팀 구성원이 저장 될 배열 입니다.
    team_arr = [[0 for j in range(5)] for i in range(N)]
    ## index 0: 팀 번호
    ## index 1: 전체 나이 합
    ## index 2: 남
    ## index 3: 여
    ## index 4: 사람 수

    for i in range(N):
        team_arr[i][0] = i+1



    ## True = Ascending,  False = Descending
    order = True
    team_index = 1

    sort_result = pd.DataFrame(columns = [])

    df_temp = pd.DataFrame(columns = [])



    for i in range(len(sort_data)):

        df_temp = df_temp.append(sort_data.iloc[i], ignore_index = True)
        team_index += 1;


        if team_index > N or i == len(sort_data)-1:
            ## 전체 나이 순 번호 정렬 내림차순
            team_arr.sort(key=lambda x: -x[1])

            ## 추출한 data 나이 순 정렬 오름차순
            df_temp = df_temp.sort_values(by=["나이"])
            df_temp = df_temp.reset_index(drop=True)


            for k in range(len(df_temp)):
                df_temp.loc[k, ["팀 번호"]] = team_arr[k][0]

                team_arr[k][1] += df_temp.loc[k,"나이"]
                team_arr[k][4] += 1

                if df_temp.loc[k, "성별"] == "남":
                        team_arr[k][2] += 1
                elif df_temp.loc[k, "성별"] == "여":
                        team_arr[k][3] += 1


            sort_result = sort_result.append(df_temp, ignore_index = True)

            df_temp = pd.DataFrame(columns = [])

            team_index = 1



    ## 남녀비율 맞추기

    team_arr.sort(key=lambda x: -x[2])

    conv = True
    conv_index_front = 0
    conv_index_back = len(team_arr)-1


    for i in range(len(sort_new)):
        if conv == True:
            if sort_new.loc[i, "성별"] == "남":
                sort_new.loc[i, ["팀 번호"]] = team_arr[conv_index_front][0]
                conv_index_front += 1
                team_arr[conv_index_front][1] += sort_new.loc[i, "나이"]

                team_arr[conv_index_front][2] += 1

                team_arr[conv_index_front][4] += 1


            elif sort_new.loc[i, "성별"] == "여" :
                sort_new.loc[i, ["팀 번호"]] = team_arr[conv_index_back][0]
                conv_index_back -= 1
                team_arr[conv_index_back][1] += sort_new.loc[i, "나이"]

                team_arr[conv_index_back][3] += 1

                team_arr[conv_index_back][4] += 1



    sort_result = sort_result.append(sort_new, ignore_index = True)

    sort_result = sort_result.sort_values(by = ["팀 번호", "학부"])
    sort_result = sort_result.reset_index(drop=True)


    statistics = pd.DataFrame(columns = ["팀 번호","남", "여", "나이 평균", "학부 수"])

    total_dep = pd.DataFrame(columns = ["학부"])

    cur_team = 0
    cur_dep = ""
    total_count = 0
    dep_count = 0

    ## 팀 번호 정렬 내림차순
    team_arr.sort(key=lambda x: x[0])



    for i in range(len(sort_result)):

        if i == 0:
            cur_team = sort_result.loc[i,"팀 번호"]-1
            cur_dep = sort_result.loc[i, "학부"]
            total_dep.loc[dep_count, "학부"] = str(cur_team+1) + "팀"
            dep_count += 1

        elif sort_result.loc[i, "팀 번호"] != (cur_team+1) or i == len(sort_result)-1:
            if (cur_team+1) < N:
                total_dep.loc[dep_count, "학부"] = str(cur_team+2) + "팀"
                dep_count += 1

            statistics.loc[cur_team, "학부 수"] = total_count
            statistics.loc[cur_team, "나이 평균"] = round(team_arr[cur_team][1]/team_arr[cur_team][4], 1)
            statistics.loc[cur_team, "팀 번호"] = team_arr[cur_team][0]
            statistics.loc[cur_team, "남"] = team_arr[cur_team][2]
            statistics.loc[cur_team, "여"] = team_arr[cur_team][3]

            cur_team = sort_result.loc[i, "팀 번호"]-1
            cur_dep = sort_result.loc[i, "학부"]
            total_count = 0
        else:

            if sort_result.loc[i, "학부"] != cur_dep:
                total_count += 1
                total_dep.loc[dep_count, "학부"] = sort_result.loc[i, "학부"]
                dep_count += 1


    sort_result = sort_result.sort_values(by = ["팀 번호", "성별", "나이", "학부"])

    print(total_dep)

    writer = pd.ExcelWriter("조 편성 결과.xlsx", engine='openpyxl')
    sort_result.to_excel(writer, index=False, sheet_name="전체 편성표")
    statistics.to_excel(writer, index=False, sheet_name="평균")
    total_dep.to_excel(writer, index=False, sheet_name="학부 구성")

    writer.save()

    label2.config(text = "생성 완료!")
## UI


label_content = Label(tk, text="전체 교수님 명단에서 조를 짜는 프로그램 입니다. \n 파일 이름을 전체 명단.xlsx로 바꿔주세요 \n 전체 sheet는 명단, 신임으로 이루어집니다. \n 각 열은 이름, 학부, 나이, 성별 순으로 입력해주세요").grid(row=0, columnspan=2, padx=40, pady=40)



label1 = Label(tk, text="전체 팀 개수(ex) 16)").grid(row=1, column=0, padx=40, pady=40)

entry = Entry(tk)
entry.grid(row=1, column=1, pady=40)

label2 = Label(tk, text=" ")
label2.grid(row=2, columnspan=2)


button = Button(tk, text="생성하기" , command=event).grid(row=3, columnspan=2)

tk.mainloop()
