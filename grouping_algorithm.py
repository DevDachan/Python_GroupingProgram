import random
import time
import math
import copy
import pandas as pd
pd.set_option('mode.chained_assignment',  None) # 경고 off


filename = "example.xlsx"
df = pd.read_excel(filename, engine="openpyxl", sheet_name=["전임"])["전임"]
df2 = pd.read_excel(filename, engine="openpyxl", sheet_name=["신임"])["신임"]

## N == 전체 팀 개수
N = 8
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

## 전체 팀에 들어갈 사람 수를 정하기 위한 과정입니다.

print(sort_data)

print(sort_new)

## team_arr는 최종적으로 팀 구성원이 저장 될 배열 입니다.
team_arr = [[0 for j in range(4)] for i in range(N)]
## index 0: 팀 번호
## index 1: 전체 나이 합
## index 2: 성비
## index 3: 사람 수

for i in range(N):
    team_arr[i][0] = i+1


## True = Ascending,  False = Descending
order = True
team_index = 1

sort_result = pd.DataFrame(columns = [])

df_temp = pd.DataFrame(columns = [])
select_temp = 0



total_age = 0



for i in range(len(sort_data)):

    df_temp = df_temp.append(sort_data.iloc[i], ignore_index = True)
    team_index += 1;


    if team_index > N or i == len(sort_data)-1:
        if order == True:
            df_temp = df_temp.sort_values(by=["나이"])
            for k in range(len(df_temp)):
                df_temp.loc[k, ["팀 번호"]] = k+1

                team_arr[k+1][1] += df_temp.loc[k,"나이"]
                team_arr[k+1][3] += 1

                total_age += df_temp.loc[k,"나이"]
                if df_temp.loc[k, "성별"] == "남":
                        team_arr[k][2] += 1
                elif df_temp.loc[k, "성별"] == "여":
                        team_arr[k][2] -= 1


            sort_result = sort_result.append(df_temp, ignore_index = True)
            order = False
            df_temp = pd.DataFrame(columns = [])

        else:
            df_temp = df_temp.sort_values(by=["나이"], ascending=False)
            for k in range(len(df_temp)):
                df_temp.loc[k, ["팀 번호"]] = k+1

                team_arr[k+1][1] += df_temp.loc[k,"나이"]
                team_arr[k+1][3] += 1

                total_age += df_temp.loc[k,"나이"]
                if df_temp.loc[k, "성별"] == "남":
                        team_arr[k][2] += 1
                elif df_temp.loc[k, "성별"] == "여":
                        team_arr[k][2] -= 1

            sort_result = sort_result.append(df_temp, ignore_index = True)
            order = True
            df_temp = pd.DataFrame(columns = [])
        team_index = 1


sort_result = sort_result.sort_values(by = "팀 번호")



result = pd.DataFrame(team_arr,columns=['팀 번호','나이 합','성비', "사람 수"])
result = result.sort_values(by=["사람 수","나이 합"], ascending=True)
result = result.reset_index(drop=True)



sort_result.to_excel("result.xlsx", index=False)
