import pygsheets
from pygsheets.custom_types import HorizontalAlignment
import time
import math

GROUP_COLOR = [(209, 224, 227),(254, 242, 207),(227, 185, 176),(216, 210, 232),(218, 234, 212),(208, 226, 242)]
for i in range(6):
    GROUP_COLOR[i] = tuple([x/255 for x in GROUP_COLOR[i]])



gc = pygsheets.authorize(service_account_file='./intense-arbor-349708-f63f307870a4.json')
url = "https://docs.google.com/spreadsheets/d/1ra4M8irPI0_zCPBqAqL1sLvTT5XGSOuQHQEYMWWGUdA/"
sht = gc.open_by_url(url)

ans = []

# sht.create_worksheet('new_worksheet', rows=100, cols=100)

wks_list = sht.worksheets()
test_wks = wks_list[4]
ansSheet = wks_list[3]
res_sheet = wks_list[1]
score_sheet = wks_list[2]
for i in range(14):
    ans.append(int(ansSheet.cell('b' + str(i+1)).value))
# print(ans)

now_row = 1
with open('test.txt', 'r') as f:
    now_row = int(f.readline())

def init():
    for i in range(6):
        standcell = pygsheets.Cell('A'+str(i+2))
        standcell.color = GROUP_COLOR[i]
        standcell.value = " "
        standcell.set_horizontal_alignment(HorizontalAlignment.CENTER)
        datarange = pygsheets.datarange.DataRange(start='B'+str(i+2),end= 'O'+str(i+2),worksheet=score_sheet)
        datarange.apply_format(standcell)
        datarange.update_values([[" " for _ in range (14)]])
        datarange.link()

# init()
print("Start from Row:", now_row)
while True:
        # time.sleep(1)
    # while res_sheet.cell('A' + str(now_row)).value != None:
        this_row = res_sheet.get_row(now_row)
        print("Row:",now_row,this_row)
        group = this_row[1]
        problem = this_row[2]
        lower = this_row[3] 
        upper = this_row[4] 
        # print(group, problem, upper, lower)
        if group not in ['1', '2', '3', '4', '5', '6']:
            break

        if problem == 15:
            now_row += 1
            continue
        
        try :
            upper = int(upper)
            lower = int(lower)
        except:
            now_row += 1
            continue

        if ( lower <= ans[int(problem)-1] <= upper):
            score = math.floor(upper/lower)
        else:
            score = "X"

        group_index = int(group) +1
        problem_index = int(problem) +1
        problem_char  = chr(64 + problem_index)
        
        cell = pygsheets.Cell(str(problem_char) + str(group_index))
        cell.set_horizontal_alignment(HorizontalAlignment.CENTER)
        cell.value = score
        cell.color = GROUP_COLOR[int(group)-1] if score != "X" else (1, 0, 0)
        cell.link(score_sheet)
        cell.set_text_format('bold', True)
        cell.update()
        print("Row:",now_row,"is done. At",str(problem_char) + str(group_index), "Score:", score)
        now_row += 1


