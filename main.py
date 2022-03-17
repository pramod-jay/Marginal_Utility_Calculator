import openpyxl

book_que= openpyxl.load_workbook("doc_que.xlsx")
sheet=book_que['Sheet1']
round_trips=[]
round_trips_tu=[]
phone_minutes=[]
phone_minutes_tu=[]
row=sheet.max_row
column=sheet.max_column

for i in range (2,(row+1)):
    round_trips.append(sheet.cell(i,1).value)
    round_trips_tu.append(sheet.cell(i,2).value)
    phone_minutes.append(sheet.cell(i,3).value)
    phone_minutes_tu.append(sheet.cell(i,4).value)

sheet.cell(1,1,'Round Trips')
sheet.cell(1,2,'Total Utility')
sheet.cell(1,3, 'Marginal Utility(Round Trips)')
sheet.cell(1,4, 'MU/P $2')
sheet.cell(1,5,'Phone Minutes')
sheet.cell(1,6,'Total Utility')
sheet.cell(1,7, 'Marginal Utility(Phone Minutes)')
sheet.cell(1,8, 'MU PM $0.05')

print(len(round_trips))

#for i in range(2, len(round_trips)):

book_que.save('book_ans.xlsx')