'Create excel object to read
Set objReadXL = CreateObject("Excel.Application")
objReadXL.Visible = True

Set objReadWB = objReadXL.Workbooks.Open("E:\Scripts\DrawBorder.xlsx") 'Change the path of your excel file
Set objReadWS = objReadWB.Sheets("Sheet1") 'Change the name of sheet as in your file

objReadWS.Range(objReadWS.Cells(1, 1), objReadWS.Cells(9, 4)).BorderAround 1
objReadWS.Range(objReadWS.Cells(1, 1), objReadWS.Cells(9, 4)).Borders(12).LineStyle = 1
objReadWS.Range(objReadWS.Cells(1, 1), objReadWS.Cells(9, 4)).Borders(11).LineStyle = 1

objReadWB.Save

Set objReadWS = Nothing
Set objReadWB = Nothing
Set objReadXL = Nothing
