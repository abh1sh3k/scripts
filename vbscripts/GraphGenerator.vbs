'Create excel object to read
Set objReadXL = CreateObject("Excel.Application")
objReadXL.Visible = True

Set objReadWB = objReadXL.Workbooks.Open("E:\Scripts\DrawBorder.xlsx") 'Change the path of your excel file
Set objReadWS = objReadWB.Sheets("Sheet1") 'Change the name of sheet as in your file

Set objChart = objReadXL.Charts.Add()
Set rng = objReadWS.Range(objReadWS.Cells(1, 1), objReadWS.Cells(9, 4))
objReadXL.ActiveChart.HasTitle = True
objReadXL.ActiveChart.ChartTitle.Characters.Text = "CPU Utilization"
objReadXL.ActiveChart.SetSourceData rng, 2
objReadXL.ActiveChart.Location 2, "Sheet2"
objReadXL.ActiveChart.Axes(1).HasTitle = True
objReadXL.ActiveChart.Axes(1).AxisTitle.Characters.Text = "Load Duration"
objReadXL.ActiveChart.Axes(2).HasTitle = True
objReadXL.ActiveChart.Axes(2).AxisTitle.Characters.Text = "Percentage"
objReadXL.ActiveChart.ChartType = 4
objReadXL.ActiveChart.Parent.Width = 350
objReadXL.ActiveChart.Parent.Height = 200
objReadXL.ActiveChart.Parent.Left = xLeft
objReadXL.ActiveChart.Parent.Top = xTop

objReadWB.Save

Set objReadWS = Nothing
Set objReadWB = Nothing
Set objReadXL = Nothing
