Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open("/home/runner/workspace/processed_inventory.xlsx")
objExcel.Run "AutoPivot"
objWorkbook.Save
objWorkbook.Close False
objExcel.Quit