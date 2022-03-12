Dim args, objExcel

set args = wscript.Arguments
set objExcel= CreateObject("Excel.Application")

objExcel.workbooks.open args(0)
objExcel.visible = True

objExcel.Run "Main"

objExcel.Activeworkbook.Save
objExcel.Activeworkbook.Close(0)
objExcel.Quit

