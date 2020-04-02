DataTable.Import "C:\Users\a107941700\Desktop\uft\scripts\Data\input.xls"
'DataTable.ImportSheet "C:\Users\a107941700\Desktop\uft\scripts\Data\input.xls","Global","Action1"
'DataTable.ImportSheet "C:\Users\a107941700\Desktop\uft\scripts\Data\input.xls","Action1","Global"
For Iterator = 1 To Datatable.GetSheet("Global").GetParameterCount Step 1
	msgbox 	Datatable.GetSheet("Global").GetParameter(Iterator).Name
	msgbox 	Datatable.GetSheet("Global").GetParameter(Iterator).Value
Next
Datatable.GetSheet("Global").AddParameter "Message","TESTING NEW COLUMN"
iRowCount=Datatable.GetSheet("Global").GetRowCount
For Iterator = 1 To iRowCount Step 1
	Datatable.SetCurrentRow(Iterator)
	msgbox DataTable("URL","Global")
	msgbox DataTable("USERID",dtGlobalSheet)
	msgbox DataTable("Password",Global)
	msgbox DataTable("URL","Action1")
	msgbox DataTable("USERID",dtLocalSheet)
	msgbox DataTable("Password","Action1")	
	DataTable("Message",Global)="TESTING NEW COLUMN"
	msgbox DataTable("Message",Global)
Next




'DataTable.Export "C:\Users\a107941700\Desktop\uft\scripts\Data\input.xls"
