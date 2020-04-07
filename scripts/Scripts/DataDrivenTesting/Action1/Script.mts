Option Explicit
Dim sUsername,sPassword,sUrl,sTriptype,sFromDay,sToDay,sToMonth,Iterator,sFromport,sToport,sServclass,sAirline,sOutflight,sInflight,sPassfirst0,sPasslast0,iCreditnumber,sMeal,iExpirydate,iExpiryyear,sFname,sLname
sUsername="mercury"
sPassword="5e8c0e37c4a1b01a2d6a5f278e8f3cf622a648b9"
sUrl="http://newtours.demoaut.com/"

Datatable.Import "C:\Users\a107941700\Desktop\uft\scripts\Data\DataDrivenTesting.xls"

For Iterator = 1 To Datatable.GetSheet("Global").GetRowCount Step 1
	DataTable.SetCurrentRow Iterator
	LoginMercuryDemoTours sUrl,sUsername,sPassword

	sTriptype=DataTable("triptype","Global") @@ script infofile_;_ZIP::ssf4.xml_;_
	sFromDay=DataTable("fromday","Global")
	sToDay=DataTable("today","Global")
	sToMonth=DataTable("tomonth","Global")
	
	sFromport=DataTable("fromport","Global")
	sToport=DataTable("toport","Global")
	sServclass=DataTable("servclass","Global")
	sAirline=DataTable("airline","Global")
	
	sOutflight=DataTable("outFlight","Global")
	sInflight=DataTable("inflight","Global")
	sPassfirst0=DataTable("passfirst0","Global")
	sPasslast0=DataTable("passlast0","Global")
	
	iCreditnumber=DataTable("creditnumber","Global")
	sMeal=DataTable("meal","Global")
	iExpirydate=DataTable("expirydate","Global")
	iExpiryyear=DataTable("expiryyear","Global")
	sFname=DataTable("fname","Global")
	sLname=DataTable("lname","Global")
	
	DataTable("orderno","Global")= replace(BookFlight(sTriptype,sFromDay,sToDay,sToMonth,sFromport,sToport,sServclass,sAirline,sOutflight,sInflight,sPassfirst0,sPasslast0,iCreditnumber,sMeal,iExpirydate,iExpiryyear,sFname,sLname),"Flight Confirmation #","")
Next
DataTable.Export "c:\temp\output.xls"

