


msgbox "Addition is" & inttotal 

print "multiplication is " & intmultiply
msgbox "multiplication is " & intmultiply



Dim intarr5(5)
intTotal=0

For Iterator = 0 To ubound(intarr5) 
	
intarr5(Iterator)=cint(inputbox("Enter 5 Numbers"))
intTotal=intTotal+intarr5(Iterator)
Next
print "Addition is " & intTotal


Dim intnumber

intnumber=cint(inputbox("Enter The Numbers"))
'For i = intnumber to 1 step -1
'	
'	print " Numbers are " & i
'Next
'
Dim strpattern : strpattern=""
For i = intnumber To 1 Step -1
	strpattern=""
	For j = intnumber To i Step -1
		strpattern=strpattern & cstr(j)
		print strpattern 
	
	Next
Next

Dim intnumber

intnumber=cint(inputbox("Enter The Numbers"))

Dim strpattern : strpattern=""
Dim strFinalpattern : strFinalpattern=""
For i = intnumber To 1 Step -1
	strpattern=""
	For j = intnumber To i Step -1
		strpattern=strpattern &" " & cstr(j)
'		print strpattern -1
	
	Next
	strFinalpattern=strFinalpattern & vbnewline & strpattern
Next
print strFinalpattern
