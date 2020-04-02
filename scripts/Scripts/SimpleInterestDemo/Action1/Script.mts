Dim doublePrinciple,doubleyear,doublerate,dblintrest,dbltotalamount
doublePrinciple=cdbl(inputbox("Enter Principle amount"))
doubleyear=cdbl(inputbox("Enter no of years"))
doublerate=cdbl(inputbox("Enter rate"))
dblintrest=CalculateSimpleIntrest(doublePrinciple,doubleyear,doublerate)
print "Total Intrest is " & dblintrest
msgbox "Total Intrest is " & dblintrest
dbltotalamount=pdoublePrinciple+intintrest
print "Total Amount is " & dbltotalamount
msgbox "Total Amount is " & dbltotalamount




Function CalculateSimpleIntrest(doublePrinciple,doubleyear,doublerate)
	CalculateSimpleIntrest=(doublePrinciple*doubleyear*doublerate)/100
End Function

