
'perform addition of two numbers
dim a,b,c
a=cdbl(inputbox("Enter first no"))
b=cdbl(inputbox("Enter second no"))
c=cdbl(inputbox("Enter Third No"))

'a=cint(inputbox("Enter first no"))
'b=cint(inputbox("Enter second no"))
'c=cint(inputbox("Enter Third No"))
'c=a+b
'call AddTwoNumbers(a,b)
'msgbox AddTwoNumbers(a,b)
msgbox AddTwoNumbers("Automation","Testing")
msgbox MultiplyThreeNumbers(a,b,c)
msgbox MultiplyThreeNumbers(2,3,4)
msgbox SubtractTwoNumbers(a,b)

Function AddTwoNumbers( a, b)
	AddTwoNumbers=a+b
End Function

Function MultiplyThreeNumbers(a,b,c)
MultiplyThreeNumbers=a*b*c
	
End Function


Function SubtractTwoNumbers(a,b)
	SubtractTwoNumbers=a-b
End Function
