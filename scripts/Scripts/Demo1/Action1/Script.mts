'interger variable to store number
Dim intvar
intvar=10
print intvar
msgbox intvar

'String variable to store name 
Dim StrName
StrName="Swati"
print StrName
msgbox StrName

Dim Intvar1 : Intvar1=30   'to store the variable by using one statement
print Intvar1
msgbox Intvar1


Dim Strname1 : Strname1="Automation Testing"
print Strname1
msgbox Strname1


Dim intarr1()	'Without Size
ReDim intarr1(3)
intarr1(0)=1
intarr1(1)=2
intarr1(2)=3
intarr1(3)=4
print ubound(intarr1)
print lbound(intarr1)

Dim intarr2(5)  

'Method 3 : using 'Array' Parameter
Dim arr3
arr3 = Array("apple","Orange","Grapes")

For Iterator = 0 To ubound(arr3) Step 1
	
	print arr3(Iterator)
Next


'for Loops to print 10 numbers
Dim intvar3
For intvar3 = 1 To 10 Step 1
	
	print intvar3
Next

'while loop to print 10 number

Dim intvar4
intvar4=10
While(intvar4<=15)
print intvar4
intvar4=intvar4+1
	
Wend


i=20
do
i=i+1
print i
Loop Until i > 25

for i=0 to ubound(arr)
	print arr(i)
next

for i=ubound(arr) to 0 step -1
	print arr(i)
next


'perform addition of two numbers
dim a,b,c
a=cint(inputbox("Enter first no"))
b=cint(inputbox("Enter second no"))
'c=a+b
'call AddTwoNumbers(a,b)
msgbox AddTwoNumbers(a,b)
msgbox AddTwoNumbers(10,30)



Function AddTwoNumbers(a,b)
	AddTwoNumbers=a+b
End Function
