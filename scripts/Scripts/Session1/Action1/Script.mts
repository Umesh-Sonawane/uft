'Variable Declaration 1
Dim intVar'we use dim keyword to declare any variables.
intVar=10
Dim strVar
strVar="Automation Testing"
print intVar
print strVar
msgbox intVar
msgbox strVar

'Variable Declaration 2
Dim intVar1 : intVar1=10 'This way we can declare the variables and initialize it at the same time.
Dim strVar1 : strVar1="Automation Testing"
print intVar1
print strVar1
msgbox intVar1
msgbox strVar1

'Array Declaration 1
'Method 1 : Using Dim
Dim arr1()	'Without Size
ReDim arr1(3)
arr1(0)=4
arr1(1)=3
arr1(2)=2
arr1(3)=1
print ubound(arr1)
print lbound(arr1)
'Method 2 : Mentioning the Size
Dim arr2(5)  'Declared with size of 5

'Method 3 : using 'Array' Parameter
Dim arr3
arr3 = Array("apple","Orange","Grapes")

'for loop
For i = 0 to 10 step 1
  r = 10 + i
  print r
Next
'while loop
Dim i: i=0
While i <= 5
print i
i = i + 1

Wend

'do while loop	
Dim j : j=0
Do

j = j + 1
print j
Loop Until j > 5
