1.
'VBScript variables are declared using the dim statement:
 ' Dim firstName
 ' Dim age
2.
 'These variables could also have been declared on the same line, separated by a comma:
 ' Dim firstName, age
3. 
 'VBScript has a way of ensuring you declare all your variables. This is done with the Option Explicit statement:
' Option Explicit
'  Dim firstName
'  Dim age
'  firstName = "Borat"
'  age = 25

'Constant is a named memory location used to hold a value that CANNOT be changed during the script execution. 
'If a user tries to change a Constant Value, the Script execution ends up with an error. 
'Constants are declared the same way the variables are declared.
'Syntax:[Public | Private] Const Constant_Name = Value

'Example 1
'In this example, the value of pi is 3.4 and it displays the area of the circle in a message box.
 Dim intRadius
 intRadius = 20
 const pi = 3.14
 Area = pi*intRadius*intRadius
 Msgbox Area

'Example 2
'The below example illustrates how to assign a String and Date Value to a Constant.
Const myString = "VBScript"
Const myDate = #01/01/2050#
Msgbox myString
Msgbox myDate

'Example 3
'In the below example, the user tries to change the Constant Value; hence, it will end up with an Execution Error.
Dim intRadius
intRadius = 20
const pi = 3.14
pi = pi*pi	'pi VALUE CANNOT BE CHANGED.THROWS ERROR'
Area = pi*intRadius*intRadius
Msgbox Area
