'Arrays are declared the same way a variable has been declared except that the declaration of an array variable uses parenthesis.
'In the following example, the size of the array is mentioned in the brackets.
'Method 1 : Using Dim
Dim arr1() 'Without Size

'Method 2 : Mentioning the Size
Dim arr2(5) 'Declared with size of 5
'Although, the Array size is indicated as 5, it can hold 6 values as array index starts from ZERO.

'Method 3 : using 'Array' Parameter
Dim arr3
arr3 = Array("apple","Orange","Grapes")

Dim arr(5)
arr(0) = "1"            'Number as String
arr(1) = "VBScript"     'String
arr(2) = 100            'Number
arr(3) = 2.45           'Decimal Number
arr(4) = #10/07/2013#   'Date
arr(5) = #12.45 PM#     'Time

'Arrays are not just limited to single dimension and can have a maximum of 60 dimensions. 
'Two-dimension arrays are the most commonly used ones.
Dim arr(2,3)	' Which has 3 rows and 4 columns
arr(0,0) = "Apple" 
arr(0,1) = "Orange"
arr(0,2) = "Grapes"           
arr(0,3) = "pineapple" 

arr(1,0) = "cucumber"           
arr(1,1) = "beans"           
arr(1,2) = "carrot"           
arr(1,3) = "tomato"    

arr(2,0) = "potato"             
arr(2,1) = "sandwitch"            
arr(2,2) = "coffee"             
arr(2,3) = "nuts"   

'ReDim Statement is used to declare dynamic-array variables and allocate or reallocate storage space.
'SYNTAX: ReDim [Preserve] varname(subscripts) [, varname(subscripts)]
Dim a()
 i = 0
 redim a(5)
 a(0) = "XYZ"
 a(1) = 41.25
 a(2) = 22

 REDIM PRESERVE a(7)
 For i = 3 to 7
 a(i) = i
 Next

 'to Fetch the output
 For i = 0 to ubound(a)
    Msgbox a(i)
 Next
 
' A Function, which returns a String that contains a specified number of substrings in an array. 
' This is an exact opposite function of Split Method.
' Syntax: Join(List[,delimiter]) 
 a = array("Red","Blue","Yellow")
 b = join(a)
 msgbox("The value of b " & " is :"  & b)

 ' Join using $
 b = join(a,"$")
 
'A Filter Function, which returns a zero-based array that contains a subset of a string array based on a specific filter criteria.
'Syntax: Filter(inputstrings,value[,include[,compare]]) 
 a = array("Red","Blue","Yellow")
 b = Filter(a,"B")
 c = Filter(a,"e")
 d = Filter(a,"Y")

 For each x in b
   msgbox("The Filter result 1: " & x )
 Next

 For each y in c
   msgbox("The Filter result 2: " & y)
 Next

 For each z in d
   msgbox("The Filter result 3: " & z )
 Next
 
'The IsArray Function returns a Boolean value that indicates whether or NOT the specified input variable is an array variable.
'Syntax:IsArray(variablename)
a = array("Red","Blue","Yellow")
b = "12345"
msgbox("The IsArray result 1 : " & IsArray(a) )
msgbox("The IsArray result 2 : " & IsArray(b) )
