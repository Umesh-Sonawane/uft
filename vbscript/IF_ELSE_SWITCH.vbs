'Decision making allows programmers to control the execution flow of a script or one of its sections. The execution is governed by one or more conditional statements.

'Following is the general form of a typical decision making structure found in most of the programming languages −

'VBScript provides the following types of decision making statements.

'Statement	Description
'if statement	An if statement consists of a Boolean expression followed by one or more statements.
'if..else statement	An if else statement consists of a Boolean expression followed by one or more statements. If the condition is True, the statements under the If statements are executed. If the condition is false, then the Else part of the script is Executed
'if...elseif..else statement	An if statement followed by one or more ElseIf Statements, that consists of Boolean expressions and then followed by an optional else statement, which executes when all the condition becomes false.
'nested if statements	An if or elseif statement inside another if or elseif statement(s).
'switch statement	A switch statement allows a variable to be tested for equality against a list of values.

'1. An If statement consists of a Boolean expression followed by one or more statements. If the condition is said to be True, the statements under If condition(s) are Executed. If the Condition is said to be False, the statements after the If loop are executed.

'Syntax
'The syntax of an If statement in VBScript is −

'If(boolean_expression) Then
'  Statement 1
'	.....
'	.....
'   Statement n
'End If

'Example
Dim a : a = 20
Dim b : b = 10

If a > b Then
  msgbox "a is Greater than b"
End If

'2.An If statement consists of a Boolean expression followed by one or more statements. If the condition is said to be True, the statements under If condition(s) are Executed. If the Condition is said to be False, the statements under Else Part would be executed.

'Syntax
'The syntax of an if…else statement in VBScript is −

'If(boolean_expression) Then
'   Statement 1
'	.....
'	.....
'   Statement n
'Else
 '  Statement 1
'	.....
'	....
'   Statement n
'End If
Dim a : a = 5
Dim b : b = 25
If a > b Then
  msgbox "a is Greater"
Else 
  msgbox "b is Greater"
End If

'An If statement followed by one or more ElseIf Statements that consists of boolean expressions and then followed by a default else statement, which executes when all the condition becomes false.

'Syntax
'The syntax of an If-ElseIf-Else statement in VBScript is −

'If(boolean_expression) Then
'   Statement 1
'	.....
'	.....
'   Statement n
'ElseIf (boolean_expression) Then
'   Statement 1
'	.....
'	....
'   Statement n
'ElseIf (boolean_expression) Then
'   Statement 1
'	.....
'	....
'   Statement n
'Else
'   Statement 1
'	.....
'	....
'   Statement n
'End If

Dim a  
a = -5

If a > 0 Then
  msgbox "a is a POSITIVE Number"
ElseIf a < 0 Then
  msgbox "a is a NEGATIVE Number"
Else
  msgbox "a is EQUAL than ZERO"
End If

'An If or ElseIf statement inside another If or ElseIf statement(s). The Inner If statements are executed based on the Outermost If statements. This enables VBScript to handle complex conditions with ease.

'Syntax
'The syntax of a Nested if statement in VBScript is −

'If(boolean_expression) Then
'   Statement 1
'	.....
'	.....
'   Statement n
'   If(boolean_expression) Then
'      Statement 1
'		.....
'		.....
'	  Statement n
'   ElseIf (boolean_expression) Then
'      Statement 1
'		.....
'		....
'      Statement n
'   Else
'	   Statement 1
'		.....
'		....
'	   Statement n
'   End If
'Else
'   Statement 1
'	.....
'	....
'   Statement n
'End If
'Example
Dim a
a = 23

If a > 0 Then
  msgbox "The Number is a POSITIVE Number"
  If a = 1 Then
     msgbox "The Number is Neither Prime NOR Composite"   
  Elseif a = 2 Then
     msgbox "The Number is the Only Even Prime Number"   
  Elseif a = 3 Then
     msgbox "The Number is the Least Odd Prime Number"   
  Else
     msgbox "The Number is NOT 0,1,2 or 3"   
  End If
ElseIf  a < 0 Then
  msgbox "The Number is a NEGATIVE Number"
Else
  msgbox "The Number is ZERO"
End If

'When a user wants to execute a group of statements depending upon a value of an expression, then he can use Select Case statements.
'Each value is called a Case, and the variable being switched ON based on each case.
'Case Else statement is executed if test expression doesn't match any of the Case specified by the user.
'Case Else is an optional statement within Select Case, however, it is a good programming practice to always have a Case Else statement.

'Syntax
'The syntax of a Select Statement in VBScript is −

'Select Case expression
'   Case expressionlist1
'      statement1
'      statement2
'      ....
'      ....
'      statement1n
'   Case expressionlist2
'      statement1
'      statement2
'      ....
'      ....
'   Case expressionlistn
'      statement1
'      statement2
'      ....
'      ....   
'  Case Else
'      elsestatement1
'      elsestatement2
'      ....
'      ....
'End Select
'Example

Dim MyVar
MyVar = 1

Select case MyVar
case 1
   msgbox  "The Number is the Least Composite Number"

case 2
   msgbox "The Number is the only Even Prime Number"

case 3
   msgbox "The Number is the Least Odd Prime Number"

case else
   msgbox "Unknown Number"
End select

'In the above example, the value of MyVar is 1. Hence, Case 1 would be executed.
