'A for loop is a repetition control structure that allows a developer to efficiently 
'write a loop that needs to execute a specific number of times.

'Syntax
'The syntax of a for loop in VBScript is −

'For counter = start To end [Step stepcount]
'   [statement 1]
'   [statement 2]
'   ....
'   [statement n]
'   [Exit For]
'   [statement 11]
'   [statement 22]
'   ....
'   [statement n]
'Next


'Here is the flow of control in a For Loop −

'1.The For step is executed first. This step allows you to initialize any loop control variables and increment the step counter variable.

'2.Secondly, the condition is evaluated. If it is true, the body of the loop is executed. If it is false, the body of the loop does not execute and flow of control jumps to the next statement just after the For Loop.

'3.After the body of the for loop executes, the flow of control jumps to the Next statement. This statement allows you to update any loop control variables. It is updated based on the step counter value.

'4.The condition is now evaluated again. If it is true, the loop executes and the process repeats itself (body of loop, then increment step, and then again condition). After the condition becomes false, the For Loop terminates.

Dim a : a = 10
For i = 0 to a Step 2 'i is the counter variable and it is incremented by 2
  msgbox("The value is i is : " & i)
Next
 
 
' A Exit For Statement is used when we want to Exit the For Loop based on certain criteria. When Exit For is executed, the control jumps to next statement immediately after the For Loop.

'Syntax
'The syntax for Exit For Statement in VBScript is −

 'Exit For
 
 Dim a : a = 10
 For i = 0 to a Step 2 'i is the counter variable and it is incremented by 2
    msgbox("The value is i is : " & i)
    

   If i = 4 Then
      i = i*10  'This is executed only if i = 4
      msgbox(" Exited The value is i is : " & i)
      Exit For 'Exited when i = 4
   End If	
Next
 
' A For Each loop is used when we want to execute a statement or a group of statements for each element in an array or collection.

'A For Each loop is similar to For Loop; however, the loop is executed for each element in an array or group. Hence, the step counter won't exist in this type of loop and it is mostly used with arrays or used in context of File system objects in order to operate recursively.

'Syntax
'The syntax of a For Each loop in VBScript is −

'For Each element In Group
'   [statement 1]
'   [statement 2]
'   ....
'   [statement n]
'   [Exit For]
'   [statement 11]
'   [statement 22]
'Next
'Example

'fruits is an array 
fruits = Array("apple","orange","cherries")
Dim fruitnames

'iterating using For each loop. 
For each item in fruits
  fruitnames = fruitnames&item&vbnewline
Next

msgbox fruitnames

'When the above code is executed, it prints all the fruitnames with one item in each line.


'In a While..Wend loop, if the condition is True, all statements are executed until Wend keyword is encountered.

'If the condition is false, the loop is exited and the control jumps to very next statement after Wend keyword.

'Syntax
'The syntax of a While..Wend loop in VBScript is −

'While condition(s)
 '  [statements 1]
 '  [statements 2]
 '  ...
 '  [statements n]
'Wend

Dim Counter :  Counter = 10   
 While Counter < 15    ' Test value of Counter.
    Counter = Counter + 1   ' Increment Counter.
    msgbox("The Current Value of the Counter is : " & Counter)    
 Wend ' While loop exits if Counter Value becomes 15.
         
'An Exit Do Statement is used when we want to Exit the Do Loops based on certain criteria. It can be used within both Do..While and Do..Until Loops.

'When Exit Do is executed, the control jumps to next statement immediately after the Do Loop.

'Syntax
'The syntax for Exit Do Statement in VBScript is −

 'Exit Do
 i = 0
 Do While i <= 100
    If i > 10 Then
       Exit Do   ' Loop Exits if i>10
    End If
    msgbox("The Value of i is : " &i)
    
    i = i + 2
 Loop   
 
 'A Do..While loop is used when we want to repeat a set of statements as long as the condition is true. The Condition may be checked at the beginning of the loop or at the end of the loop.

'Syntax
'The syntax of a Do..While loop in VBScript is −

'Do While condition
'  [statement 1]
'   [statement 2]
 '  ...
'   [statement n]
'   [Exit Do]
'   [statement 1]
'   [statement 2]
'   ...
 '  [statement n]
'Loop    
i=0
 Do While i < 5
    i = i + 1
    msgbox("The value of i is : " & i)
    
 Loop   
 
 There is also an alternate Syntax for Do..while loop which checks the condition at the end of the loop. The Major difference between these two syntax is explained below with an example.

'Do 
 '  [statement 1]
 '  [statement 2]
 '  ...
 '  [statement n]
 '  [Exit Do]
 '  [statement 1]
 '  [statement 2]
 '  ...
 '  [statement n]
'Loop While condition
i=0
 Do 
    i = i + 1
    msgbox("The value of i is : " & i)
    
 Loop While i > 5
 
' A Do..Until loop is used when we want to repeat a set of statements as long as the condition is false. The Condition may be checked at the beginning of the loop or at the end of loop.

'Syntax
'The syntax of a Do..Until loop in VBScript is −

'Do Until condition
'   [statement 1]
'   [statement 2]
 '  ...
 '  [statement n]
  ' [Exit Do]
 '  [statement 1]
  ' [statement 2]
 '  ...
'   [statement n]
'Loop           

i = 10
Do Until i>15  'Condition is False.Hence loop will be executed
  i = i + 1
  msgbox("The value of i is : " & i)
  
Loop 
