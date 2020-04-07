'systemutil.Run "chrome.exe","http://newtours.demoaut.com/"
'wait(15)
'msgbox "Recovery Scenario In Process"
'
'
'
'msgbox "Recovery Scenario Executed"

On Error Resume Next '(Putting error handling statement)
Dim result
result = 20/0 '(Performing division by 0 Scenario)
If not isempty(result) Then '(Checking the value of result variable)
	Msgbox "Result is 0."
Else
	Msgbox "Result is non-zero."
End If
msgbox err.Number & " " & err.Description
err.clear
msgbox err.Number & " " & err.Description

On error goto 0 'disable on error resume next
'Dim result
result = 20/0 '(Performing division by 0 Scenario)
If not isempty(result) Then '(Checking the value of result variable)
	Msgbox "Result is 0."
Else
	Msgbox "Result is non-zero."
End If
