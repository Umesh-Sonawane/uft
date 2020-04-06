Systemutil.Run "chrome.exe","http://newtours.demoaut.com/"

wait 10
'Creating a description object


Set objBrowser = Description.Create()
'Add descriptions and properties
objBrowser("micclass").value = "Browser"
objBrowser("name").value = "Welcome: Mercury Tours"
objBrowser("title").value = "Welcome: Mercury Tours"

Set objBrowsersList=Desktop.ChildObjects(objBrowser)

msgbox objBrowsersList.count

Set objUserID = Description.Create()
objUserID("micclass").value = "WebEdit"
objUserID("html tag").value = "INPUT"
objUserID("name").value = "userName"

set objWebEditList=Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").ChildObjects(objUserID)
If objWebEditList.count>0 Then
	objWebEditList(0).set "mercury"
End If 

Set objPassword = Description.Create()
objPassword("micclass").value = "WebEdit"
objPassword("html tag").value = "INPUT"
objPassword("name").value = "password"

set objWebEditList=Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").ChildObjects(objPassword)
If objWebEditList.count>0 Then
	objWebEditList(0).set "mercury"
End If 

Set objSignIn = Description.Create()
objSignIn("micclass").value = "Image"
objSignIn("html tag").value = "INPUT"
objSignIn("alt").value = "Sign-In"

set objImageList=Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").ChildObjects(objSignIn)
If objImageList.count>0 Then
	objImageList(0).click
End If


'
'
'
'' Use the same to script it
'Browser("Math Calc").Page("Num Calculator").WebButton(btncalc).Click
'
'Dim oDesc
'Set oDesc = Description.Create
'oDesc("micclass").value = "Link"
'
''Find all the Links
'Set obj = Browser("Math Calc").Page("Math Calc").ChildObjects(oDesc)
'
'Dim i
''obj.Count value has the number of links in the page
'
'For i = 0 to obj.Count - 1	 
'   'get the name of all the links in the page			
'   x = obj(i).GetROProperty("innerhtml") 
'   print x 
'Next
'
'' Using Location
'Dim Obj
'Set Obj = Browser("title:=.*google.*").Page("micclass:=Page")
'Obj.WebEdit("name:=Test","location:=0").Set "ABC"
'Obj.WebEdit("name:=Test","location:=1").Set "123"
' 
'' Index
'Obj.WebEdit("name:=Test","index:=0").Set "1123"
'Obj.WebEdit("name:=Test","index:=1").Set "2222"
' 
'' Creation Time
'Browser("creationtime:=0").Sync
'Browser("creationtime:=1").Sync
'Browser("creationtime:=2").Sync
