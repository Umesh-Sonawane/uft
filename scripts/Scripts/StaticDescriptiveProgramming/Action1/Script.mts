'Open Google.com
SystemUtil.Run "chrome.exe", "http://www.google.com"
 
Set objPage=Browser("title:=Google").Page("title:=Google","url:=https://www\.google\.com")
'Set value in Google Search Box
objPage.WebEdit("name:=q","type:=text").Set "automation repository"

objPage.WebEdit("html tag:=INPUT","name:=q").Set "Static Descriptive Programming"
'Click on Search button
objPage.WebButton("acc_name:=Google Search","html tag:=INPUT","type:=submit","index:=1").Click 
'Find out the number of search results displayed
sResults = Browser("opentitle:=Google","openurl:=https://www\.google\.com").Page("title:=.+Google Search").WebElement("html id:=result-stats","html tag:=DIV").GetROProperty("innertext")
 
'Display the Result in a message box
Msgbox sResults
'Open Google.com
SystemUtil.CloseProcessByName "chrome.exe"
