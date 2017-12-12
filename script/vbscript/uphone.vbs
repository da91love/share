Dim IE 

'On Error Resume Next

Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = 1  
IE.navigate "http://www.uphone.co.kr/"

While IE.Busy
  WScript.Sleep 1000
Wend


If Instr(IE.Document.title,"Study different") Then
'	Wscript.Echo(IE.Document.title)
	
	IE.Document.getElementByID("ID").value = "da91love@naver.com"

	IE.Document.getElementByID("PWD").value = "rlaektmf21"

	IE.Document.getElementsByTagName("fieldset")(1).getElementsByTagName("button")(0).click()
	
Else
	Wscript.Echo("page can not be found")
End If

'If Err.Number = 0 
'Else
'	
'End IF

'IE.Quit