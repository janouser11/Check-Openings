Dim URL 
Dim IE 
Set IE = CreateObject("internetexplorer.application")
URL = "https://my.vcu.edu/myvcu-portlets/ssologin.do?applicationId=BANNER_SELF_SERVICE" 
IE.Visible = True
IE.Navigate URL

Call Sleep

With IE.Document
.getElementByID("username").value = "username"
.getElementByID("password").value = "password"
IE.document.Forms(0).submit
End With


Call Sleep

Set a = ie.Document.GetElementsByTagName("a")

' Find the link titled "Student" and click it...
For i = 0 To a.Length - 1
If a(i).innerText = "Student" Then a(i).Click
Next

Call Sleep

Set a = ie.Document.GetElementsByTagName("a")

' Find the link titled "Student" and click it...
For i = 0 To a.Length - 1
If a(i).innerText = "Registration" Then a(i).Click
Next

Call Sleep

Set a = ie.Document.GetElementsByTagName("a")

For i = 0 To a.Length - 1
If a(i).innerText = "Add or Drop Classes" Then a(i).Click
Next

Call Sleep

For Each btn In IE.Document.getElementsByTagName("input")
If btn.type = "submit" Then btn.Click()
Next

Call Sleep

With IE.Document
'.getElementByID("crn_id1").value = "20866"
'   .getElementByID("crn_id2").value = "31522"
'   .getElementByID("crn_id3").value = "31528"
'.getElementByID("crn_id4").value = "17997"
'   .getElementByID("crn_id5").value = "29459"
'.getElementByID("crn_id6").value = "24537"
'.getElementByID("crn_id6").value = "13368"
End With

Call Sleep

For Each btn In IE.Document.getElementsByTagName("input")
If btn.Value = "Submit Changes" Then btn.Click()
Next

Call Sleep

test="Closed Section"
value = true

Do Until value = false
	'Wscript.echo "In do loop"
		If instr (ie.document.body.innerText,test) then
			'Wscript.echo "I have found " & test 
				With IE.Document
					.getElementByID("crn_id1").value = "13368"
				End With
				For Each btn In IE.Document.getElementsByTagName("input")
					If btn.Value = "Submit Changes" Then btn.Click()
				Next
			Call Sleep
			WScript.Sleep 60000
		else
			'Wscript.echo "Loop is now ending"	
			Call email
			value = false
		end if
Loop 

 ' Wscript.echo "Loop ended"

Call Sleep

Wscript.echo "The selected class is open and an email has been sent!"	

Function email()
	
Set emailObj      = CreateObject("CDO.Message")

emailObj.From     = "email"
emailObj.To       = "email"

emailObj.Subject  = "Info 350 Class Is open"
emailObj.TextBody = "CRN is 13368"
'If err.number = 0 then Msgbox "Email is sending"
Set emailConfig = emailObj.Configuration

emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")      = true 
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")    = "username"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")    = "password"
emailConfig.Fields.Update
emailObj.Send
'If err.number = 0 then Msgbox "Email Sent"


End Function

Function sleep()
	Do While IE.Busy
	WScript.Sleep 1000
	Loop
End Function
