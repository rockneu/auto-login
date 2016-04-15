' This script is designed to auto login some sites
'
'
'
'			author: Rock Liu 2016.03.18
'			
'

'1)

dim url, user, pswd
user="liuwei"
pswd="1234"
url="http://10.10.10.57:8011/"

DIM ie ', oIEres
Set ie = CreateObject("InternetExplorer.Application")
ie.navigate url
ie.Visible = True


'Do While ie.readyState<>4 Or ie.busy	'page loaded completely
Do While ie.readyState<>4 'Or Mid(ie.locationURL,8,3)<>"sso"	'page loaded completely
	wscript.sleep 300
Loop  

'wscript.echo Mid(ie.LocationURL,8,3)
'MsgBox(ie.Document.body.innerHTML)
'wscript.sleep 3000



Dim all, obj
Set all=ie.Document.all
'MsgBox all.length
For i=0 To all.length-1
	'If all.item(i).id="txtUserName" Then
	If all(i).Id="txtUserName" Then
		Set obj=all.item(i)
		'MsgBox all(i).width
		'all(i).value=user
		
	End If

	If all(i).Id="txtPassword" Then
		Set obj=all.item(i)
		'all(i).value=pswd
		Exit For
	End If
Next 

Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file=fso.CreateTextFile("D:\pwork\source.code\auto-login\cs.htm.txt", True)
file.WriteLine(ie.Document.body.innerHTML)
file.close()
Set file=Nothing
Set fso=Nothing


'wscript.echo ie.locationURL & "----" & ie.Document.body.innerHTML


'MsgBox(ie.Document.body.innerHTML)
'wscript.echo ie.Document.getElementById("txtUserName").class
'wscript.echo ie.Document.getElementByName("txtUserName").width


Set ie.document.getElementById("tbUserName").value=user
Set ie.Document.getElementById("tbPassword").value=pswd




ie.Document.getElementById("btnLogin").click
