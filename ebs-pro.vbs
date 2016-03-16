' This script is designed to auto login Oracle EBS


' learned from  http://www.anwarsayed.com/3/post/2011/07/vbs-script-login-to-website-automatically.html

'1)

dim url, user, pswd
user="liuwei"
pswd="rock000"
url="http://ebs.szgdjt.com:8000/OA_HTML/fndvald.jsp?username="+user+"&password="+pswd+"&langCode=ZHS"
'url="http://ebs.szgdjt.com:8020/OA_HTML/RF.jsp?function_id=28910&resp_id=-1&resp_appl_id=-1&security_group_id=0&lang_code=ZHS&params=DA0YNVtselNaA5IFP3G0zw&oas=47-Zwvp4Pawyo1xIQml-fQ.."
urlres="http://ebs.szgdjt.com:8000/OA_HTML/RF.jsp?function_id=261&resp_id=50862&resp_appl_id=20003&security_group_id=0&lang_code=ZHS&oas=Vq_HgvEtAic5YzvnljNBiA.."

DIM oIE ', oIEres
Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate url
oIE.Visible = True

'Do While oIE.readyState<>4 Or oIE.busy	'page loaded completely
Do While oIE.readyState<>4 	'page loaded completely
	wscript.sleep 300
Loop  

oIE.navigate urlres
oIE.Visible = True

Do While oIE.readyState<>4 	'page loaded completely
	wscript.sleep 300
Loop


'wscript.sleep 3000
'Dim treeItem
'Set treeItem=oIE.Document.getElementById("SZMTR_客户化职责")
'MsgBox "initialize treeItem"
'treeItem.click
'MsgBox "initialize treeItem2"

'Dim obj, btnOrg
'Set obj = oIE.document
'For i=0 To obj.all.length-1
'	If obj.all(i).tagname = "SZMTR_客户化职责" Then
'		MsgBox "found SZMTR"
'		obj.all(i).click
'	End if
'next	

'MsgBox "SZMTR_客户化职责 to click"
'oIE.Document.getElementById("50862:20003:-1:0").click
'MsgBox "SZMTR_客户化职责 clicked"


'set btnOrg=oIE.Document.getElementById("更改组织")
'btnOrg.click

'msgbox "xx"
