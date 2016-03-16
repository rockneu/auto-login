' learned from  http://www.anwarsayed.com/3/post/2011/07/vbs-script-login-to-website-automatically.html

'1)

dim url, user, pswd
user="liu"
pswd="xxx"
url="http://ebs.szgdjt.com:8000/OA_HTML/fndvald.jsp?username="+user+"&password="+pswd+"&langCode=ZHS"
'url="http://ebs.szgdjt.com:8020/OA_HTML/RF.jsp?function_id=28910&resp_id=-1&resp_appl_id=-1&security_group_id=0&lang_code=ZHS&params=DA0YNVtselNaA5IFP3G0zw&oas=47-Zwvp4Pawyo1xIQml-fQ.."
DIM oIE
Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate url
oIE.Visible = True
