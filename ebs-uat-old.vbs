'1)

dim url, user, pswd
url="http://ebs.szgdjt.com:8020/OA_HTML/RF.jsp?function_id=28910&resp_id=-1&resp_appl_id=-1&security_group_id=0&lang_code=ZHS&params=m-gisey8PDGZZYbzm2wZ.gLbJr1taCPfyOCZRReE.6Y&oas=PXMNTSeFZEflkG4iSI5o-w.."
user="liu"
pswd="xxx"

DIM oIE
Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate url
oIE.Visible = True

n = 0
Do while oIE.Busy
  n = n + 1
  WSH.Sleep 200
Loop
if n = 301 then wsh.echo "ERROR" : wsh.quit 1

'2)

Set UID = oIE.document.all.usernameField
UID.value = user

Set PWD = oIE.document.all.passwordField
PWD.value = pswd

'try clicking Chinese 
'oIE.document.all.langOptionsTable.click
oIE.document.all.SubmitButton.click

'3)
'sLocation = "*ERROR*"
'n = 0
'Do until oIE.document.ReadyState = "complete" or n=100
'  n=n+1
'  WSH.Sleep 50
'Loop

'sLocation = lcase(unescape(oIE.document.location))

'set oLogfile = CreateObject("scripting.filesystemobject").opentextfile("D:\tech\source\web\auto-login-ebs\logfile.txt", 8, true)
'if not sLocation = "https://xyz.com/the/expected/destination_page.ext" then
'  oLogfile.writeline "Login failed"
'else

'4)
'   oLogfile.writeline "You made it to " & sLocation
'end if
