' learned from  http://www.anwarsayed.com/3/post/2011/07/vbs-script-login-to-website-automatically.html

'1)

dim url, user, pswd
url="http://maximo.sz-mtr.com/maximo/webclient/login/login.jsp?uisessionid=1366092630829"
user="zhangtingting"
pswd="7788520"

'http://maximo.sz-mtr.com/maximo/webclient/login/login.jsp?uisessionid=139873913878&allowinsubframe=false&username=zhangtingting&password=7788520

DIM oIE
Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate url
oIE.Visible = True

WSH.Sleep 500

n = 0
'Do while oIE.Busy or n = 101
Do while oIE.Busy
  n = n + 1
  WSH.Sleep 100
Loop
if n = 201 then wsh.echo "ERROR" : wsh.quit 1

'2)

'Set UID = oIE.document.all.username
Set UID = oIE.document.GetElementById("username")
UID.value = user

'Set PWD = oIE.document.all.password
Set PWD = oIE.document.GetElementById("password")
PWD.value = pswd

'try clicking Chinese 
'oIE.document.all.langOptionsTable.click
'oIE.document.all.loginbutton.click
'Set BTN = oIE.document.all.loginbutton
Set BTN = oIE.document.GetElementById("loginbutton")
BTN.click

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
