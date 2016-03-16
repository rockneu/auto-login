' learned from  http://www.anwarsayed.com/3/post/2011/07/vbs-script-login-to-website-automatically.html

'1)

dim url, user, pswd
url="http://maximo.sz-mtr.com/maximo/webclient/login/login.jsp?uisessionid=1366092630829"
user="zhang"
pswd="xxx"

url="http://maximo.sz-mtr.com/maximo/webclient/login/login.jsp?uisessionid=139873913878&allowinsubframe=false&username=zhangtingting&password=7788520"

DIM oIE
Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate url
oIE.Visible = True

