' learned from  http://www.anwarsayed.com/3/post/2011/07/vbs-script-login-to-website-automatically.html

'1)

dim url, user, pswd

user="liuwei"
pswd="rock000"
url="http://ebsuat.szgdjt.com:8000/OA_HTML/fndvald.jsp?username="+user+"&password="+pswd+"&langCode=ZHS"
urlres="http://ebsuat.szgdjt.com:8000/OA_HTML/RF.jsp?function_id=261&resp_id=50862&resp_appl_id=20003&security_group_id=0&lang_code=ZHS&oas=Vq_HgvEtAic5YzvnljNBiA.."


DIM oIE
Set oIE = CreateObject("InternetExplorer.Application")
oIE.navigate url
oIE.Visible = True

Do While oIE.readyState<>4 	'page loaded completely
	wscript.sleep 300
Loop  

oIE.navigate urlres
oIE.Visible = True

Do While oIE.readyState<>4 	'page loaded completely
	wscript.sleep 300
Loop
