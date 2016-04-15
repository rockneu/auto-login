Dim objIE, IEObject, Info, all, hasOwnProperty
Info = ""

Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.Navigate "http://www.baidu.com"

Do While objIE.Busy Or (objIE.READYSTATE <> 4)
 Wscript.Sleep 100
Loop
WScript.Sleep 50

objIE.Document.getElementById("kw").value="test kw"
objIE.Document.getElementById("su").click

'Set hasOwnProperty = objIE.Document.ParentWindow.Object.prototype.hasOwnProperty

' Can I use objIE.Document.all.Item.length to count all elements on the webpage?
Set all = objIE.Document.all
For i = 0 To all.Item.Length - 1
Set IEObject = all.Item(i)

'If this is not the first item, place this data on a new line.
If i > 0 Then Info = Info & vbCrLf

' Number each line.
Info = Info & i + 1 & ". "

' Specify the ID if it is given.
If hasOwnProperty.call(IEObject, "id") Then
  Info = Info & IEObject.id
Else
  Info = Info & "[NO ID]"
End If

' Specify the title if it is given.
If hasOwnProperty.call(IEObject, "title") Then
  Info = Info & "-" & IEObject.title
Else
  Info = Info & "-[NO TITLE]"
End If

' Specify the name if it is given.
If hasOwnProperty.call(IEObject, "name") Then
  Info = Info & "-" & IEObject.name
Else
  Info = Info & "-[NO NAME]"
End If
Next

Wscript.Echo Info

Set IEObject = Nothing
Set objIE = Nothing