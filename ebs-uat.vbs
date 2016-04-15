' learned from  http://www.anwarsayed.com/3/post/2011/07/vbs-script-login-to-website-automatically.html

'1)

dim url, user, pswd

user="bGl1d2Vp"
pswd="cGFzczAwMA=="
url="http://ebsuat.szgdjt.com:8000/OA_HTML/fndvald.jsp?username="+Base64Decode(user)+"&password="+Base64Decode(pswd)+"&langCode=ZHS"
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



Function Base64Decode(ByVal base64String)
  
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function