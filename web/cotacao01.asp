<%
Set fso = CreateObject("Scripting.FileSystemObject")
caminho = Server.MapPath("teste02.txt")
Set SITE = fso.CreateTextFile(caminho, True)
url="http://ultimosegundo.ig.com.br/economia/painel/bvsp.txt"
Set objHTTP = CreateObject("MSXML2.XMLHTTP")
Call objHTTP.Open("GET", url, FALSE)
objHTTP.Send
SITE.WriteLine (objHTTP.ResponseText)                             
SITE.Close


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(caminho, 1)


Do Until objFile.AtEndOfStream

 strLine = objFile.Readline
 StrServiceList = Split(strLine , ";" , -1, 1)
 'Codigo  = StrServiceList(1)    
 'Nome  = StrServiceList(2)
' Valor  = StrServiceList(3)
 'Percentagem = StrServiceList(4)

 'If Codigo = "PETR4.SA" or Codigo = "VALE5.SA" Then

 ' Corpo = Corpo & "Código: " & StrServiceList(1) & vbcrlf    
 ' Corpo = Corpo & "Nome:   " & StrServiceList(2) & vbcrlf
 ' Corpo = Corpo & "Valor:     " & StrServiceList(3) & vbcrlf
 ' Corpo = Corpo & "%          " & StrServiceList(4) & vbcrlf
 ' Corpo = Corpo & "----------------------------------" & vbcrlf
 
' End If

Loop


'response.write Corpo
response.write StrServiceList

%>