<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Documento sem t&iacute;tulo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<!--#include file="dsn.asp"-->
<%

dim rs,SQL,Conexao

	  Set Conexao = Server.CreateObject("ADODB.Connection")
	  Set rs = Server.CreateObject("ADODB.RecordSet")
	  
	  Conexao.open dsn





dim Variavel
dim Retorno
dim i
Variavel = "Nova Petrópolis, Jordanópolis, Rudge Ramos"
Retorno = Split(Variavel,", ")

i=0


for i=0 to UBound(Retorno)
%>

<%
SQL = "select * from imoveis where bairro like '"& Retorno(i) &"' ORDER BY Cod_imovel   DESC"




rs.open SQL,Conexao,2,1

while not rs.eof


response.write rs("bairro")&"<br>"

rs.MoveNext
Wend

rs.close




%>


<%

next

%>

</body>
</html>
