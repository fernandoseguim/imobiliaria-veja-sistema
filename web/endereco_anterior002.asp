<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<%
dim endereco01,endereco02

endereco01 = request.serverVariables("remote_addr")
endereco02 = request.serverVariables("http_referer")

response.write endereco01&"||"&endereco02
%>
<a href="endereco_anterior001.asp">Teste</a>
</body>
</html>
