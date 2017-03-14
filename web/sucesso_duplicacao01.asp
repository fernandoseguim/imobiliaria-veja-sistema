<!--#include file="cores.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sucesso na duplicação</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>">
<%
dim varSucesso_duplicacao

varSucesso_duplicacao = request.querystring("varSucesso_duplicacao")

%>
<br>
<br>
<br>
<center>
<strong><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_duplicacao%> foi duplicado 
com sucesso.</font></strong> </center>
</body>
</html>
