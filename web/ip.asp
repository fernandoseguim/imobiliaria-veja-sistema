<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>


<body>
<%
dim EnderecoIP

EnderecoIP = request.ServerVariables("REMOTE_ADDR")

%>
<center><font size="4" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=EnderecoIP%></font></center>
</body>
</html>
