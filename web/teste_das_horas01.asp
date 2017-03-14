<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<%
dim hora, dia, mes, ano

hora = hour(now())
dia = day(now())
mes = month(now())
ano = year(now())

response.write hora&"/"&dia&"/"&mes&"/"&ano

%>
</body>
</html>
