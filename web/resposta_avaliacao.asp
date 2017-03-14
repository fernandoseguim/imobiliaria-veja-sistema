<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<%

dim resposta

resposta = request.form("txt_pergunta")



if resposta = "" then
response.redirect "default.asp"

else


end if


if resposta <> "nao" then
response.redirect "procurar_avaliacao02.asp"

else

response.redirect "primeira01.asp"
end if





%>
</body>
</html>
