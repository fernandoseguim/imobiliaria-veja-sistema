<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%


%>

<html>
<body bgcolor="<%=escuro%>">
<%
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")

Mailer.ContentType = "text/html"
Mailer.FromName = "Imobiliária Veja"
Mailer.FromAddress= "wsbraga@imobiliariaveja.com.br"
Mailer.RemoteHost = "smtp2.locaweb.com.br"
Mailer.AddRecipient "Imobiliária Veja", "wentznico@terra.com.br"
Mailer.Subject = "teste de email"
Mailer.Bodytext = "Caro <b>" & "Nico" & ",</b>" & "<br><a href='www.cbn.com.br' target='_blank'>vamos testar</a>"
'Mailer.BodyText = varMensagem
If Mailer.SendMail Then
Response.Write "<br><br><center><strong><font color='#FFFFFF' size='2' face='Verdana, Arial, Helvetica, sans-serif'>Mensagem enviada!!</font> </strong></center>"
Else
Response.Write "Erro " & Mailer.Response
End If

Set Mailer = Nothing
%>
</body>
</html>

<%
conexao.close
set conexao = nothing
%>
<!--#include file="dsn2.asp"-->