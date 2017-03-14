
<!--#include file="cores.asp"-->
<%
dim varDe,varPara,varAssunto,varMensagem
varDe = request.form("txtDe")
varPara = request.form("txtPara")
varAssunto = request.form("txtAssunto")
varMensagem = request.form("txtMensagem")

%>

<html>
<body bgcolor="<%=escuro%>">
<%
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName = "Imobiliária Veja"
Mailer.FromAddress= "veja@imobiliariaveja.com.br"
Mailer.RemoteHost = "smtp2.locaweb.com.br"
Mailer.AddRecipient "Imobiliária Veja", varPara
Mailer.Subject = varAssunto
Mailer.BodyText = varMensagem
If Mailer.SendMail Then
Response.Write "<br><br><center><strong><font color='#FFFFFF' size='2' face='Verdana, Arial, Helvetica, sans-serif'>Mensagem enviada!!</font> </strong></center>"
Else
Response.Write "Erro " & Mailer.Response
End If

Set Mailer = Nothing
%>
</body>
</html>