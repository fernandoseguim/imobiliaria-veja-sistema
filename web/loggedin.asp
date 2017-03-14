
<%
Session.TimeOut = 480
	If Session("isLoggedIn") <> True Then
		Response.Redirect "senha.asp"
	End If
	dim varEnderecoIP,SQLIP,rsIP
	'varEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	'varEnderecoIP = "123456"
	
	
	
	Set rsIP = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQLIP = "Select * from ip where ip ='"&varEnderecoIP&"' ORDER BY id_ip DESC"

rsIP.Open SQLIP, Conexao
 
 if rsIP.eof then
' Response.Redirect "verificacao_ip.asp"
 
 end if
 

%>