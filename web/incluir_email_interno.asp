<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%


if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if


Dim Conexao,strSQL,rs,vdata,vNome,vEmail,vAssunto,vMensagem,vTelefone
 Dim vdata2

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
   
   dim txtPara
   dim txtDe
   dim txtAssunto
   dim txtMensagem
   
   
    txtPara=request.form("txtPara")
    txtDe=request.form("txtDe")
    txtAssunto=request.form("txtAssunto")
    txtMensagem=request.form("txtMensagem")
   
   
  
	
	
	
	
	                                                  
															  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	 
	
	
	 
	'--------------------Verificar se já tem cadastro-------------- 
	 

	 
	 '--------------------------------------------------------------
	 
	  Conexao.execute"Insert into email_interno(quem_mandou, quem_recebeu, foi_visto ,assunto,mensagem,telefone_quem_mandou,data,origem_franquia,cod_quem_mandou,cod_quem_recebeu) values( '"& txtDe &"','"& txtPara &"','"& "não" &"','"& txtAssunto &"','"& txtMensagem &"','"& session("telefone_interno") &"','"& now() &"','"& session("vOrigem_Franquia") &"','"& "não informado" &"','"& "não informado" &"')" 
	dim varSucesso
	
	response.Redirect "form_enviar_email_interno.asp?varSucesso="&"Mensagem enviada com sucesso"&""
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sugestão incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"  ><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
<td width="590" height="105" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    
          <td width="202" height="156"><img src="sorriso_email.jpg" width="202" height="156" border="0"></img></td>   
          <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36"></img></td>

</tr>


</table>







 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
