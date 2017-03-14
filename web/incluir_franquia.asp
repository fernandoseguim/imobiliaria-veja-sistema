<%
Option Explicit
%>
<!--#include file="dsn.asp"-->


<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vVagas,vValor,vNegociacao,vFoto
Dim vdata2
dim vVila

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
 
 
 
  
	
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
    
	
   
	dim vFranquia
	dim vEndereco
	dim vTelefone
	dim vEmail
	
	
	vFranquia = request.form("txt_franquia")
	vEndereco = request.form("txt_endereco") 
	vTelefone = request.form("txt_telefone")
	vEmail = request.form("txt_email") 
	 
	
	Conexao.execute"Insert into franquia(nome_franquia,data_franquia,endereco,telefone,email) values('"& vFranquia &"','"& vdata &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"')"
	 
	 
	 response.Redirect "form_incluir_franquia.asp?varSucesso_franquia="&vFranquia&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
          <td width="202" height="156"></img></td>   
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
