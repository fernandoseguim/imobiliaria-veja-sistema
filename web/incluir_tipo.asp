<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin02.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vTipo,vBairro
Dim vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vVagas,vValor,vNegociacao,vFoto
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
 
 
 
  
	
    vTipo=request.form("txt_tipo")  
   
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
 
	                           
	dim rs4,SQL4
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select * from tipo where tipo like '"& vTipo&"'"
 
 rs4.open SQL4,Conexao,2,1
 
 if not rs4.eof then
 dim existente
 existente = " já existe no banco de dados"
 response.Redirect "form_incluir_tipo.asp?varExistente="&vTipo&existente&""
 
 end if						   
							              
															  
	
   
	
	 
	
	Conexao.execute"Insert into tipo(tipo,data_tipo) values( '"& vTipo &"','"& now() &"')"
	 
	 response.Redirect "form_incluir_tipo.asp?varSucesso_tipo="&vTipo&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
          <td width="202" height="156"><img src="sorriso_sugestao.jpg" width="202" height="156" border="0"></img></td>   
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
