

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodProposta,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,vcheck01,SearchFor

dim varCodClicado


dim SearchWhere

 dim varCidade, varBairro,varNegociacao,varQuartos,page,varValor
dim varValor1,varValor2
 
 
 SearchWhere = request.querystring("SearchWhere")
 varCidade = request.querystring("varCidade")
  varBairro = request.querystring("varBairro")  
  varNegociacao = request.querystring("varNegociacao")
  varQuartos = request.querystring("varQuartos")
  varValor1 = request.querystring("varValor1")
  varValor2 = request.querystring("varValor2")
  varValor = request.querystring("varValor")
   page = request.querystring("page")        
	
	




SearchWhere = request.Querystring("SearchWhere")
SearchFor = request.Querystring("SearchFor")
 varCodClicado = request.form("varCodClicado")
 session("varClicado") = varCodClicado
 
 vcheck01 = request.form("check01")
 session("vcheck01") = vcheck01
 
 if varCodClicado = "" and vcheck01 = "" then
 response.Redirect "archive_comprador_clicado_corretor.asp?searchFor="&searchFor&"&searchWhere="&searchWhere&""
 end if
 
                                                            
															  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	
	
	Conexao.execute"delete from comprador_clicado where cod_clicado in ("& request.form("check01") &")"
	 
	 conexao.close
	 set conexao = nothing
	 
	 
	 
	 response.redirect "archive_comprador_clicado_corretor.asp?page="&cInt(Page)&"&SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varCidade="&varCidade&"&varBairro="&varBairro&"&varNegociacao="&varNegociacao&"&varQuartos="&varQuartos&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varValor="&varValor&""
	 
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
    <td width="590" height="105" >&nbsp;</td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    <td width="202" height="156" ><img src="sorriso_proposta2.jpg"></img></td>   <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36" ></img></td>

</tr>


</table>







 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


