

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodProposta,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,vcheck01,SearchFor


dim page
  page = request.querystring("page")
dim varIDProposta_de_fora

varIDProposta_de_fora = request.QueryString("varIDProposta_de_fora")



SearchFor = request.Querystring("SearchFor")
 varCodProposta = request.form("varCodProposta")
 vcheck01 = request.form("check01")
 
 if varIDProposta_de_fora = "" and vcheck01 = "" then
 response.Redirect "archive_proposta_de_fora.asp?SearchFor="&SearchFor&""
 end if
 
                                                              
															  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	
	
	Conexao.execute"delete from proposta_de_fora where id_proposta_de_fora in ("& request.form("check01") &")"
	 
	 
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	 
	 
	 response.redirect "archive_proposta_de_fora.asp?SearchFor="&SearchFor&"&page="&cInt(Page)&""
	
	  
	  
   
   
   
   
   
  
   
   
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
 
     
 
 
           
         
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


