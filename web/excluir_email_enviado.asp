

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCod_email_enviado,SearchFor
 varCod_email_enviado = request.QueryString("varCod_email_enviado")
SearchFor = request.Querystring("searchFor")
 dim page
 page = request.Querystring("page")
                                                            
															  
	Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	
	
	Conexao.Open dsn
	
	

        
	
	
	Conexao.execute"delete from email_enviado where cod_email_enviado="&VarCod_email_enviado 
	
	
	
	set rs = nothing
	
	conexao.close
	
	set conexao = nothing
	 
	 response.redirect "archive_email_enviado.asp?searchFor="&SearchFor&"&page="&page&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
<td width="590" height="105" ></td>
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




