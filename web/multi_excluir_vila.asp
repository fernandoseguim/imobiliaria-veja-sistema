

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCod_bairro,vcheck01,SearchFor,SearchWhere,SearchWhereFor,page

SearchWhereFor = request.querystring("SearchWhereFor")
SearchWhere = request.querystring("SearchWhere")
SearchFor= request.querystring("SearchFor")
page = request.querystring("page")

 varCod_bairro = request.form("varCod_bairro")
 vcheck01 = request.form("check01")                                                           
	
	if varCod_bairro = "" and vcheck01 = "" then
	response.Redirect "archive_vila.asp?SearchWhere="&SearchWhere&"&SearchFor="&SearchFor&"&SearchWhereFor="&SearchWhereFor&"&page="&page&""
	end if														  
	
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	
	
	
	
	Conexao.Open dsn
	
	
	

	 
	
		   
    
    
   
	
	
	
	
	
	
	
	
	 
	 Conexao.execute"delete from combo3 where id_combo3 in ("& request.form("check01") &")"
	 
 
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	 
	
	  response.Redirect "archive_vila.asp?SearchWhere="&SearchWhere&"&SearchFor="&SearchFor&"&SearchWhereFor="&SearchWhereFor&"&page="&page&""
     
	  
	  
   
   
   
   
   
  
   
   
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


