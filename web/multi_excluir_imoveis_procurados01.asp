

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCod_imoveis_procurados,vcheck01,SearchFor



SearchFor = request.Querystring("SearchFor")
 varCod_imoveis_procurados = request.form("varCod_imoveis_procurados")
 vcheck01 = request.form("check01") 
 dim page
 
 page = request.querystring("page")                                                          
	
	if varCod_imoveis_procurados = "" and vcheck01 = "" then
	response.Redirect "archive_imoveis_procurados.asp?SearchFor="&SearchFor&""
	end if														  
	
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	
	
	
	
	Conexao.Open dsn
	
	
	

	 
	
		   
    
    
   
	
	
	
	
	
	
	
	Conexao.execute"delete from imoveis_procurados where cod_procurados in ("& request.form("check01") &")"
	 
	 
 
	 
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	
	  response.Redirect "archive_imoveis_procurados.asp?page="&cInt(Page)&"&SearchFor="&SearchFor&""
     
	  
	  
   'response.write page
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
 
     
 
 
           
           
           
           %>
<% response.flush%>
<%response.clear%>
</body>
</html>




