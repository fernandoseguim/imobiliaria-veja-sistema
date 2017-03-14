

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,varCod_imovel,SearchFor,objFSO,FotoPeqEx,vProprietario,vEndereco 
dim rs3
dim strSQL3

dim varCodCompradoresProcurados
varCodCompradoresProcurados = request.QueryString("varCodCompradoresProcurados")
 SearchFor = request.querystring("SearchFor")
 
  
  
   dim varCidade, varBairro,varNegociacao,varQuartos,page,SearchWhere,varValor
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
	
	
	
	
	
	
	                                                        
															  
	
	Set rs3 = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	

	
	Conexao.Open dsn
	
	

RS3.CursorLocation = 3
RS3.CursorType = 3

        
		
		
		
		
		
		
		
		
		
		
		
	
	
	Conexao.execute"Delete from compradores_procurados where cod_compradoresprocurados="&varCodCompradoresProcurados
	 
	 
	
	 conexao.close
	 
	 
	 Set rs3 = nothing
    Set Conexao = nothing
	Set objFSO = nothing
	
	 
	 
	 
	 response.redirect "archive_compradores_procurados.asp?page="&cInt(Page)&"&SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varCidade="&varCidade&"&varBairro="&varBairro&"&varNegociacao="&varNegociacao&"&varQuartos="&varQuartos&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varValor="&varValor&""
	
	

	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">










 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>



