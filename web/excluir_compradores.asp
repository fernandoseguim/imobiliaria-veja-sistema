

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,varCod_imovel,SearchFor,objFSO,FotoPeqEx,vProprietario,vEndereco 
dim rs3
dim strSQL3

dim varCodCompradores
 varCodCompradores = request.QueryString("varCodCompradores")
 SearchFor = request.querystring("SearchFor")
 
  
  
   dim varCidade, varBairro,varNegociacao,varQuartos,page,SearchWhere,varValor
dim varValor1,varValor2
 
 
 
 dim varCidade2
dim varBairro2
dim varTipo

dim varSuites 
dim varPiscina 
dim varPortaria 
dim varQuintal 
dim varQuadras 
dim varEdicula 
dim varCondominio
dim varCondominio1
dim varCondominio2
dim varAreaTotal 
dim varAreaTotal1 
dim varAreaTotal2 
dim varOcupacao 
dim varAtendimento
dim varCondicoes 
dim varStandby
dim varVagas
 
 
 
 SearchWhere = request.querystring("SearchWhere")
 varCidade = request.querystring("varCidade")
  varBairro = request.querystring("varBairro")  
  varNegociacao = request.querystring("varNegociacao")
  varQuartos = request.querystring("varQuartos")
  varValor1 = request.querystring("varValor1")
  varValor2 = request.querystring("varValor2")
  varValor = request.querystring("varValor")
   page = request.querystring("page")        
	
	
varCidade2 = request.querystring("varCidade2")
 varBairro2 = request.querystring("varBairro2")
 varTipo = request.querystring("varTipo")

 varSuites  = request.querystring("varSuites")
 varPiscina = request.querystring("varPiscina") 
 varPortaria = request.querystring("varPortaria") 
 varQuintal = request.querystring("varQuintal") 
 varQuadras = request.querystring("varQuadras") 
 varEdicula = request.querystring("varEdicula") 
 varCondominio = request.querystring("varCondominio")
 varCondominio1 = request.querystring("varCondominio1")
 varCondominio2 = request.querystring("varCondominio2")
 varAreaTotal = request.querystring("varAreaTotal") 
 varAreaTotal1 = request.querystring("varAreaTotal1") 
 varAreaTotal2  = request.querystring("varAreaTotal2")
 varOcupacao  = request.querystring("varOcupacao")
 varAtendimento = request.querystring("varAtendimento")
 varCondicoes = request.querystring("varCondicoes") 
 varStandby = request.querystring("varStandby")
 varVagas = request.querystring("varVagas")
	
	
	
	                                                        
															  
	
	Set rs3 = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	

	
	Conexao.Open dsn
	
	

RS3.CursorLocation = 3
RS3.CursorType = 3

        
		
		
		
		
		
		
		
		
		
		
		
	
	
	Conexao.execute"Delete from compradores where cod_compradores="&VarCodCompradores
	 
	 
	
	 conexao.close
	 
	 
	 Set rs3 = nothing
    Set Conexao = nothing
	Set objFSO = nothing
	
	 
	 
	 
	' response.redirect "archive_compradores.asp?SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varCidade="&varCidade&"&varBairro="&varBairro&"&varNegociacao="&varNegociacao&"&varQuartos="&varQuartos&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varValor="&varValor&"&page="&cInt(Page)&""
	
	response.redirect "archive_compradores.asp?page="&cInt(Page)&"&varCidade="&varCidade&"&varCidade2="&varCidade2&"&varBairro="&varBairro&"&varBairro2="&varBairro2&"&varNegociacao="&varNegociacao&"&varTipo="&varTipo&"&varQuartos="&varQuartos&"&varVagas="&varVagas&"&SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varValor="&varValor&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varSuites="&varSuites&"&varPiscina="&varPiscina&"&varPortaria="&varPortaria&"&varQuintal="&varQuintal&"&varQuadras="&varQuadras&"&varEdicula="&varEdicula&"&varCondominio="&varCondominio&"&varCondominio1="&varCondominio1&"&varCondominio2="&varCondominio2&"&varAreaTotal="&varAreaTotal&"&varAreaTotal1="&varAreaTotal1&"&varAreaTotal2="&varAreaTotal2&"&varOcupacao="&varOcupacao&"&varStandby="&varStandby&"&varAtendimento="&varAtendimento&""


	  
	  
   
   
   
   
   
  
   
   
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


