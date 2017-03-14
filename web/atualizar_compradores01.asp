
<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodimovel,rsFoto,vFoto,vNome,strSQL2,vProposta,vimagem,vdata,vHorario,vInteresse
 
 
 dim vCidade,vBairro,vTipo,vNegociacao,vValor
 Dim vProprietario,vTelefone,vEmail,vEndereco
 dim vQuartos,vDescricao
 dim vVila,vVila2

 dim varCodCompradores
 
 varCodCompradores = request.Querystring("varCodCompradores")
 

 
 vProprietario=request.form("txt_proprietario")
 vTelefone=request.form("txt_telefone")
 vEmail=request.form("txt_email")

 vCidade = request.form("combo1")
 vBairro = request.form("combo2")
  vVila=request.form("combo5")
 vNegociacao = request.Form("example2")
 vTipo=request.form("txt_tipo")
 vValor=request.form("stage22")
 vQuartos=request.Form("txt_quartos")
 
 
  if vQuartos = "não informado" then
 vQuartos = "0"
 end if
 
 
 vDescricao=request.Form("txt_descricao")
 

 if vDescricao = "" then
 vDescricao = "não informado"
 end if
 
 
 
 

 

 
 

 
 
	
	 
  
  
   Dim vdata2

vdata2 = now()


  if vNegociacao = "nqualquer" then
  vNegociacao = "qualquer um"
  end if
  
 
	  
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	if vCidade <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 else
 vCidade2 = "não informado"
 end if
 
 
 
 if vBairro <> "bqualquer" then
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2
 vBairro2 = rs3("nome_combo2")
 else
 vBairro2 = "não informado"
 end if
 
	
	
	 if vVila <> "vlqualquer" then
 
 dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1

 vVila2 = rs333("nome_combo3")
 
 else
 
 vVila2 = "não informado"
 
 end if
	
	
	
	
	
	
	
	dim varCodimovel2
	varCodimovel2 = "não informado"
	
	
	
	
	
	dim vVagas
	vVagas = request.Form("txt_vagas")
	
	if vVagas = "não informado" then
 vVagas = "0"
 end if
	
	dim vOcupacao
	vOcupacao = request.Form("txt_ocupacao")
	
	if vOcupacao = "" then
	vOcupacao = "não informado"
	end if
	
	dim vAtendimento
	vAtendimento = request.Form("txt_atendimento")
	
	if vAtendimento = "" then
	vAtendimento = "não informado"
	end if
	 
	 
	 
	 Conexao.execute"update compradores set nome='"&vProprietario&"',telefone='"&vTelefone&"',email='"&vEmail&"',cidade='"&vCidade2&"',bairro='"&vbairro2&"',tipo='"&vTipo&"',quartos='"&vQuartos&"',valor='"&vValor&"',descricao='"&vDescricao&"',negociacao='"&vNegociacao&"',data_atualizacao='"&vData2&"',vila='"&vVila2 &"',vagas='"&vVagas &"',ocupacao='"&vOcupacao &"',atendimento='"&vAtendimento &"' where cod_compradores="&varCodCompradores
	 
	
	  response.Redirect "visualizar_compradores02.asp?varSucesso_imovel="&vProprietario&"&varCodCompradores="&varCodCompradores&""
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel incluído</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





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
<td width="217" height="156" ></td>    <td width="202" height="156" ><img src="sorriso_proposta.jpg" border="0"></img></td>   <td width="217" height="156" ></td>
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
 
     
 rs7.Close
           
           Set rs7 = Nothing
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>

