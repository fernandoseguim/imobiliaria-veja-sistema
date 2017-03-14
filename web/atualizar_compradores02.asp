
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
 

 
 vProprietario=request.form("txt_nome")
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
	
	
	if  vCidade <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 from combo1 where id_combo1 ="&vCidade
 
 rs2.open SQL2,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 
 rs2.close
 
 set rs2 = nothing
 
 
 else
 vCidade2 = "não informado"
 end if
 
 
'------------------------------pegar os vários bairros-----------------------
 
 
 
 dim vBairro2
 
 
 if vBairro <> "bqualquer" then
 
 
 dim rsMultiBairros
 dim sqlMultiBairros
 
 Set rsMultiBairros = Server.CreateObject("ADODB.RecordSet")
 

dim Variavel
dim Retorno
dim i
Variavel = vBairro
Retorno = Split(Variavel,", ")

i=0


for i=0 to UBound(Retorno)



sqlMultiBairros = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 from combo2 where id_combo2 ="& Retorno(i)


rsMultiBairros.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMultiBairros.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMultiBairros.ActiveConnection = Conexao




rsMultiBairros.open sqlMultiBairros,Conexao,2,1



 while not rsMultiBairros.eof

vBairro2 = rsMultiBairros("nome_combo2")&", "&vBairro2 

rsMultiBairros.MoveNext 
Wend



rsMultiBairros.close








next


set rsMultiBairros = nothing

 else
 vBairro2 = "não informado"
 end if
 
	
	
	 
	
	
	
	
	
	
	
	dim varCodimovel2
	varCodimovel2 = "não informado"
	
	dim vAtendimento
	vAtendimento = request.Form("txt_atendimento")
	
	if vAtendimento = "" then
	vAtendimento = "não informado"
	end if
	
	
	
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
	 
	 
	 
	 Conexao.execute"update compradores set nome='"&vProprietario&"',telefone='"&vTelefone&"',email='"&vEmail&"',cidade='"&vCidade2&"',bairro='"&vbairro2&"',tipo='"&vTipo&"',quartos='"&vQuartos&"',valor='"&int(vValor)&"',descricao='"&vDescricao&"',negociacao='"&vNegociacao&"',atendimento='"&vAtendimento&"',data_atualizacao='"&vData2&"',vagas='"&vVagas &"',ocupacao='"&vOcupacao &"' where cod_compradores="&varCodCompradores
	 
	
	conexao.close
	
	set conexao = nothing
	
	
	  response.Redirect "conta_comprador01.asp?varSucesso_imovel="&vProprietario&"&varCodCompradores="&varCodCompradores&""
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel incluído</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
 
     
 rs7.Close
           
           Set rs7 = Nothing
 
           
           Conexao.Close
           
           %>
<% response.flush%>
<%response.clear%>
</body>
</html>

