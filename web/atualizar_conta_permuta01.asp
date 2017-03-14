
<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodimovel,rsFoto,vFoto,vNome,strSQL2,vProposta,vimagem,vdata,vHorario,vInteresse
 
 Dim vCidade_vend,vBairro_vend,vTipo_vend,vDescricao_vend
 Dim vCidade_comp,vBairro_comp,vTipo_comp,vDescricao_comp,vLink
 Dim vProprietario,vTelefone,vEmail,vEndereco
 Dim vQuartos_vend,vQuartos_comp
 Dim vVila_vend,vVila_comp
 dim vVagas_vend,vVagas_comp

 Dim vValor_vend, vValor_comp
 dim vStandby
 
 vStandby = request.form("txt_standby")
 
 dim varCodPermuta
 
 varCodPermuta = request.Querystring("varCodPermuta")
 
 varCodImovel=request.form("txt_cod_imovel")
 vLink=request.form("txt_link")
 vProprietario=request.form("txt_proprietario")
 vTelefone=request.form("txt_telefone")
 vEmail=request.form("txt_email")
 vEndereco=request.form("txt_endereco")
 vQuartos_vend = request.form("txt_quartos_vend")
 vCidade_vend=request.form("combo1")
 vBairro_vend=request.form("combo2")
 vVila_vend=request.form("combo5")
 
 vTipo_vend=request.form("txt_tipo")
 vValor_vend=request.form("txt_valor_vend")
 vDescricao_vend=request.form("txt_descricao_vend")
  vVagas_vend=request.form("txt_vagas_vend")
 
 
 vCidade_comp=request.form("combo3")
 vBairro_comp=request.form("combo4")
 vVila_comp=request.form("combo7")
 vTipo_comp=request.form("txt_tipo2")
 vQuartos_comp = request.form("txt_quartos_comp")
 vValor_comp=request.form("txt_valor_comp")
 vDescricao_comp=request.form("txt_descricao_comp")
 vVagas_comp=request.form("txt_vagas_comp")
 
 
 if vDescricao_comp = "" then
 vDescricao_comp = "não informado"
 end if
 
 if vDescricao_vend = "" then
 vDescricao_vend = "não informado"
 end if
 
 
 if vValor_comp = "" then
 vValor_comp = "0,00"
 end if
 
 if vValor_vend = "" then
 vValor_vend = "0,00"
 end if
 
 if vLink = "" then
 vLink="não informado"
 end if
 
 
 if varCodImovel = "" then
 varCodimovel = "não informado"
 end if
 
 
 
 
 if vVagas_vend = "não informado" then
 vVagas_vend = "0"
 end if
 
 if vVagas_comp = "não informado" then
 vVagas_comp = "0"
 end if
 
 if vQuartos_vend = "não informado" then
 vQuartos_vend = "0"
 end if
 
 if vQuartos_comp = "não informado" then
 vQuartos_comp = "0"
 end if
 
 
 

  
   Dim vdata2

vdata2 = now()


  
  
 
	  
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  dim vAtendimento
	  
	  vAtendimento = request.Form("txt_atendimento")
	  
	  if vAtendimento = "" then
	  vAtendimento = "não informado"
	  end if
	  
	  													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	if vCidade_vend <> "cqualquer" and vCidade_vend <> "" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 from combo1 where id_combo1 ="& vCidade_vend
 
 rs2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs2.ActiveConnection = Conexao
 
 
 rs2.open SQL2,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_vend
 vCidade2_vend = rs2("nome_combo1")
 
 '---------------------
 rs2.close
 set rs2 = nothing
 
 '----------------------
 else
 vCidade2_vend = "não informado"
 end if
 
 
 
 
 

 
 dim rs3,SQL3
 
 if vBairro_vend <> "bqualquer" and vBairro_vend <> ""   then
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& vBairro_vend
 
 
 rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2_vend
 vBairro2_vend = rs3("nome_combo2")
 
 
rs3.close

set rs3 = nothing
 
	else
	
	vBairro2_vend = "não informado"
	
	end if
	
	
	
	
	
	
	
	
dim rs333,SQL333
 
 if vVila_vend <> "vlqualquer"  and vVila_vend <> ""  then
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="& vVila_vend
 
 rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs333.ActiveConnection = Conexao
 
 
 
 rs333.open SQL333,Conexao,2,1
 dim vVila2_vend
 vVila2_vend = rs333("nome_combo3")
 
 
 rs333.close
 
 set rs333 = nothing
	else
	
	vVila2_vend = "não informado"
	
	end if
	
		
	
	
	
	
	
	if vCidade_comp <> "cqualquer" and vCidade_comp <> "" then
	dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL5 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="& vCidade_comp
 
 rs5.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs5.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs5.ActiveConnection = Conexao
 
 
 
 rs5.open SQL5,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_comp
 vCidade2_comp = rs5("nome_combo1")
 
 rs5.close
 
 set rs5 = nothing
 
 
 else
 vCidade2_comp = "não informado"
 end if
 
 
 
 
 dim rs4,SQL4
 
 
 '------------------------------pegar os vários bairros-----------------------
 
 
 
 dim vBairro2_comp
 
 
 if vBairro_comp <> "bqualquer" and vBairro_comp <> "" then
 
 
 dim rsMultiBairros
 dim sqlMultiBairros
 
 Set rsMultiBairros = Server.CreateObject("ADODB.RecordSet")
 

dim Variavel
dim Retorno
dim i
Variavel = vBairro_comp
Retorno = Split(Variavel,", ")

i=0


for i=0 to UBound(Retorno)



sqlMultiBairros = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& Retorno(i)

rsMultiBairros.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMultiBairros.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMultiBairros.ActiveConnection = Conexao


rsMultiBairros.open sqlMultiBairros,Conexao,2,1



 while not rsMultiBairros.eof

vBairro2_comp = rsMultiBairros("nome_combo2")&", "&vBairro2_comp 

rsMultiBairros.MoveNext 
Wend



rsMultiBairros.close








next
 set rsMultiBairros = nothing
 
 else
  vBairro2_comp = "não informado" 
   
   end if
	
	
	
	
		
	
dim rs444,SQL444
 
 if vVila_comp <> "vlqualquer" and vVila_comp <> ""   then
 Set rs444 = Server.CreateObject("ADODB.RecordSet")
 SQL444 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="& vVila_comp
 
 
 rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444.ActiveConnection = Conexao
 
 rs444.open SQL444,Conexao,2,1
 dim vVila2_comp
 vVila2_comp = rs444("nome_combo3")
 
	else
	
	vVila2_comp = "não informado"
	
	end if
	
		
	
	
	
	
	
	
	
	
	
	dim varCodimovel2
	varCodimovel2 = "não informado"
	
	
	
	 
	 
	 
	 Conexao.execute"update permuta set cod_imovel='"&varCodImovel&"',cidade_vend='"&vCidade2_vend&"',cidade_comp='"&vCidade2_comp&"',bairro_vend='"&vbairro2_vend&"',bairro_comp='"&vbairro2_comp&"',tipo_vend='"&vTipo_vend&"',tipo_comp='"&vTipo_comp&"',quartos_vend='"&vQuartos_vend&"',quartos_comp='"&vQuartos_comp&"',valor_vend='"&int(vValor_vend)&"',valor_comp='"&int(vValor_comp)&"',descricao_vend='"&vDescricao_vend&"',descricao_comp='"&vDescricao_comp&"',vagas_vend='"&vVagas_vend&"',vagas_comp='"&vVagas_comp&"' where cod_permuta="&varCodPermuta
	 
	
	conexao.close
	
	set conexao = nothing
	
	
	
	
	
	  response.Redirect "conta_permuta01.asp?varSucesso_imovel="&vProprietario&"&varCodPermuta="&varCodPermuta&""
     
	  
	  
   
   
   
   
   
  
   
   
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

