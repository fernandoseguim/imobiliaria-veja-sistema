<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vValor,vNegociacao,vFoto
Dim vdata2
Dim vTitulo_anuncio,vTexto_anuncio

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
 
 
 
 
   
  vProprietario=request.form("txt_proprietario")
 
 
 
  
  
	
	vEmail=request.form("txt_email") 
	
	if vEmail = "" then
	vEmail = "Não informado"
	end if
	
	 
    vTelefone=request.form("txt_telefone")
	
	if vTelefone = "" then
	vTelefone = "Não informado"
	end if
	
	   
      
	 
	 
	  
    vCidade=request.form("combo1")  
    vBairro=request.form("combo2")   
     vTipo=request.form("txt_tipo")   
	 
	
    vQuartos=request.form("txt_quartos")  
	
	  if vQuartos = "não informado" then
 vQuartos = "0"
 end if
    
	 
	 vNegociacao=request.form("txt_negociacao")
	 
	 if vNegociacao = "nqualquer" then
	 vNegociacao = "qualquer um" 
	 end if
	    
     vValor=request.form("txt_valor") 
	 
	 if vValor = "vqualquer" then
	 vValor = "00" 
	 end if 
	    
	 
	 dim vDescricao
		
		vDescricao=request.form("txt_descricao")
		
			
		
							
		

 

   
   
	 
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	if vCidade <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
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
 vBairro2= "não informado"
 end if
                                     
	'------------------------------------------informações da permuta------------------------------------
	
	
	Dim vCidade_vend,vBairro_vend,vTipo_vend,vDescricao_vend
 dim vQuartos_vend,vValor_vend
 dim vPergunta
	
	vPergunta=request.form("txt_pergunta")
	  vCidade_vend=request.form("combo3")  
    vBairro_vend=request.form("combo4")   
     vTipo_vend=request.form("txt_tipo_vend")   
	 
	
    vQuartos_vend=request.form("txt_quartos_vend")
	
	 if vQuartos_vend = "não informado" then
 vQuartos_vend = "0"
 end if
	
	
	   
    vDescricao_vend = request.form("txt_descricao_vend")
	 
	 
	 
	
	    
     vValor_vend=request.form("txt_valor_vend") 
	 
	 if vValor_vend = "vqualquer" then
	 vValor_vend = "00" 
	 end if 
	
	dim vVagas
	vVagas = request.Form("txt_vagas")
	
	
	if vVagas = "não informado" then
 vVagas = "0"
 end if
	
	
	dim vOcupacao
	vOcupacao = request.Form("txt_ocupacao")
	
	dim vVagas_vend
	vVagas_vend = request.Form("txt_vagas_vend")
	
	
	if vVagas_vend = "não informado" then
 vVagas_vend = "0"
 end if
	
	
	
	
	dim rs22,SQL22
	if vCidade_vend <> "cqualquer" then
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL22 = "select * from combo1 where id_combo1 ="& vCidade_vend
 
 rs22.open SQL22,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_vend
 vCidade2_vend = rs22("nome_combo1")
 else
 vCidade2_vend = "não informado"
 end if
 
 
 if vBairro_vend <> "bqualquer" then
 dim rs33,SQL33
 Set rs33 = Server.CreateObject("ADODB.RecordSet")
 SQL33 = "select * from combo2 where id_combo2 ="& vBairro_vend
 
 rs33.open SQL33,Conexao,2,1
 dim vBairro2_vend
 vBairro2_vend = rs33("nome_combo2")
 else
 vBairro2_vend = "não informado"
 end if
 
 
 
 
 
	dim vVila_comp2
	dim vVila_comp
	 vVila_comp2=request.form("combo5")
	 session("vVila_comp2") = vVila_comp2
	 if session("vVila_comp2") = "" then
session("vVila_comp2") = request.querystring("vVila_comp2")

end if
	 
	 if session("vVila_comp2") <> "vlqualquer" then
	  dim rs88,SQL88
 Set rs88 = Server.CreateObject("ADODB.RecordSet")
 SQL88 = "select * from combo3 where id_combo3 ="& session("vVila_comp2")
 
 rs88.open SQL88,Conexao,2,1

 vVila_comp = rs88("nome_combo3")
 else
 vVila_comp = "não informado"
	end if                                      
									
	 	
	
	
	 dim vVila_vend2
	 dim vVila_vend
	 vVila_vend2=request.form("combo7")
	 session("vVila_vend2") = vVila_vend2
	 if session("vVila_vend2") = "" then
session("vVila_vend2") = request.querystring("vVila_vend2")

end if
	 
	 if session("vVila_vend2") <> "vlqualquer" then
	  dim rs99,SQL99
 Set rs99 = Server.CreateObject("ADODB.RecordSet")
 SQL99 = "select * from combo3 where id_combo3 ="& session("vVila_vend2")
 
 rs99.open SQL99,Conexao,2,1

 vVila_vend = rs99("nome_combo3")
 else
 vVila_vend = "não informado"
	end if                                      
	
	
 
 
 
 
 
 
 
	
	if vPergunta = "sim" then
	
	
	 dim vimagem,varCodImovel,vLink
 vimagem = "imovel00000.jpg"
 varCodImovel = "00"
 vLink= "não informado"
 Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby) values( '"& vimagem &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& "não informado" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vDescricao_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vDescricao &"','"& varCodImovel &"','"& vLink &"','"& vData2 &"','"& vQuartos_vend &"','"& vQuartos &"','"& vValor_vend &"','"& vValor &"','"& "internet" &"','"& vData2 &"','"& vVila_vend &"','"& vVila_comp &"','"& vVagas_vend &"','"& vVagas &"','"& "excluido" &"')"
 
Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade) values( '"& vProprietario &"','"& "não informado" &"','"& vTelefone &"','"& vEmail &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "icon_foto2.gif" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& "00" &"','"& "00" &"','"& vQuartos_vend &"','"& "00" &"','"& vVagas_vend &"','"& "venda" &"','"& vValor_vend &"','"& vData2 &"','"& vDescricao_vend &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "excluirdo" &"','"& "não informado" &"','"& "internet" &"','"& vData2 &"','"& vVila_vend&"','"& "Sem Placa"&"','"& "0"&"','"& "0"&"','"& "0"&"','"& "negócio comum"&"')"
	
	
	
	
		
 	
	
	
	
	
	
	
	 end if
	 
	
	'---------------------------------------------------------------------------------------------------														  
	
   
	
	 
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vQuartos &"','"& vNegociacao &"','"& vValor &"','"& vdata &"','"& vDescricao &"','"& "internet" &"','"& vData2 &"','"& vVila_comp &"','"& vVagas &"','"& vOcupacao &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"')"
	 
	  Conexao.Close



'----------------------------------------------------------------


	 
	 
	'---------------------------------------------------------------------
	if vPergunta = "sim" then
	
	
	dim ConexaoPermuta
	Set ConexaoPermuta = Server.CreateObject("ADODB.Connection")
	ConexaoPermuta.Open dsn
	
	
	
	dim rs555,SQL555
 Set rs555 = Server.CreateObject("ADODB.RecordSet")
 SQL555 = "select * from permuta ORDER BY cod_permuta ASC" 
	
	rs555.open SQL555,ConexaoPermuta,2,1  
	 rs555.moveLast
	
	dim varCodPermuta555
	
	
	varCodPermuta555 = rs555("cod_permuta")
	
	ConexaoPermuta.Close
	'-----------------------------------------------------------------------
	
	dim ConexaoImovel
	Set ConexaoImovel = Server.CreateObject("ADODB.Connection")
	ConexaoImovel.Open dsn
	
	
	
	 dim rs444,SQL444
 Set rs444 = Server.CreateObject("ADODB.RecordSet")
 SQL444 = "select * from imoveis ORDER BY cod_imovel ASC" 
	
	rs444.open SQL444,ConexaoImovel,2,1  
	 rs444.moveLast
	
	dim varCodImovel444
	
	
	varCodImovel444 = rs444("cod_imovel")
	
	
	ConexaoImovel.Close
	
	end if
	
	dim ConexaoCompra
	Set ConexaoCompra = Server.CreateObject("ADODB.Connection")
	ConexaoCompra.Open dsn
	
	  dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from compradores ORDER BY Cod_compradores ASC" 
	
	rs333.open SQL333,ConexaoCompra,2,1  
	rs333.movelast
	dim varNome
	dim varTelefone
	dim varCodComprador
	
	varNome = rs333("nome")
	varTelefone = rs333("telefone")
	varCodComprador = rs333("cod_compradores")
	
	
	
	'--------------------------------------------------------------------
	if vPergunta = "sim" then
	
	
	
	 ConexaoCompra.execute"update imoveis set cod_permuta='"&varCodPermuta555&"',cod_comprador='"&varCodComprador &"' where cod_imovel="&varCodImovel444
	 
	 
	 ConexaoCompra.execute"update compradores set cod_imovel='"&varCodImovel444&"',cod_permuta='"&varCodPermuta555&"' where cod_compradores="&varCodComprador
	 
	 
	  
	 ConexaoCompra.execute"update permuta set cod_imovel='"&varCodImovel444&"',cod_comprador='"&varCodComprador&"' where cod_permuta="&varCodPermuta555
	 
	
	end if
	
	ConexaoCompra.Close
	
	
	
	
	
	 response.Redirect "mostrar_conta01.asp?varNome="&vProprietario&"&varTelefone="&vTelefone&"&varCodComprador="&varCodComprador&"&vPergunta="&vPergunta&"&varCodImovel="&varCodImovel444&"&varCodPermuta="&varCodPermuta555&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sugestão incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
 
     
 
 
           
          
            
           %>
<% response.flush%>
<%response.clear%>
</body>
</html>
