

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,varCod_imovel,SearchFor,objFSO,FotoPeqEx,vProprietario,vEndereco 
dim rs3
dim strSQL3
dim varCidade, varBairro,varNegociacao,varQuartos,page,SearchWhere,varValor
dim varValor1,varValor2
dim varFoto

dim varCidade2
dim varBairro2
dim varTipo
dim varDevedor 
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
dim varCaptacao 
dim varStandbyImovel 
dim varVagas


 varCidade2 = request.QueryString("varCidade2")
 varBairro2 = request.QueryString("varBairro2")
 varTipo = request.QueryString("varTipo")
 varDevedor = request.QueryString("varDevedor")
 varSuites = request.QueryString("varSuites") 
 varPiscina = request.QueryString("varPiscina") 
 varPortaria = request.QueryString("varPortaria") 
 varQuintal = request.QueryString("varQuintal") 
 varQuadras = request.QueryString("varQuadras") 
 varEdicula = request.QueryString("varEdicula") 
 varCondominio = request.QueryString("varCondominio") 
 varCondominio1 = request.QueryString("varCondominio1") 
 varCondominio2 = request.QueryString("varCondominio2") 
 varAreaTotal = request.QueryString("varAreaTotal") 
 varAreaTotal1 = request.QueryString("varAreaTotal1") 
 varAreaTotal2 = request.QueryString("varAreaTotal2") 
 varOcupacao = request.QueryString("varOcupacao") 
 varCaptacao = request.QueryString("varCaptacao") 
 varStandbyImovel = request.QueryString("varStandbyImovel")
varVagas = request.QueryString("varVagas")





 varCod_imovel = request.QueryString("varCod_imovel")
 SearchFor = request.querystring("SearchFor")
 SearchWhere = request.querystring("SearchWhere")
 varCidade = request.querystring("varCidade")
  varBairro = request.querystring("varBairro")  
  varNegociacao = request.querystring("varNegociacao")
  varQuartos = request.querystring("varQuartos")
  varValor1 = request.querystring("varValor1")
  varValor2 = request.querystring("varValor2")
  varValor = request.querystring("varValor")
   page = request.querystring("page")   
   varFoto = request.querystring("varFoto")                                                     
															  
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set rs3 = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	strSQL = "select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  from imoveis,primeira_pagina where primeira_pagina.foto_grande = imoveis.foto_grande and imoveis.cod_imovel = "&varCod_imovel
	strSQL3 = "select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  from imoveis where cod_imovel = "&varCod_imovel
	
	Conexao.Open dsn
	
	RS.CursorLocation = 3
RS.CursorType = 3

RS3.CursorLocation = 3
RS3.CursorType = 3

        rs.Open strSQL, Conexao 
		rs3.Open strSQL3, Conexao
		if not rs3.eof then
		
		If objFSO.FileExists(Server.MapPath(rs3("Foto_pequena"))) = True and objFSO.GetFileName(Server.MapPath(rs3("Foto_pequena"))) <> "mini_imovel00000.jpg" and objFSO.GetFileName(Server.MapPath(rs3("Foto_pequena"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs3("Foto_pequena")))
	 
	 end if
		end if
		
		
		
		
		
		if not rs.eof then
		
		
	FotoPeqEx=rs("Foto_Grande")
	vProprietario=rs("proprietario")
	 vEndereco=rs("endereco")
	 
	Conexao.execute"Delete from primeira_pagina where Foto_grande like '"&FotoPeqEx&"' and proprietario like '"&vProprietario&"' and endereco like '"&vEndereco&"'"
	Conexao.execute"Delete from imoveis where cod_imovel="&VarCod_imovel
	
	If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True and objFSO.GetFileName(Server.MapPath(rs("Foto_grande"))) <> "imovel00000.jpg" Then
	
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande")))
	
	 
	 end if
	 
	
	 
	
	 If objFSO.FileExists(Server.MapPath(rs("Foto_grande2"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande2"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande2")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande3"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande3"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande3")))
	 
	 end if
	 
	 
	 If objFSO.FileExists(Server.MapPath(rs("Foto_grande4"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande4"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande4")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande5"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande5"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande5")))
	 
	 end if
	 
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande6"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande6"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande6")))
	 
	 end if
	 
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande7"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande7"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande7")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande8"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande8"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande8")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande9"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande9"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande9")))
	 
	 end if
	 
	  If objFSO.FileExists(Server.MapPath(rs("Foto_grande10"))) = True  and objFSO.GetFileName(Server.MapPath(rs("Foto_grande10"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande10")))
	 
	 end if
	 
	 
	else
	Dim rs2,strSQL2
	
	Set rs2 = Server.CreateObject("ADODB.RecordSet")
	strSQL2 = "select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  from imoveis where cod_imovel = "&varCod_imovel
	rs2.open strSQL2,Conexao
	
	If objFSO.FileExists(Server.MapPath(rs2("Foto_grande"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande"))) <> "imovel00000.jpg" Then
	
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande")))
	
	 
	 end if
	 
	
	 
	
	 If objFSO.FileExists(Server.MapPath(rs2("Foto_grande2"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande2"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande2")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande3"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande3"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande3")))
	 
	 end if
	 
	 
	 If objFSO.FileExists(Server.MapPath(rs2("Foto_grande4"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande4"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande4")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande5"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande5"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande5")))
	 
	 end if
	 
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande6"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande6"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande6")))
	 
	 end if
	 
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande7"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande7"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande7")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande8"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande8"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande8")))
	 
	 end if
	 
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande9"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande9"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande9")))
	 
	 end if
	 
	  If objFSO.FileExists(Server.MapPath(rs2("Foto_grande10"))) = True  and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande10"))) <> "imovel00000.jpg" Then
	 objFSO.DeleteFile(Server.MapPath(rs2("Foto_grande10")))
	 
	 end if
	 
	 
	
	Conexao.execute"Delete from imoveis where cod_imovel="&VarCod_imovel
	  'response.redirect "archive_imoveis.asp?SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varCidade="&varCidade&"&varBairro="&varBairro&"&varNegociacao="&varNegociacao&"&varQuartos="&varQuartos&"&page="&cInt(Page)&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varValor="&varValor&"&varFoto="&varFoto&""
	  
	  response.redirect "archive_imoveis.asp?page="&cInt(Page)&"&varCidade="&varCidade&"&varCidade2="&varCidade2&"&varBairro="&varBairro&"&varBairro2="&varBairro2&"&varNegociacao="&varNegociacao&"&varTipo="&varTipo&"&varQuartos="&varQuartos&"&varVagas="&varVagas&"&SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varValor="&varValor&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varFoto="&varFoto&"&varDevedor="&varDevedor&"&varSuites="&varSuites&"&varPiscina="&varPiscina&"&varPortaria="&varPortaria&"&varQuintal="&varQuintal&"&varQuadras="&varQuadras&"&varEdicula="&varEdicula&"&varCondominio="&varCondominio&"&varCondominio1="&varCondominio1&"&varCondominio2="&varCondominio2&"&varAreaTotal="&varAreaTotal&"&varAreaTotal1="&varAreaTotal1&"&varAreaTotal2="&varAreaTotal2&"&varOcupacao="&varOcupacao&"&varCaptacao="&varCaptacao&"&varStandbyImovel="&varStandbyImovel&""

	  
	  end if
	 
	 
	 '-----------------------------------
	 
	 rs.close
	 
	 set rs = nothing
	 
	 '----------------------------------------
	 
	 '-----------------------------------
	 
	 rs2.close
	 
	 set rs2 = nothing
	 
	 '----------------------------------------
	 
	 
	 
	 '-----------------------------------
	 
	 rs3.close
	 
	 set rs3 = nothing
	 
	 '----------------------------------------
	 
	 
	 '-----------------------------------
	 
	 
	 
	 set objfso = nothing
	 
	 '----------------------------------------
	 
	 
	 '-----------------------------------
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	 '----------------------------------------
	 
	 
	 
	 
	 response.redirect "archive_imoveis.asp?page="&cInt(Page)&"&varCidade="&varCidade&"&varCidade2="&varCidade2&"&varBairro="&varBairro&"&varBairro2="&varBairro2&"&varNegociacao="&varNegociacao&"&varTipo="&varTipo&"&varQuartos="&varQuartos&"&varVagas="&varVagas&"&SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varValor="&varValor&"&varValor1="&varValor1&"&varValor2="&varValor2&"&varFoto="&varFoto&"&varDevedor="&varDevedor&"&varSuites="&varSuites&"&varPiscina="&varPiscina&"&varPortaria="&varPortaria&"&varQuintal="&varQuintal&"&varQuadras="&varQuadras&"&varEdicula="&varEdicula&"&varCondominio="&varCondominio&"&varCondominio1="&varCondominio1&"&varCondominio2="&varCondominio2&"&varAreaTotal="&varAreaTotal&"&varAreaTotal1="&varAreaTotal1&"&varAreaTotal2="&varAreaTotal2&"&varOcupacao="&varOcupacao&"&varCaptacao="&varCaptacao&"&varStandbyImovel="&varStandbyImovel&""

     
	  
	  
   
   
   
   
   
  
   
   
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



