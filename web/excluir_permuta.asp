

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,varCod_imovel,SearchFor,objFSO,FotoPeqEx,vProprietario,vEndereco 
dim rs3,rs333
dim strSQL3
dim SearchWhere

dim varCodPermuta,page
 varCodPermuta = request.QueryString("varCodPermuta")

 SearchFor = request.querystring("SearchFor")
 SearchWhere = request.querystring("SearchWhere")
 page = request.QueryString("page")
 
 
                                                            
															  
	
	Set rs3 = Server.CreateObject("ADODB.RecordSet")
	Set rs333 = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	
	strSQL3 = "select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais   from permuta where cod_permuta = "&varCodPermuta
	
	Conexao.Open dsn
	
	

RS3.CursorLocation = 3
RS3.CursorType = 3

        
		rs3.Open strSQL3, Conexao
		
		
		if not rs3.eof then
		
		dim strSQL333
		strSQL333 = "select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  from imoveis where foto_pequena = '"&rs3("foto_imovel")&"' or foto_grande = '"&rs3("foto_imovel")&"' or foto_grande1 = '"&rs3("foto_imovel")&"'"
		
		
		dim valor
		
		valor = rs3("foto_imovel")
		
		if valor <> "" then
		
		valor = rs3("foto_imovel")
		
		else
		
		valor = "imovel00000.jpg"
		
		end if
		
		
		
		rs333.Open strSQL333, Conexao
		if  rs333.eof then
		
		If objFSO.FileExists(Server.MapPath(valor)) = True and objFSO.GetFileName(Server.MapPath(valor)) <> "mini_imovel00000.jpg" and objFSO.GetFileName(Server.MapPath(valor)) <> "imovel00000.jpg" and  objFSO.GetFileName(Server.MapPath(valor)) <> "" Then
	 objFSO.DeleteFile(Server.MapPath(rs3("Foto_imovel")))
	 
	 end if
		
		end if
		
		
	
	
	Conexao.execute"Delete from permuta where cod_permuta="&VarCodPermuta
	 
	 
	 
	 
	 
	
	
     end if
	  
	  
	  
	  rs333.close
	  set rs333 = nothing
	  
	  set objfso = nothing
	  
	  rs3.close
	  
	  set rs3 = nothing
	  
	  
	  
	  
	   response.redirect "archive_permuta.asp?SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&page="&page&""
	
   
   
   
   
   
  
   
   
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


