

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,vProprietario,vEndereco,vFoto_Grande,vFoto_Pequena,vTitulo_Anuncio,vTexto_Anuncio,vLink_Foto,vCidade,vBairro,vTipo,vArea_Total,vArea_Construida,vQuartos,vBanheiros,vVagas,vNegociacao,vValor,vblob,vblob2,numBlob,numBlob2,vdata,vTelefone,vEmail
 Dim varSucesso_imovel,varCod_imovel,vPresenca_primeira
 dim vVila,vVila2
 
 varCod_imovel=request.querystring("varCod_imovel")
   vdata = left(now(),10)
   
   vTitulo_anuncio = request.form("txt_titulo")
   
   if vTitulo_anuncio = "" then
   vTitulo_anuncio = "não informado"
   end if
   
   vTexto_anuncio = request.form("txt_anuncio")
  
   if vTexto_anuncio = "" then
   vTexto_anuncio = "não informado"
   end if
   
  vProprietario=request.form("txt_proprietario")
  
 vPresenca_primeira = request.Form("txt_presenca_primeira")
      vEndereco=request.form("txt_endereco")
	 
	  vTelefone=request.form("txt_telefone")
	 
	  if vTelefone = "" then
	  vTelefone = "Não informado"
	  end if
	  
	  dim vCaptacao
	  
	  vCaptacao=request.form("txt_captacao")
	 
	  if vCaptacao = "" then
	  vCaptacao = "Não informado"
	  end if
	  
	  
	  
	  
	  
	  vEmail=request.form("txt_email")
	  
	  if vEmail = "" then
	  vEmail = "Não informado"
	  end if
	 
	  
	  
     
	  
	  
	  
      vLink_Foto=request.form("txt_link_foto")
	  vCidade=request.form("combo1") 
	                                                            
	 vBairro=request.form("combo2")
	 
	 vVila=request.form("combo5")
	 
	 
	 
      vTipo=request.form("txt_tipo")
      vArea_Total=request.form("txt_a_total")
	 if vArea_Total = "" then
	 vArea_Total = "00"
	 end if
	 
	 
	 
	  vArea_Construida=request.form("txt_a_constr")
	  
	  
	  if vArea_Construida = "" then
	  vArea_Construida = "00"
	  end if
	   						
									
		vQuartos=request.form("txt_quartos")
		
		
		 if vQuartos = "não informado" then
 vQuartos = "0"
 end if
		
		
      vBanheiros=request.form("txt_banheiros")
      vVagas=request.form("txt_vagas")
	  
	  if vVagas = "não informado" then
 vVagas = "0"
 end if
	  
	  
	  vNegociacao=request.form("txt_negociacao") 						
		vValor=request.form("txt_valor")
		
		dim standby, vOcupacao
		standby=request.form("txt_standby")
		vOcupacao=request.form("txt_ocupacao")
		
		dim vObs_imovel
		
		vObs_imovel=request.form("obs_imovel")
		
		if vObs_imovel = "" then
		vObs_imovel = "sem observações"
			
		
		end if
		
		
		dim vQualidade
		
		vQualidade = request.form("txt_qualidade")		
		
							
		dim vObs_proprietario
		
		vObs_proprietario=request.form("obs_proprietario")
		
		if vObs_proprietario = "" then
		vObs_proprietario = "sem observações"
		end if		
		
		
		 
		 dim vData2
		 
		 vData2 = now()
		 
		 
															  
	
	
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
 
 
 if vVila <> "vlqualquer" then
 
 dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1

 vVila2 = rs333("nome_combo3")
 
 else
 
 vVila2 = "não informado"
 
 end if
 
	
	 
 
 
 
 
	
	 
	 
	
	 
	 
	 Conexao.execute"update imoveis set proprietario='"&vProprietario&"',telefone='"&vTelefone&"',email='"&vEmail&"',endereco='"&vEndereco&"',link_foto='"&vLink_Foto&"',cidade='"&vCidade2&"',bairro='"&vBairro2&"',tipo='"&vTipo&"',area_total='"&vArea_Total&"',area_construida='"&vArea_Construida&"',quartos='"&vQuartos&"',banheiros='"&vBanheiros&"',vagas='"&vVagas&"',negociacao='"&vNegociacao&"',valor='"&vValor&"',data_atualizacao='"&vdata2&"',obs_imovel='"&vObs_imovel &"',obs_proprietario='"&vObs_proprietario &"',presenca_primeira='"&vPresenca_primeira &"',titulo_anuncio='"&vTitulo_anuncio &"',texto_anuncio='"&vTexto_anuncio &"',standby='"&standby &"',ocupacao='"&vOcupacao &"',vila='"&vVila2 &"',qualidade='"&vQualidade &"',captacao='"&vCaptacao &"' where cod_imovel="&varCod_imovel
	 
	 
	 response.Redirect "visualizar_imovel.asp?varSucesso_imovel="&vProprietario&"&varCod_imovel="&varCod_imovel&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
