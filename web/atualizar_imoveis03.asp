

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,vProprietario,vEndereco,vFoto_Grande,vFoto_Pequena,vTitulo_Anuncio,vTexto_Anuncio,vLink_Foto,vCidade,vBairro,vTipo,vArea_Total,vArea_Construida,vQuartos,vBanheiros,vVagas,vNegociacao,vValor,vblob,vblob2,numBlob,numBlob2,vdata,vTelefone,vEmail
 Dim varSucesso_imovel,varCod_imovel,vPresenca_primeira
 dim vVila,vVila2
 
 varCod_imovel=request.querystring("varCod_imovel")
   vdata = left(now(),10)
   
   
   
  
   
  vProprietario=request.form("txt_proprietario")
  
 
      vEndereco=request.form("txt_endereco")
	 
	  vTelefone=request.form("txt_telefone")
	 
	  if vTelefone = "" then
	  vTelefone = "Não informado"
	  end if
	  
	  
	  
	  vEmail=request.form("txt_email")
	  
	  if vEmail = "" then
	  vEmail = "Não informado"
	  end if
	 
	  
	  
     
	  
	  
	  
     
	  vCidade=request.form("combo1") 
	                                                            
	 vBairro=request.form("combo2")
	 
	 vVila=request.form("combo5")
	 
	 
	 
      vTipo=request.form("txt_tipo")
      vArea_Total=request.form("txt_a_total")
	 if vArea_Total = "" then
	 vArea_Total = "não informado"
	 end if
	 
	 
	 
	  vArea_Construida=request.form("txt_a_constr")
	  
	  
	  if vArea_Construida = "" then
	  vArea_Construida = "não informado"
	  end if
	   						
									
		vQuartos=request.form("txt_quartos")
      vBanheiros=request.form("txt_banheiros")
      vVagas=request.form("txt_vagas")
	  
	  if vVagas = "não informado" then
 vVagas = "0"
 end if
	  
	  
	  vNegociacao=request.form("txt_negociacao") 						
		vValor=request.form("txt_valor")
		
		
		dim vObs_imovel
		
		vObs_imovel=request.form("obs_imovel")
		
		if vObs_imovel = "" then
		vObs_imovel = "sem observações"
		end if		
		
							
		dim vObs_proprietario
		
		vObs_proprietario=request.form("obs_proprietario")
		
		if vObs_proprietario = "" then
		vObs_proprietario = "sem observações"
		end if		
		
		
		
		dim vOcupacao
		vOcupacao=request.form("txt_ocupacao")
		
		if vOcupacao = "" then
		vOcupacao = "não informado"
		end if		
		
		
		 
		 dim vData2
		 
		 vData2 = now()
		 
		 
															  
	
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	if vCidade <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1   from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 
 rs2.close
 
 
 set rs2 = nothing
 else
 vCidade2 = "não informado"
 end if
 
 
 
 if vBairro <> "bqualquer" then
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& vBairro
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2
 vBairro2 = rs3("nome_combo2")
 
 rs3.close
 
 set rs3 = nothing
 else
 vBairro2 = "não informado"
 end if
 
 
 if vVila <> "vlqualquer" then
 
 dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1

 vVila2 = rs333("nome_combo3")
 
 rs333.close
 
 set rs333 = nothing
 
 else
 
 vVila2 = "não informado"
 
 end if
 
	
	 
 
 
 
 
	
	 
	 
	
	 
	 
	 Conexao.execute"update imoveis set proprietario='"&vProprietario&"',telefone='"&vTelefone&"',email='"&vEmail&"',endereco='"&vEndereco&"',cidade='"&vCidade2&"',bairro='"&vBairro2&"',tipo='"&vTipo&"',area_total='"&vArea_Total&"',area_construida='"&vArea_Construida&"',quartos='"&vQuartos&"',banheiros='"&vBanheiros&"',vagas='"&vVagas&"',negociacao='"&vNegociacao&"',valor='"&int(vValor)&"',data_atualizacao='"&vdata2&"',obs_imovel='"&vObs_imovel &"',obs_proprietario='"&vObs_proprietario &"',ocupacao='"&vOcupacao &"',vila='"&vVila2 &"' where cod_imovel="&varCod_imovel
	 
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	 
	 
	 response.Redirect "acesso_imoveis03.asp?varSucesso_imovel="&vProprietario&"&varCod_imovel="&varCod_imovel&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
