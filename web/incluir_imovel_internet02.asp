<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vVagas,vValor,vNegociacao,vFoto
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
 
 
 dim vFoto2,vFoto3,vFoto4,vFoto5,vFoto6,vFoto7
   vFoto="imovel00000.jpg"
   vFoto2="mini_imovel00000.jpg"
   vFoto3="imovel00000.jpg"
   vFoto4="imovel00000.jpg"
   vFoto5="imovel00000.jpg"
   vFoto6="imovel00000.jpg"
   vFoto7="imovel00000.jpg"
   
  vProprietario=request.form("txt_proprietario")
  
  vTitulo_anuncio=request.form("txt_titulo")
  if vTitulo_anuncio = "" then
  vTitulo_anuncio = "não informado"
  end if
 
 
  vTexto_anuncio=request.form("txt_anuncio")
  if vTexto_anuncio = "" then
  vTexto_anuncio = "não informado"
  end if
  
  
   dim vPresenca_primeira
    vPresenca_primeira=request.form("txt_presenca_primeira")
	
	vEmail=request.form("txt_email") 
	
	if vEmail = "" then
	vEmail = "Não informado"
	end if
	
	 
    vTelefone=request.form("txt_telefone")
	
	if vTelefone = "" then
	vTelefone = "Não informado"
	end if
	
	   
     vEndereco=request.form("txt_endereco")  
	 
	 
	  vLink_Foto=request.form("txt_link_foto")
    vCidade=request.form("combo1")  
    vBairro=request.form("combo2")   
     vTipo=request.form("txt_tipo")   
	 
	 
	 vAreaTotal=request.form("txt_a_total")
    vAreaConstruida=request.form("txt_a_constr")  
    vQuartos=request.form("txt_quartos") 
	
	 if vQuartos = "não informado" then
 vQuartos = "0"
 end if
	
	  
     vBanheiros=request.form("txt_banheiros")   
	 
	 vNegociacao=request.form("txt_negociacao")
	  vVagas=request.form("txt_vagas")  
	  
	  
	  if vVagas = "não informado" then
 vVagas = "0"
 end if
	  
	  
	   
     vValor=request.form("txt_valor")   
	 
	 dim vOcupacao,standby
	 
	  vOcupacao=request.form("txt_ocupacao")
	  standby=request.form("txt_standby")
	  
	 
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
	 

 
if vAreaTotal = "" then
   vAreaTotal = "00"
   end if
   
   if vAreaConstruida = "" then
   vAreaConstruida = "00"
   end if


	 
	  
	
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
		vBairro2 = "não informado"
		end if
		
		
 
	dim vVila_vend2
	dim vVila_vend
	 vVila_vend2=request.form("combo5")
	 session("vVila_vend2") = vVila_vend2
	 if session("vVila_vend2") = "" then
session("vVila_vend2") = request.querystring("vVila_vend2")

end if
	 
	 if session("vVila_vend2") <> "vlqualquer" then
	  dim rs88,SQL88
 Set rs88 = Server.CreateObject("ADODB.RecordSet")
 SQL88 = "select * from combo3 where id_combo3 ="& session("vVila_vend2")
 
 rs88.open SQL88,Conexao,2,1

 vVila_vend = rs88("nome_combo3")
 else
 vVila_vend = "não informado"
	end if                                      
									
	 	
	
	
	 dim vVila_comp2
	 dim vVila_comp
	 vVila_comp2=request.form("combo7")
	 session("vVila_comp2") = vVila_comp2
	 if session("vVila_comp2") = "" then
session("vVila_comp2") = request.querystring("vVila_comp2")

end if
	 
	 if session("vVila_comp2") <> "vlqualquer" then
	  dim rs99,SQL99
 Set rs99 = Server.CreateObject("ADODB.RecordSet")
 SQL99 = "select * from combo3 where id_combo3 ="&session("vVila_comp2")
 
 rs99.open SQL99,Conexao,2,1

 vVila_comp = rs99("nome_combo3")
 else
 vVila_comp = "não informado"
	end if                                      
	
	
															  
	
   
	
	 
	
	  dim vVagas_comp
	  
	  vVagas_comp = request.form("txt_vagas_comp")
	  
	  
	  if vVagas_comp = "não informado" then
 vVagas_comp = "0"
 end if
	  
	  
	 
	 dim vCidade_comp,vBairro_comp
	 dim vCidade2_comp,vBairro2_comp
	 dim vTipo_comp,vQuartos_comp,vValor_comp
	 dim vDescricao_comp
	 
	 vCidade_comp=request.form("combo3")
	 vBairro_comp=request.form("combo4")
	 vTipo_comp=request.form("txt_tipo_comp")
	 vQuartos_comp=request.form("txt_quartos_comp")
	 
	 
	  if vQuartos_comp = "não informado" then
 vQuartos_comp = "0"
 end if
	 
	 
	 
	 vValor_comp=request.form("txt_valor_comp")
	 vDescricao_comp=request.form("txt_descricao_comp")
	 
	 if vValor_comp="" then
	 vValor_comp="00"
	 end if
	 
	 if vDescricao_comp="" then
	 vDescricao_comp="não informado"
	 end if
	 
	 
	 if vCidade_comp <> "cqualquer" then
	 
	 dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL5 = "select * from combo1 where id_combo1 ="& vCidade_comp
 
 rs5.open SQL5,Conexao,2,1
 'o recordset é aberto.
 
 vCidade2_comp = rs5("nome_combo1")
 else
 vCidade2_comp = "não informado"
 end if
 
 
 
 if vBairro_comp <> "bqualquer" then
 
 dim rs4,SQL4
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select * from combo2 where id_combo2 ="& vBairro_comp
 
 rs4.open SQL4,Conexao,2,1


 vBairro2_comp = rs4("nome_combo2")
 
 else
 vBairro2_comp = "não informado"
 end if
 
 
 dim vimagem,varCodImovel,vLink
 vimagem = "imovel00000.jpg"
 varCodImovel = "00"
 vLink= "não informado"
	
 
	 
	 dim vPergunta
	 
	 vPergunta = request.form("txt_pergunta")
	 if vPergunta = "sim" then
	  
	  
	
	
	
	 
	 
	  

 Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby) values( '"& vimagem &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& vEndereco &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vObs_imovel &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vDescricao_comp &"','"& varCodImovel &"','"& vLink &"','"& now() &"','"& vQuartos &"','"& vQuartos_comp &"','"& vValor &"','"& vValor_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"& vVila_comp &"','"& vVagas &"','"& vVagas_comp &"','"& "excluido" &"')"
 
 
 
 
 Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& vValor_comp &"','"& now() &"','"& vDescricao_comp &"','"& "internet" &"','"& now() &"','"& vVila_comp &"','"& vVagas_comp &"','"& "não informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "Internet" &"')"
	

	
	
	
	end if
		
Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& vFoto &"','"& vFoto2 &"','"& vFoto3 &"','"& vFoto4 &"','"& vFoto5 &"','"& vFoto6 &"','"& vFoto7 &"','"& vFoto &"','"& vFoto &"','"& vFoto &"','"& vFoto &"','"& vFoto &"','"& "icon_foto2.gif" &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vAreaTotal &"','"& vAreaConstruida &"','"& vQuartos &"','"& vBanheiros &"','"& vVagas &"','"& vNegociacao &"','"& vValor &"','"& now() &"','"& vObs_imovel &"','"& vObs_proprietario &"','"& "excluido" &"','"& vTitulo_anuncio &"','"& vTexto_anuncio &"','"& "excluido" &"','"& vOcupacao &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"& "Sem Placa"&"','"& "0" &"','"& "0" &"','"& "0" &"','"& "negócio comum" &"')"	 


	Conexao.Close
	
     Set Conexao = nothing

	
	 
	 
	 if vPergunta = "sim" then

dim ConexaoPermuta
	Set ConexaoPermuta = Server.CreateObject("ADODB.Connection")
	ConexaoPermuta.Open dsn
	
	
	
	dim rs555,SQL555
 Set rs555 = Server.CreateObject("ADODB.RecordSet")
 SQL555 = "select * from permuta  ORDER BY cod_permuta ASC" 
	
	rs555.open SQL555,ConexaoPermuta,2,1  
	 rs555.moveLast
	
	dim varCodPermuta555
	
	
	varCodPermuta555 = rs555("cod_permuta")
	
	ConexaoPermuta.Close
	Set ConexaoPermuta = nothing
	end if
	
	
	
	if vPergunta = "sim" then
	dim ConexaoCompra
	Set ConexaoCompra = Server.CreateObject("ADODB.Connection")
	ConexaoCompra.Open dsn
	
	  dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from compradores ORDER BY Cod_compradores ASC" 
	
	rs333.open SQL333,ConexaoCompra,2,1  
	 rs333.moveLast
	dim varNome
	dim varTelefone
	dim varCodComprador
	
	varNome = rs333("nome")
	varTelefone = rs333("telefone")
	varCodComprador = rs333("cod_compradores")
	
	ConexaoCompra.Close
	 Set ConexaoCompra = nothing

end if

	 
 
 
 
 
 	 
	 dim ConexaoImovel
	Set ConexaoImovel = Server.CreateObject("ADODB.Connection")
	ConexaoImovel.Open dsn
		
 dim rsImovel,SQLImovel
 dim varCodImovel444
 
 Set rsImovel = Server.CreateObject("ADODB.RecordSet")
 SQLImovel = "select * from imoveis ORDER BY cod_imovel ASC" 
	
	rsImovel.open SQLImovel,ConexaoImovel,2,1  
	 rsImovel.movelast
	
	
	varCodImovel444 = rsImovel("cod_imovel")
	
	
	
	
	
	if vPergunta = "sim" then
	
	
	
	 ConexaoImovel.execute"update imoveis set cod_permuta='"&varCodPermuta555&"',cod_comprador='"&varCodComprador &"' where cod_imovel="&varCodImovel444
	 
	 
	 ConexaoImovel.execute"update compradores set cod_imovel='"&varCodImovel444&"',cod_permuta='"&varCodPermuta555&"' where cod_compradores="&varCodComprador
	 
	 
	  
	 ConexaoImovel.execute"update permuta set cod_imovel='"&varCodImovel444&"',cod_comprador='"&varCodComprador&"' where cod_permuta="&varCodPermuta555
	 
	
	end if
	
		
	
	ConexaoImovel.Close
	 Set ConexaoImovel = nothing
	
	
	
	
	 
	 
	 response.Redirect "mostrar_conta02.asp?varNome="&vProprietario&"&varTelefone="&vTelefone&"&varCodComprador="&varCodComprador&"&vPergunta="&vPergunta&"&varCodImovel="&varCodImovel444&"&varCodPermuta="&varCodPermuta555&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sugestão incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"  ><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
<td width="590" height="105" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    
          <td width="202" height="156"><img src="sorriso_sugestao.jpg" width="202" height="156" border="0"></img></td>   
          <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36"></img></td>

</tr>


</table>







 
 <%
 
     
 
 
           
           
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
