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
Dim vVila,vVila2

vdata = now()

vdata2 = now()
 
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
    vVila=request.form("combo5") 
	
	 vTipo=request.form("txt_tipo")   
	 
	 
	 vAreaTotal=request.form("txt_a_total")
    vAreaConstruida=request.form("txt_a_constr")  
    vQuartos=request.form("txt_quartos")   
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
		
		dim vQualidade
		
		vQualidade = request.form("txt_qualidade")
		
							
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

dim vCaptacao


vCaptacao=request.form("txt_captacao")	

if vCaptacao = "" then
vCaptacao = "não informado"
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
 
 
 
 
 
 
 
 if vVila <> "vlqualquer" then
  dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1
 
 vVila2 = rs333("nome_combo3")
 else
 vVila2 = "não informado"
 end if
 
 
 
 if vQuartos = "não informado" then
 vQuartos = "0"
 end if
 
 
 
 
	                                      
															  
	
   
	
	 
	
	Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,qualidade) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& vFoto &"','"& vFoto2 &"','"& vFoto3 &"','"& vFoto4 &"','"& vFoto5 &"','"& vFoto6 &"','"& vFoto7 &"','"& vLink_Foto &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vAreaTotal &"','"& vAreaConstruida &"','"& vQuartos &"','"& vBanheiros &"','"& vVagas &"','"& vNegociacao &"','"& vValor &"','"& vdata &"','"& vObs_imovel &"','"& vObs_proprietario &"','"& vPresenca_primeira &"','"& vTitulo_anuncio &"','"& vTexto_anuncio &"','"& standby &"','"& vOcupacao &"','"& vCaptacao &"','"& vData2 &"','"& vVila2 &"','"& vQualidade &"')"
	 
	 response.Redirect "form_incluir_imovel02.asp?varSucesso_imovel="&vProprietario&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
