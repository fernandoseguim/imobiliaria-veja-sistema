
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
 
 
 
 
 
 dim Conexao2,rs7
 Set Conexao2 = Server.CreateObject("ADODB.Connection")
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	Conexao2.Open dsn
	dim strSQL7
	
	
	 strSQL7 = "SELECT * FROM imoveis where cod_imovel="&varCodImovel
	 rs7.CursorLocation = 3
      rs7.CursorType = 3
	 rs7.Open strSQL7, Conexao2
   if not rs7.eof then
   vimagem = rs7("foto_grande")
   else
   vimagem = "imovel00000.jpg"
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
	
	
	if vCidade_vend <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade_vend
 
 rs2.open SQL2,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_vend
 vCidade2_vend = rs2("nome_combo1")
 else
 vCidade2_vend = "não informado"
 end if
 
 
 
 
 

 
 dim rs3,SQL3
 
 if vBairro_vend <> "bqualquer"   then
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro_vend
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2_vend
 vBairro2_vend = rs3("nome_combo2")
 
	else
	
	vBairro2_vend = "não informado"
	
	end if
	
	
	
	
	
	
	
	
dim rs333,SQL333
 
 if vVila_vend <> "vlqualquer"   then
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from combo3 where id_combo3 ="& vVila_vend
 
 rs333.open SQL333,Conexao,2,1
 dim vVila2_vend
 vVila2_vend = rs333("nome_combo3")
 
	else
	
	vVila2_vend = "não informado"
	
	end if
	
		
	
	
	
	
	
	if vCidade_comp <> "cqualquer" then
	dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL5 = "select * from combo1 where id_combo1 ="& vCidade_comp
 
 rs5.open SQL5,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_comp
 vCidade2_comp = rs5("nome_combo1")
 else
 vCidade2_comp = "não informado"
 end if
 
 
 
 
 dim rs4,SQL4
 
 if vBairro_comp <> "bqualquer"  then
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select * from combo2 where id_combo2 ="& vBairro_comp
 
 rs4.open SQL4,Conexao,2,1
 dim vBairro2_comp
 vBairro2_comp = rs4("nome_combo2")
 else
  vBairro2_comp = "não informado" 
   
   end if
	
	
	
	
		
	
dim rs444,SQL444
 
 if vVila_comp <> "vlqualquer"   then
 Set rs444 = Server.CreateObject("ADODB.RecordSet")
 SQL444 = "select * from combo3 where id_combo3 ="& vVila_comp
 
 rs444.open SQL444,Conexao,2,1
 dim vVila2_comp
 vVila2_comp = rs444("nome_combo3")
 
	else
	
	vVila2_comp = "não informado"
	
	end if
	
		
	
	
	
	
	
	
	
	
	
	dim varCodimovel2
	varCodimovel2 = "não informado"
	
	
	
	 
	 
	 
	 Conexao.execute"update permuta set cod_imovel='"&varCodImovel&"',foto_imovel='"&vimagem&"',nome='"&vProprietario&"',telefone='"&vTelefone&"',email='"&vEmail&"',endereco_vend='"&vEndereco&"',link_imovel='"&vLink&"',cidade_vend='"&vCidade2_vend&"',cidade_comp='"&vCidade2_comp&"',bairro_vend='"&vbairro2_vend&"',bairro_comp='"&vbairro2_comp&"',tipo_vend='"&vTipo_vend&"',tipo_comp='"&vTipo_comp&"',quartos_vend='"&vQuartos_vend&"',quartos_comp='"&vQuartos_comp&"',valor_vend='"&vValor_vend&"',valor_comp='"&vValor_comp&"',descricao_vend='"&vDescricao_vend&"',descricao_comp='"&vDescricao_comp&"',atendimento='"&vAtendimento&"',data_atualizacao='"&vData2&"',vila_vend='"&vVila2_vend&"',vila_comp='"&vVila2_comp&"',vagas_vend='"&vVagas_vend&"',vagas_comp='"&vVagas_comp&"' where cod_permuta="&varCodPermuta
	 
	
	  response.Redirect "visualizar_permuta02.asp?varSucesso_imovel="&vProprietario&"&varCodPermuta="&varCodPermuta&""
     
	  
	  
   
   
   
   
   
  
   
   
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

