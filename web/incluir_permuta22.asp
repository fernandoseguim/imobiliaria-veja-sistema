

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
 Dim vValor_vend, vValor_comp
 Dim vVagas_vend,vVagas_comp
 dim vStandby
 
 vStandby = request.form("txt_standby")
 
 
 varCodImovel=request.form("txt_cod_imovel")
 vLink=request.form("txt_link")
 vProprietario=request.form("txt_proprietario")
 vTelefone=request.form("txt_telefone")
 vEmail=request.form("txt_email")
 vEndereco=request.form("txt_endereco")
 vQuartos_vend = request.form("txt_quartos_vend")
 vCidade_vend=request.form("combo1")
 vBairro_vend=request.form("combo2")
 vTipo_vend=request.form("txt_tipo")
 vValor_vend=request.form("txt_valor_vend")
 vDescricao_vend=request.form("txt_descricao_vend")
 vVagas_vend=request.form("txt_vagas_vend")
 
 
 
 
 vCidade_comp=request.form("combo3")
 vBairro_comp=request.form("combo4")
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
 varCodimovel = "00"
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


dim vAtendimento

vAtendimento = request.form("txt_atendimento")

if vAtendimento = "" then
vAtendimento = "não informado"
end if



  
  
 
	  
	  if vEmail = "" then
	  vEmail = "não informado"
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
 
 
 if vBairro_vend <> "bqualquer" then
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro_vend
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2_vend
 vBairro2_vend = rs3("nome_combo2")
 else
 vBairro2_vend = "não informado"
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
 vcidade2_comp = "não informado"
 end if
 
 
  
 if vBairro_comp <> "bqualquer" then
 dim rs4,SQL4
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select * from combo2 where id_combo2 ="& vBairro_comp
 
 rs4.open SQL4,Conexao,2,1
 dim vBairro2_comp
 vBairro2_comp = rs4("nome_combo2")
 else
 vBairro2_comp = "não informado"
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
 SQL99 = "select * from combo3 where id_combo3 ="& session("vVila_comp2")
 
 rs99.open SQL99,Conexao,2,1

 vVila_comp = rs99("nome_combo3")
 else
 vVila_comp = "não informado"
	end if                                      
	
	
	
	
	
	
	
	
	
	
	
	dim varCodimovel2
	varCodimovel2 = "não informado"
	
	
	
	Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby) values( '"& vimagem &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& vEndereco &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vDescricao_vend &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vDescricao_comp &"','"& varCodImovel &"','"& vLink &"','"& vData2 &"','"& vQuartos_vend &"','"& vQuartos_comp &"','"& vValor_vend &"','"& vValor_comp &"','"& vAtendimento &"','"& vData2 &"','"& vVila_vend &"','"& vVila_comp &"','"& vVagas_vend &"','"& vVagas_comp &"','"& vStandby &"')" 
	 
	 
	
	  response.Redirect "form_permuta_incluir22.asp?varSucesso_imovel="&vProprietario&""
     
	  
	  
   
   
   
   
   
  
   
   
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


