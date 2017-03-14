

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodimovel,rsFoto,vFoto,vNome,strSQL2,vProposta,vimagem,vdata,vHorario,vInteresse
 
 Dim vCidade_vend,vBairro_vend,vTipo_vend,vDescricao_vend
 Dim vCidade_comp,vBairro_comp,vTipo_comp,vDescricao_comp,vLink
 Dim vProprietario,vTelefone,vEmail,vEndereco
 dim vQuartos_vend,vQuartos_comp,vValor_vend,vValor_comp
 dim vVagas_vend,vVagas_comp
 dim vPergunta
 
 vPergunta = request.form("txt_pergunta")
 
 
 vProprietario=request.form("txt_proprietario")
 vTelefone=request.form("txt_telefone")
 vEmail=request.form("txt_email")
 vEndereco=request.form("txt_endereco")
 vCidade_vend=request.form("combo1")
 vBairro_vend=request.form("combo2")
 vTipo_vend=request.form("txt_tipo_vend")
 vDescricao_vend=request.form("txt_descricao_vend")
 vVagas_vend=request.form("txt_vagas_vend")
 
 vCidade_comp=request.form("combo3")
 vBairro_comp=request.form("combo4")
 vTipo_comp=request.form("txt_tipo_comp")
 vDescricao_comp=request.form("txt_descricao_comp")
 vQuartos_comp=request.form("txt_quartos_comp")
 vVagas_comp=request.form("txt_vagas_comp")
 
 vQuartos_vend=request.form("txt_quartos_vend")
 
 vValor_vend=request.form("txt_valor_vend")
 vValor_comp=request.form("txt_valor_comp")
 
 
 
if vEmail = "" then
vEmail = "não informado"
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
 
 
 
 
 
 vLink="não informado"
 
 
 
 varCodimovel = "00"
   vimagem = "imovel00000.jpg"
  
  
  dim vdata2
  vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
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
 vCidade2_comp = "não informado"
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
	
	
 
 dim vFoto2,vFoto3,vFoto4,vFoto5,vFoto6,vFoto7
 
 vFoto = "imovel00000.jpg"
 vFoto2 = "imovel00000.jpg"
 vFoto3 = "imovel00000.jpg"
 vFoto4 = "imovel00000.jpg"
 vFoto5 = "imovel00000.jpg"
 vFoto6 = "imovel00000.jpg"
 vFoto7 = "imovel00000.jpg"
 	
	Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby) values( '"& vimagem &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& vEndereco &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vDescricao_vend &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vDescricao_comp &"','"& varCodImovel &"','"& vLink &"','"& now() &"','"& vQuartos_vend &"','"& vQuartos_comp &"','"& vValor_vend &"','"& vValor_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"& vVila_comp &"','"& vVagas_vend &"','"& vVagas_comp &"','"& "excluido" &"')" 
	
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& vValor_comp &"','"& now() &"','"& vDescricao_comp &"','"& "internet" &"','"& now() &"','"& vVila_comp &"','"& vVagas_comp &"','"& "não informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"')"
	
	
	 
	 
	 if vPergunta = "sim" then
	 

	
	 
	 Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_comprador,cod_permuta,qualidade) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& vFoto &"','"& vFoto2 &"','"& vFoto3 &"','"& vFoto4 &"','"& vFoto5 &"','"& vFoto6 &"','"& vFoto7 &"','"& vFoto7 &"','"& vFoto7 &"','"& vFoto7 &"','"& vFoto7 &"','"& vFoto7 &"','"& "icon_foto2.gif" &"','"& vCidade2_vend&"','"& vBairro2_vend &"','"& vTipo_vend &"','"& "00" &"','"& "00" &"','"& vQuartos_vend &"','"& "00" &"','"& vVagas_vend &"','"& "venda" &"','"& vValor_vend &"','"& now() &"','"& vDescricao_vend &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"& "Sem Placa" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "negócio comum" &"')"
	
		 
	 		
 dim rs444,SQL444
 Set rs444 = Server.CreateObject("ADODB.RecordSet")
 SQL444 = "select * from imoveis ORDER BY cod_imovel ASC" 
	
	rs444.open SQL444,Conexao,2,1  
	 rs444.moveLast
	
	dim varCodImovel444
	
	
	varCodImovel444 = rs444("cod_imovel")	
	
	
	
	 end if
	 
				
 dim rs555,SQL555
 Set rs555 = Server.CreateObject("ADODB.RecordSet")
 SQL555 = "select * from permuta ORDER BY cod_permuta ASC" 
	
	rs555.open SQL555,Conexao,2,1  
	 rs555.moveLast
	
	dim varCodPermuta555
	
	
	varCodPermuta555 = rs555("cod_permuta")
	
	
	 dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from compradores ORDER BY Cod_compradores ASC" 
	
	rs333.open SQL333,Conexao,2,1  
	 rs333.moveLast
	dim varNome
	dim varTelefone
	dim varCodComprador
	
	varNome = rs333("nome")
	varTelefone = rs333("telefone")
	varCodComprador = rs333("cod_compradores")
	 
	
	dim ConexaoCompra
	
	Set ConexaoCompra = Server.CreateObject("ADODB.Connection")
	
	ConexaoCompra.open dsn
	
	if vPergunta = "sim" then
	
	
	
	 ConexaoCompra.execute"update imoveis set cod_permuta='"&varCodPermuta555&"',cod_comprador='"&varCodComprador &"' where cod_imovel="&varCodImovel444
	 
	 
	 ConexaoCompra.execute"update compradores set cod_imovel='"&varCodImovel444&"',cod_permuta='"&varCodPermuta555&"' where cod_compradores="&varCodComprador
	 
	 
	  
	 ConexaoCompra.execute"update permuta set cod_imovel='"&varCodImovel444&"',cod_comprador='"&varCodComprador&"' where cod_permuta="&varCodPermuta555
	 
	 
	 else
	 
	 
	  ConexaoCompra.execute"update compradores set cod_permuta='"&varCodPermuta555&"' where cod_compradores="&varCodComprador
	 
	 
	  
	 ConexaoCompra.execute"update permuta set cod_comprador='"&varCodComprador&"' where cod_permuta="&varCodPermuta555
	 
	 
	 
	 
	
	end if
	
	
	ConexaoCompra.close
	Set ConexaoCompra = nothing
	
	
	
	 response.Redirect "mostrar_conta03.asp?varNome="&vProprietario&"&varTelefone="&vTelefone&"&varCodComprador="&varCodComprador&"&vPergunta="&vPergunta&"&varCodImovel="&varCodImovel444&"&varCodPermuta="&varCodPermuta555&""
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel incluído</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="<%=escuro%>">

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
<td width="217" height="156" ></td>    
          <td width="202" height="156" ></img>
            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Seu 
              im&oacute;vel foi inclu&iacute;do com sucesso!!</strong></font></div></td>
          <td width="217" height="156" ></td>
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
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


