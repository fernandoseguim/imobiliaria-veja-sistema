

<%
Option Explicit
%>
<!--#include file="cores.asp"-->
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if


Dim Conexao,strSQL,rs,varCodImovel,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,vHorario,vNegociacao
 
 varCodImovel = request.QueryString("varCodImovel")
   
   Dim vdata2

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
  vNome=session("nome")
      vTelefone=session("telefone")
      vEmail=session("email")
	  
	   if vNome = "" then
	  vNome = "não informado"
	  end if
	  
	  if vTelefone = "" then
	  vTelefone = "00"
	  end if
	  
	  
	  
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  dim vValorPagar
	  dim vPagamentoPagar
	  
	  vValorPagar = request.form("txt_valor")
	  vPagamentoPagar = request.form("txt_pagamento")
	  
	  
	  vProposta=" Eu quero pagar "&vValorPagar&" por esse imóvel, a forma de pagamento que eu proponho é "&vPagamentoPagar&"."                                                           
		vHorario="não informado"
		vNegociacao="não informado"													  
	
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis  Where cod_imovel = "&varCodImovel
	
	
	 Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
	vimagem=rs("foto_grande")
	
	
	
	Conexao.execute"Insert into proposta(Foto_proposta, Nome_proposta, Email_proposta, Telefone_proposta, Proposta_proposta, data_proposta,horario_proposta,interesse_proposta,cod_imovel_proposta,origem_franquia) values( '"& vimagem &"','"& vNome &"','"& vEmail &"','"& vTelefone &"','"& vProposta &"','"& now() &"','"& vHorario &"','"& vNegociacao &"','"& varCodImovel &"','"& session("vOrigem_Franquia") &"')" 
	 
	 
	
	  dim PropostaFeita
	  
	  PropostaFeita = "PropostaFeita"
	  
	  
	  
	 response.Redirect "mostrar_imovel2.asp?varCodImovel="&varCodImovel&"&PropostaFeita="&PropostaFeita&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Proposta pelo imóvel incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">

<table width="590" height=A?g?U?"462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
    <td width="590" height="105" bgcolor="<%=escuro%>" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
          <td width="217" height="156" bgcolor="<%=escuro%>" ></td>    
          <td width="202" height="156" bgcolor="<%=escuro%>" ></img> 
            <div align="center"><font color="#FFFFFF"><strong>Sua proposta foi 
              inclu&iacute;da com sucesso!!</strong></font></div></td>   
          <td width="217" height="156" bgcolor="<%=escuro%>" ></td>
</tr>

</table>



</td>
</tr>
<tr>
    <td width="590" height="117" bgcolor="<%=escuro%>" ></td>
</tr>


<tr>
    <td width="590" height="36" bgcolor="<%=escuro%>" ></img></td>

</tr>


</table>












 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>



