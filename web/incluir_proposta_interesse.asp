

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodimovel,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,vHorario,vInteresse
 vInteresse=request.form("txtNegociacao")
 varCodimovel = request.QueryString("varCodimovel")
   vimagem = "imovel00000.jpg"
   
   
   
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
 
 
 
 
  vNome=request.form("txtNome")
      vTelefone=request.form("txtTelefone")
      
	  vEmail=request.form("txtEmail")
	 
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  
	  vProposta=request.form("txtProposta")                                                           
		vHorario=request.form("txtHorario")													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	dim  varCodimovel2
	varCodimovel2 = "não informado"
	
	dim varSucesso
	
	
	
	Conexao.execute"Insert into proposta(Foto_proposta, Nome_proposta, Email_proposta, Telefone_proposta, Proposta_proposta, data_proposta,horario_proposta,interesse_proposta,cod_imovel_proposta) values( '"& vimagem &"','"& vNome &"','"& vEmail &"','"& vTelefone &"','"& vProposta &"','"& vdata &"','"& vHorario &"','"& vInteresse &"','"& varCodimovel2 &"')" 
	 
	 response.Redirect "proposta_interesse.asp?varSucesso="&"Proposta enviada com sucesso."&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Pedido incluído</title>
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
<td width="217" height="156" ></td>    
          <td width="202" height="156" ><img src="sorriso_pedido.jpg" width="202" height="156" border="0"></img></td>   
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


