<%
Option Explicit
%>
<!--#include file="dsn.asp"-->


<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vVagas,vValor,vNegociacao,vFoto
Dim vdata2
dim vVila

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
 
 
 
  
	
    vCidade=request.form("txt_cidade")  
   
	  vBairro = request.form("txt_bairro")
	  vVila = request.form("txt_vila")
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
 
	                           
	dim rs4,SQL4
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 from combo3 where nome_combo3 like '"& vVila&"' and id_combo2 like '"&vBairro&"' and id_combo1 like '"&vCidade&"'"
 
 rs4.open SQL4,Conexao,2,1
 
 if not rs4.eof then
 dim existente
 existente = " já existe no banco de dados"
 
 response.Redirect "form_incluir_vila.asp?varExistente="&vVila&existente&""
 
 end if						   
							              
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 dim vCidade2
 vCidade2 = rs2("nome_combo1")														  
	
   
   dim rs33,SQL33
 Set rs33 = Server.CreateObject("ADODB.RecordSet")
 SQL33 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& vBairro
 
 rs33.open SQL33,Conexao,2,1
 dim vBairro2
 vBairro2 = rs33("nome_combo2")														  
	
   
	
	 
	
	Conexao.execute"Insert into combo3(nome_combo3,id_combo1,id_combo2,cidade_combo3,bairro_combo3,data_combo3) values( '"& vVila &"','"& vCidade &"','"& vBairro &"','"& vCidade2 &"','"& vBairro2 &"','"& now() &"')"
	 
	 dim varCidade
	 
	 
	 rs4.close
	 
	 set rs4 = nothing
	 
	 rs2.close
	 
	 set rs2 = nothing
	 
	 rs33.close
	 
	 set rs33 = nothing
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	 
	 
	 
	 response.Redirect "form_incluir_vila.asp?varSucesso_bairro="&vVila&"&varCidade="&vCidade2&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
          <td width="202" height="156"></img></td>   
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
