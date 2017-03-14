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
 
 
 
   
  vProprietario=request.form("txt_proprietario")
 
 
 
  
  
	
	vEmail=request.form("txt_email") 
	
	if vEmail = "" then
	vEmail = "Não informado"
	end if
	
	 
    vTelefone=request.form("txt_telefone")
	
	if vTelefone = "" then
	vTelefone = "Não informado"
	end if
	
	   
      
	 
	 
	  
    vCidade=request.form("combo1")  
    vBairro=request.form("combo2")   
     vTipo=request.form("txt_tipo")   
	 
	
    vQuartos=request.form("txt_quartos")   
    
	 
	 vNegociacao=request.form("example2")
	 
	 if vNegociacao = "nqualquer" then
	 vNegociacao = "qualquer um" 
	 end if
	    
     vValor=request.form("stage22") 
	 
	 if vValor = "vqualquer" then
	 vValor = "qualquer um" 
	 end if 
	    
	 
	 dim vDescricao
		
		vDescricao=request.form("txt_descricao")
		
			
		
							
		

 

   
   
	 
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 
 
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2
 vBairro2 = rs3("nome_combo2")
	                                      
															  
	
   
	
	 
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vQuartos &"','"& vNegociacao &"','"& vValor &"','"& vdata &"','"& vDescricao &"','"& "não informado" &"','"& vData &"')"
	 
	 response.Redirect "form_incluir_compradores01.asp?varSucesso_imovel="&vProprietario&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
