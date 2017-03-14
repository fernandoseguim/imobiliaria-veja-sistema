<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vValor,vNegociacao,vFoto
Dim vdata2
Dim vTitulo_anuncio,vTexto_anuncio
dim vVila,vVila2
vdata2 = now()

vdata = now()

 
 
 
   
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
    
	
	 if vQuartos = "não informado" then
 vQuartos = "0"
 end if
	
	 
	 vNegociacao=request.form("example2")
	 
	 if vNegociacao = "nqualquer" then
	 vNegociacao = "qualquer um" 
	 end if
	    
     vValor=request.form("stage22") 
	 
	 if vValor = "vqualquer" then
	 vValor = "0" 
	 end if 
	    
	 
	 dim vDescricao
		
		vDescricao=request.form("txt_descricao")
		
		if vDescricao = "" then
		vDescricao = "não informado"
		end if
		
		
		dim vAtendimento
		
		vAtendimento=request.form("txt_atendimento")
		
		if vAtendimento = "" then
		vAtendimento = "não informado"
		end if
			
		
		dim vVagas
		
		vVagas=request.form("txt_vagas")
		
		 
		
		
		
		if vVagas = "" then
		vVagas = "0"
		end if
		
		
		
		if vVagas = "não informado" then
 vVagas = "0"
 end if
		
		
	
 
 dim vOcupacao
		
		vOcupacao=request.form("txt_ocupacao")
		
		if vOcupacao = "" then
		vOcupacao = "não informado"
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
	                                      
															  
	
   vVila=request.form("combo5")
	 
 if vVila <> "vlqualquer" then
  dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1
 
 vVila2 = rs333("nome_combo3")
 else
 vVila2 = "não informado"
 end if
 
	
	 
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vQuartos &"','"& vNegociacao &"','"& vValor &"','"& vdata &"','"& vDescricao &"','"& vAtendimento &"','"& vData &"','"& vVila2 &"','"& vVagas &"','"& vOcupacao &"')"
	 
	 response.Redirect "form_incluir_compradores02.asp?varSucesso_imovel="&vProprietario&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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
