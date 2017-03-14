<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->





<%
'linha 580 SQL de busca


'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 


<%


dim varCidade,stringCidade,varBairro,stringBairro,varNegociacao
dim stringNegociacao,varQuartos,stringQuartos,varCidade2
dim varCodClicado



dim strSQL4VerificaAtendimento

dim rs4VerificaAtendimento

set rs4VerificaAtendimento = Server.CreateObject("ADODB.RecordSet")



 varCidade2 = request.querystring("combo1")
 if varCidade2 = "" then
 varCidade2 = "cqualquer"
 end if
 
session("varCidade2") = varCidade2
 
 if varCidade2 <> "cqualquer" then
 dim rrs2,SSQL2


Set Conexao4 = Server.CreateObject("ADODB.Connection")
Conexao4.open dsn

 Set rrs2 = Server.CreateObject("ADODB.RecordSet")
 SSQL2 = "select * from combo1 where id_combo1="&varCidade2
 
 rrs2.open SSQL2,Conexao4,2,1
 
 varCidade = rrs2("nome_combo1")
 else
 varCidade = varCidade2
 end if
 
 if request.QueryString("varCidade")<>"" then
  varCidade = request.QueryString("varCidade")
 session("varCidade") = varCidade
 else
 
 session("varCidade") = varCidade
 end if
 
 
 
 
 
 
 
 
 
 
 
 
 
varBairro2 = request.querystring("combo2")
if varBairro2 = "" then
varBairro2 = "bqualquer"
end if

 if varBairro2 <> "bqualquer" then
	  dim rrs3,SSQL3,conexao5
	  Set Conexao5 = Server.CreateObject("ADODB.Connection")
	  Conexao5.open dsn
 Set rrs3 = Server.CreateObject("ADODB.RecordSet")
 SSQL3 = "select * from combo2 where id_combo2 ="&varBairro2
 
 rrs3.open SSQL3,Conexao5,2,1

 varBairro = rrs3("nome_combo2")
 else
 varBairro = varBairro2
	end if                                      
									
	  if request.QueryString("varBairro")<>"" then
  varBairro = request.QueryString("varBairro")
 session("varBairro") = varBairro
 else
 
 session("varBairro") = varBairro
 end if




varNegociacao = request.querystring("example2")
varQuartos = request.querystring("txt_quartos")





dim varValor, varValor1,varValor2
 

varValor = request.QueryString("stage22")

session("varValor")=varValor
  


if request.QueryString("stage22")<>"" then
varValor = request.QueryString("stage22")
 session("varValor") = varValor
 else
 
 session("varValor") = varValor
 end if
 
  if request.QueryString("varValor")<>"" then
  varValor = request.QueryString("varValor")
 session("varValor") = varValor
 else
 
 varValor = session("varValor") 
 end if 
	
 session("varValor1")=left(varValor,10)
   session("varValor2")=right(varValor,10)	
	
	
	varValor1 = session("varValor1")
	varValor2 = session("varValor2")
	
	 if request.QueryString("varNegociacao")<>"" then
varNegociacao = request.QueryString("varNegociacao")
 session("varNegociacao") = varNegociacao
 else
 
 session("varNegociacao") = varNegociacao
 end if
   
   
   
   
    if request.QueryString("varQuartos")<>"" then
 varQuartos = request.QueryString("varQuartos")
 session("varQuartos") = varQuartos
 else
 
 session("varQuartos") = varQuartos
 end if



dim varNotFind



dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao

dim rs4Verifica 
dim strSQL4Verifica

Set rs4Verifica = Server.CreateObject("ADODB.RecordSet")



dim CompradorExiste

CompradorExiste = "não"


dim cod_Comprador
%>




<html>
<head>
<title>Imóveis clicados</title>

<script>

function check(acao){
if(document.Formulario.selTodos.checked){
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked = acao;
}
}
else
{
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked =! acao;
}
}



}





</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE="Javascript">
<!--

//showSubTopNav();
//showSubLeftNav(0, 1);

var popupVisible = false;

function show_info_popup(thisObj,menu_id,vertical_offset) {
	if (popupVisible == false) {
		menuObj = document.getElementById(menu_id);
		position = getAnchorPosition(thisObj.id);
		moveObject(menu_id,position.x+35,position.y - vertical_offset);
		changeObjectVisibility(menu_id,'visible');
		popupVisible = true;
	}
}

function hide_info_popup(thisObj,menu_id) {
	menuObj = document.getElementById(menu_id);
	// moveObject(menu_id,1,1);
	changeObjectVisibility(menu_id,'hidden');
	popupVisible = false;
}

function changeObjectVisibility(objectId, newVisibility) {
    // get a reference to the cross-browser style object and make sure the object exists
    var styleObject = getStyleObject(objectId);
    if(styleObject) {
	styleObject.visibility = newVisibility;
	return true;
    } else {
    	return false;
    }
}

function getStyleObject(objectId) {
     if(document.getElementById(objectId)){
	   return (document.getElementById(objectId).style);
     } else {
	   return false;
     }
}

function moveObject(objectId, newXCoordinate, newYCoordinate) {
    var styleObject = getStyleObject(objectId);
    if(styleObject) {
	styleObject.left = newXCoordinate;
	styleObject.top = newYCoordinate;
    }
}

function getAnchorPosition(anchor_id) {// This function will return an Object with x and y properties
	var position=new Object();
	// Logic to find position
	position.x=AnchorPosition_getPageOffsetLeft(document.getElementById(anchor_id));
	position.y=AnchorPosition_getPageOffsetTop(document.getElementById(anchor_id));
	return position;
}

function AnchorPosition_getPageOffsetLeft (el) {
	var ol=el.offsetLeft;
	while((el=el.offsetParent) != null) {
	  ol += el.offsetLeft;
	}
	return ol;
}

function AnchorPosition_getPageOffsetTop (el) {
	var ot=el.offsetTop;
	while( (el=el.offsetParent) != null) {
	  ot += el.offsetTop;
	}
	return ot;
}
//-->
</SCRIPT>











</head>
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="675" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
          <tr> 
            
            
          <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis.asp" target="_blank">Im&oacute;veis</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores.asp" target="_blank">Compradores</a></strong></font></div></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta.asp" target="_blank">Permuta</a></strong></font></div></td>
            
          <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta.asp" target="_blank">Proposta</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email.asp" target="_blank">Email</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow7777('procurar_avaliacao_corretor.asp')" style="color:#FFFFFF">Avaliação </a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ligar_urgente_comprador.asp" target="_blank" style="color:#FFFFFF">Ligar 
                urgente</a></strong></font></div></td>
            
          <td width="135" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imovel_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Imóveis 
              clicados</a></strong></font></div></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_contas_procuradas_corretor.asp" target="_blank" style="color:#FFFFFF">Contas 
                acessadas</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_imovel.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                imóvel</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_comprador.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                compradores</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo02.asp" target="_blank" style="color:#FFFFFF">Captação 
                bloco</a></strong></font></div>
				<%else%>
				<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação 
                bloco</strong></font></div>
				
				
				<%end if%></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo01.asp" target="_blank" style="color:#FFFFFF">Atendente 
                bloco</a></strong></font></div>
			<%else%>
			
			<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente 
                bloco</strong></font></div>
			
			<%end if%></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_financiamentos.asp" target="_blank" style="color:#FFFFFF">Financiamentos</a></strong></font></div>
			<%else%>
			<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Financiamentos</strong></font></div>
			<%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp" target="_blank" style="color:#FFFFFF">Cidade</a></strong></font></div>
			  <%else%>
             
			  <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div>
			 
			  <%end if%></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro.asp" target="_blank" style="color:#FFFFFF">Bairro</a></strong></font></div>
			  <%else%>
			  <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div>
			  
              <%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_vila.asp" target="_blank" style="color:#FFFFFF">Vila</a></strong></font></div>
			  <%else%>
			  
			   <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div>
			  
              <%end if%></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_comprador_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Compradores 
                Clicados</a></strong></font></div></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_procurados.asp" target="_blank">Im&oacute;veis 
          procurados</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_referencia_procurados.asp" target="_blank">Refer&ecirc;ncias 
          procuradas</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_procurados.asp" target="_blank">Permutantes 
          procurados</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_origem.asp" target="_blank">Origem</a></strong></font></div>
	  <%else%>
	  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Origem</strong></font></div>
      <%end if%>
	  
	  </td>
            <td width="135" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_tipo.asp" target="_blank">Tipos de imóveis</a></strong></font></div>
			  <%else%>
              
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipos de imóveis</strong></font></div>
		<% end if %>  
		</td> 
		   
		    <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_enviado.asp" target="_blank">Emails 
                enviados </a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_oficial.asp" target="_blank">Proposta oficial 
                 </a></strong></font></div></td>
    
  </tr>
  
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_procurados.asp" target="_blank">Compradores procurados</a></strong></font></div></td>
            
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_imoveis.asp" target="_blank" style="color:#FFFFFF">corretores externos imóveis</a></strong></font></div>
			  <%else%>
              <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>corretores externos imóveis</strong></font></div>
			
			<%end if%> 
            </td>
            
            <td width="135" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_compradores.asp" target="_blank" style="color:#FFFFFF">Corretores externos compradores</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Corretores externos compradores</strong></font></div>
			
			<%end if%> </td> 
		   
		    
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_visualiza_paginas.asp" target="_blank" style="color:#FFFFFF">Visualização de páginas</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Visualização página</strong></font></div>
			
			<%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if  (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_de_fora.asp" target="_blank" style="color:#FFFFFF">Proposta 
                de Fora</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Proposta de fora</strong></font></div>
			
			<%end if%></td>
    
  </tr>
   <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_interno.asp" target="_blank">Email interno</a></strong></font></div></td>
            
            <td width="134" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <div align="center"></div>
			  </td>
            
            <td width="135" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td> 
		   
		    
            <td width="136" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td>
            <td width="134" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td>
    
  </tr>
  
</table></td>
  </tr>
  </table>
<center>
<br>
  <br>
  <br>
  <font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lista de 
  im&oacute;veis clicados por internautas</strong></font> 
</center>


<center>
</center>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere
dim varCod_imovel

SearchWhere = request.querystring("SearchWhere")
SearchFor = request.querystring("SearchFor")

session("SearchWhere") = SearchWhere
session("SearchFor") = SearchFor


Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio
color2 = claro




Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

  

'SQL ="SELECT * FROM imovel_clicado ORDER BY cod_clicado DESC" 
	
 

 
if session("permissao") <> "6" then
SQL ="SELECT imovel_clicado.cod_clicado,imovel_clicado.nome,imovel_clicado.telefone,imovel_clicado.codigo_clicado,imovel_clicado.endereco_ip,imovel_clicado.data,imovel_clicado.tipo,imovel_clicado.quartos,imovel_clicado.vagas,imovel_clicado.cidade,imovel_clicado.bairro,imovel_clicado.valor,imovel_clicado.negociacao,imovel_clicado.tarja02,imovel_clicado.origem,imovel_clicado.atendimento,imovel_clicado.origem_franquia  FROM imovel_clicado  WHERE origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' ORDER BY cod_clicado DESC" 

else
SQL ="SELECT imovel_clicado.cod_clicado,imovel_clicado.nome,imovel_clicado.telefone,imovel_clicado.codigo_clicado,imovel_clicado.endereco_ip,imovel_clicado.data,imovel_clicado.tipo,imovel_clicado.quartos,imovel_clicado.vagas,imovel_clicado.cidade,imovel_clicado.bairro,imovel_clicado.valor,imovel_clicado.negociacao,imovel_clicado.tarja02,imovel_clicado.origem,imovel_clicado.atendimento,imovel_clicado.origem_franquia FROM imovel_clicado WHERE origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY cod_clicado DESC" 

end if
 



%>
<%

Dim LinkTemp
'essa variável vai ser usada como contador


'as variáveis acima são usadas para trocar a cor das tabelas que conterão os valores
'dos recordsets.






dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeira página.

dim intPageCount
'Essa variável vai receber o valor da quantidade de páginas do recordset.

dim intRecordCount
'Essa variável vai receber o número de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a variável intPage recebe o valor "1" na primeira página.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	
RS.Open SQL, Conn, 1, 3
'o recordset é aberto
	
RS.PageSize = 20
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%><br>
<center><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%rs.movefirst%><strong><%=rs("data")%></strong> até  <%rs.movelast%><strong><%=rs("data")%></strong><br><br> foram <strong><%=rs.RecordCount%></strong> acessos.</font></center>
<form  Method="Post" name="Formulario" action="multi_excluir_imovel_clicado01.asp?varCodClicado=<%=varCodClicado%>&page=<%=cInt(intPage)%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>" >
  <table width="1745" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="22" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <input type="checkbox" name="selTodos" onclick="check(true);">
      </td>
   
     
   
      <td width="97" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <% if   session("permissao") = "6" then %><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"><%else%><img src="bt_excluir002.jpg" width="95" height="18" border="0"></img><%end if%></td>
   
      <td width="25" bgcolor="#000000" style="border:1px solid #FFFFFF;"></td>
      <td width="60" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Origem</strong></font></div></td>
     
	 
	 
	 
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome 
          do internauta</strong></font></div></td>
	 
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>atendimento</strong></font></div></td>
	 
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone 
          do internauta</strong></font></div></td>
      
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
          de acesso</strong></font></div></td>
      
     
	  
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
          de inclus&atilde;o</strong></font></div></td>
      
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o</strong></font></div></td>
    
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>c&oacute;digo 
          do im&oacute;vel </strong></font></div></td>
      
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
     
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div></td>
      
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
    
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>ip</strong></font></div></td>
      
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></div></td>
    
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></div></td>
    
	  <td width="110" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
    
	
	
	</tr>
    <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage é menor ou igual a zero então intPage igual a "1"
	'a variável intPage sempre vai ser forçada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados então.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a página exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posição exata do primeiro registro da página correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual ao número de páginas no recordset , estamos na última 
			'página então.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor do número do último recordset.
				'intFinish corresponde ao valor do último registro da página correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'do número de registros na página menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros então
		For intRecord = 1 to RS.PageSize
		'um contador inRecord é colocado até o número de registros na página.
%>
    <%










'------------------------------------------------

%>
    <%
If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If

dim vValor


 session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)


%>
    <% session("page")=intPage%>
	<% varCodClicado= rs("cod_clicado") %>
   
    <tr> 
      <td width="22" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <input type="checkbox" name="check01" value="<%=rs("cod_clicado")%>"></td>
   
     
   
      <td width="97" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <% if  session("permissao") = "6"  then %><a href="excluir_imovel_clicado01.asp?page=<%=cInt(intPage)%>&varCodClicado=<%=varCodClicado%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a><%else%><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img><%end if%>
	 
	 <td width="25" bgcolor="<% if rs("tarja02") <> "" then response.write "green" else response.write color1 end if%>" style="border:1px solid #FFFFFF;"></td>
     
	  <td width="60" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("origem")%></strong></font></div></td>
     
	 
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("nome")%></strong></font></div></td>
	 
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
          <%


strSQL4VerificaAtendimento = "select*from compradores where telefone='"&rs("telefone")&"' or telefone02='"&rs("telefone")&"' or telefone03='"&rs("telefone")&"' "


rs4VerificaAtendimento.open strSQL4VerificaAtendimento,Conexao 

if not rs4VerificaAtendimento.eof then



if rs4VerificaAtendimento("atendimento") = Session("Admin_ID") then

'response.write "<font size='4'>"&rs4VerificaAtendimento("atendimento")&"</font>"

else

'response.write rs4VerificaAtendimento("atendimento")

end if




rs4VerificaAtendimento.close
else

'response.write "não informado"
rs4VerificaAtendimento.close
end if
%><%=rs("atendimento")%>
          </strong></font></div></td>
	 
	 
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
<% if session("permissao") <> "6" and session("permissao") <> "5" then %>
<%

strSQL4Verifica = "select*from compradores where (telefone like'"&rs("telefone")&"' or telefone02 like'"&rs("telefone")&"' or telefone03 like'"&rs("telefone")&"') and (atendimento='"&UCase(Session("Admin_ID"))&"' or atendimento='"&Session("Admin_ID")&"')"


rs4Verifica.open strSQL4Verifica,Conexao 

if not rs4Verifica.eof then

CompradorExiste = "sim"
cod_Comprador = rs4Verifica("cod_compradores")

%>

<a href="javascript:newWindow2('visualizar_compradores33.asp?varCodCompradores=<%=rs4Verifica("cod_compradores")%>')"><%=rs("telefone")%>   </a>

 <%else%>
 
 <% CompradorExiste = "não"%>
 
 não informado
 
 <%end if%>
 <%
 rs4Verifica.close
 %>

<% else%>


<%
strSQL4Verifica = "select*from compradores where telefone like'"&rs("telefone")&"' or telefone02 like'"&rs("telefone")&"' or telefone03 like'"&rs("telefone")&"'"

rs4Verifica.open strSQL4Verifica,Conexao 

if not rs4Verifica.eof then

%>

<a href="javascript:newWindow2('visualizar_compradores33.asp?varCodCompradores=<%=rs4Verifica("cod_compradores")%>')"><%=rs("telefone")%></a>


<%else%>
<%="não informado"%>
<%end if%>

<%
 rs4Verifica.close
 %>
<% end if%>
 


</strong></font></div></td>
	  
       
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("data")%><strong></strong></font></div></td>
	 
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
		
		
		 
        <%







'---------------------------------------------------------------------------
dim rs444P
dim strSQL444P

    Set rs444P = Server.CreateObject("ADODB.RecordSet")
'se no cliente ou no servidor.


	strSQL444P = "SELECT compradores.data,compradores.cod_compradores,compradores.telefone,compradores.telefone02,compradores.telefone03 FROM compradores where (telefone like '%"&rs("telefone")&"%' or telefone02 like '%"&rs("telefone")&"%' or telefone03 like '%"&rs("telefone")&"%')"
	
	
	
	
 
	
	
	
	
	
	 
	
	 
		
Rs444P.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
Rs444P.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.
 
	 
	 rs444P.Open strSQL444P,Conexao 
	 
	 
	 
	
	   
	   
     %>
        <%
	 
 if not rs444P.eof then

 %>
        
		<%=rs444P("data")%>
       
	   
	    <%  end if
		 rs444P.close 
		  set rs444P = nothing %>
		
		
		
		
		
		</font></div></td>
	 
      
		  
		  
      
      
	  
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("negociacao")%></strong></font></div></td>
     
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('visualizar_imovel33.asp?varCod_imovel=<%=rs("codigo_clicado")%>')"><%=rs("codigo_clicado")%></a></strong></font></div></td>
    
	  <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cidade")%></font></div></td>
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("bairro")%></font></div></td>
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%if rs("tipo") <> "tqualquer" then response.write rs("tipo") else response.write "qualquer um" end if%>
          </font></div></td>
    
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("endereco_ip")%></font></div></td>
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("quartos")%></font></div></td>
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("vagas")%></font></div></td>
    
		 
      <td width="110" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=formatNumber(rs("valor"),2)%></font></div></td>
	
	
	</tr>
    <%
'-----------------------------------------------









rs.movenext
If RS.EOF Then Exit for
Next

%>
  </table>
</form>





<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#003366"> 
        <%If cInt(intPage) > 1 Then%>
        <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
      <a href="?page=<%=intPage - 1%>" style="color:#000000">
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF">
	
	
	
	
      <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
       
	   <%dim cont,cont2,i %>
	 
	 
	 <%if int(intPageCount) > 1 then%>
<%
If int(intPage)-5 > 1 then
cont=int(intPage)-5
else
cont=1
end if
%>
<%if cint(cont+10) > cint(intPageCount) then 
cont2=int(intPageCount)
else
cont2=int(cont)+10
end if
%>
<%for i=int(cont) to int(cont2)%>
<%

%>
<a href="?page=<%=i%>" style="color:#000000"><%if int(intPage) = int(i) then %><font color="#FF0000"><%else%><font color="#000000"><%end if%><%=i%></font>
</a> 
<%next%>
<%end if%>

	 
	   
	   
	   
	   
        <%End If%></font>
        </div>
	
	
	
	
	
	</td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
        <a href="?page=<%=intPage + 1%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Próximo</b></font> 
        </a> 
        <%End If%>
        </font></div></td>
        </tr>
      </table>










 <%else%>
 
  <table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
 </table>
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Im&oacute;veis não encontrados</font></div>
</font> 
<%
End If%>
<%else%>
<table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
	
 </table>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Im&oacute;veis </font><font color="<%=escuro%>"> 
  não encontrados</font></div>
</font> 
<%
End if
%>
<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui é criada uma variável "groups" que receberá os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a variável group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero até o número de elementos do array "groups" */

group2[i2]=new Array()
/* aqui é criado o array "group" que receberá valores conforme o número de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receberá valores de opções. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receberá valores de opções. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("201,00 até 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 até 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 até 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 até 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 até 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 até 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 até 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 até 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 até 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Até  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 até 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 até 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 até 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 até 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 até 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 até 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 até 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 até 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 até 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("Acima de 400.000,00","0000400001 1000000000")









/* aqui temos um array bidimensional "group" que receberá valores de opções. */


var temp2=document.doublecombo.stage22
/* aqui a variável "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui é criada a função "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que dá um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que é escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" será o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a variável "location" recebe os valores de "stage2" que corresponde ao endereço de
link para o carregamento de página. */


//-->
</script>     
<%
  rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
           %>
  <% response.flush%>
  <%response.clear%>
  <br>
  <br>
  <center>

</center>




</body>
</html>

