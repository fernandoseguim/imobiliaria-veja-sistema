<% response.Buffer = true %>
<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->

<%

session("permissao") = "6"

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if

%>


<html>
<head>
<title>Imóveis acessados por corretores de fora</title>

<!--#include file="style6_imoveis.asp"-->

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
   openWindow = window.open(abrejanela,'openWin','width=590,height=530,resizable=yes')
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
		moveObject(menu_id,position.x+25,position.y - vertical_offset);
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
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="#FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
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
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
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
            
            
          <td width="134" height="20" bgcolor="<%=medio%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
            <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_imoveis.asp" target="_blank" style="color:#FFFFFF">corretores 
              externos imóveis</a></strong></font></div>
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
<br>
<center>

<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>A sua permissão é <%=session("permissao")%></strong></font>

</center>
<br>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere
SearchWhere = request.querystring("SearchWhere")
SearchFor = request.querystring("SearchFor")

session("SearchFor") = SearchFor
session("SearchWhere") = SearchWhere

dim rsVerifica
dim SQLVerifica



Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
	session("SearchFor")=SearchFor	


SQL = "Select corretores_fora_imoveis.id_corretores_fora,corretores_fora_imoveis.corretores_fora_nome,corretores_fora_imoveis.corretores_fora_id_imovel,corretores_fora_imoveis.corretores_fora_data,corretores_fora_imoveis.origem_franquia from corretores_fora_imoveis  where origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY id_corretores_fora DESC"
if session("SearchFor") = "" then
SQL = "Select corretores_fora_imoveis.id_corretores_fora,corretores_fora_imoveis.corretores_fora_nome,corretores_fora_imoveis.corretores_fora_id_imovel,corretores_fora_imoveis.corretores_fora_data,corretores_fora_imoveis.origem_franquia  from corretores_fora_imoveis  where origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY id_corretores_fora DESC"

end if

if session("SearchFor") <> "" and session("SearchWhere") = "Data" then
SQL = "select corretores_fora_imoveis.id_corretores_fora,corretores_fora_imoveis.corretores_fora_nome,corretores_fora_imoveis.corretores_fora_id_imovel,corretores_fora_imoveis.corretores_fora_data,corretores_fora_imoveis.origem_franquia   from corretores_fora_imoveis where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "corretores_fora_data like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "corretores_fora_data like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if

end if




if session("SearchFor") <> "" and session("SearchWhere") = "Atendente" then
SQL = "select corretores_fora_imoveis.id_corretores_fora,corretores_fora_imoveis.corretores_fora_nome,corretores_fora_imoveis.corretores_fora_id_imovel,corretores_fora_imoveis.corretores_fora_data,corretores_fora_imoveis.origem_franquia   from corretores_fora_imoveis where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "corretores_fora_nome like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "corretores_fora_nome like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if

end if


dim intRecordCount
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS.CursorType = 3






	

%>
<form action="archive_corretores_fora_imoveis.asp?SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>" Method="GET" name="b2" >
<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td bgcolor="#DAE3F0">
<table border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="<%=claro%>">
          <tr>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>Procurar</b></font></td>
            <td bgcolor="<%=claro%>"> 
              <input type="text" name="SearchFor" class="inputBox" value="<%=SearchFor%>" style="background:<%=medio%>;">
            </td>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>em</b></font></td>
            <td bgcolor="<%=claro%>"> 
              <select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
              <% if session("SearchWhere") <> "" then %>
			  <option value="<%=session("SearchWhere")%>" ><%=session("SearchWhere")%></option>
			  <% end if %>
			  

<option value="Data" >Data</option>
<option value="Atendente" >Atendente</option>
</select>
            </td>
            <td bgcolor="<%=claro%>"> 
              <input type="submit" value="Buscar" class="inputSubmit" style="background:<%=medio%>;"></td>
</tr>
</table>
</td>
</tr>
</table>
</form>


  <br>
  <br>
           
          
<%

Dim LinkTemp
'essa variável vai ser usada como contador


'as variáveis acima são usadas para trocar a cor das tabelas que conterão os valores
'dos recordsets.






dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeira página.

dim intPageCount
'Essa variável vai receber o valor da quantidade de páginas do recordset.


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

	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%>
</div>
<center>
<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Foram encontrados  <%=intRecordCount%> registros na busca.</strong></font>
</center>
   <center>      
 <form  Method="Post" name="Formulario" action="multi_excluir_corretores_fora_imoveis.asp?SearchFor=<%=SearchFor%>&page=<%=intpage%>" >           
<table width="740" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="740" height="18"><table width="740" height="18" border="0" cellpadding="0" cellspacing="0">
        <tr>
		
            <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);"></td>
            <td width="95" height="18"><% if  session("permissao") = ""  or session("permissao") = "4"  or session("permissao") = "5" or session("permissao") = "6" then %><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"><%else%><img src="bt_excluir002.jpg" width="95" height="18" border="0"></img><%end if%></td>
             
              <td width="95" height="18" bgcolor="#000000"></img></td>
              <td width="20" height="20" bgcolor="#000000" style="border:1px solid #FFFFFF;"></td>
           
		      <td width="140" height="20" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>c&oacute;digo 
                  do im&oacute;vel</strong></font></div></td>
           
		   
		   
		    <td width="250" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <strong>Atendente</strong></font></div></td>
            <td width="140" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                </strong></font></div></td>
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
If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If


%>




	<% session("page")=intPage%>
	<%
	if session("page") = "" then
	session("page") = request.querystring("page")
	end if
	%>
		
		
		
		
		<tr>
		<% varCodCorretores_fora_imoveis = rs("id_corretores_fora") %>
                <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("id_corretores_fora")%>"></td>
                    <td width="95" height="18" bgcolor="<%=color1%>"> 
                      <% if  session("permissao") = "" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6"  then%>
                      <a href="excluir_corretores_fora_imoveis.asp?varCodCorretores_fora_imoveis=<%=varCodCorretores_fora_imoveis%>&SearchFor=<%=SearchFor%>&page=<%=intpage%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a> 
                      <%else%>
                      <img src="bt_excluir001.jpg" width="95" height="18" border="0"></img> 
                      <%end if%>
                    </td>
                  
              <td width="95" height="18" bgcolor="<%=color1%>">&nbsp; </td>
                 
				 
              <td width="20" height="20" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"></td>
				  
              <td width="140" height="20" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
                
				  
				  
                <div align="center"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
				
				
				
				 
                  </font> </div>
				<%
				Set rsVerifica = Server.CreateObject("ADODB.RecordSet")
				 
				 SQLVerifica = "Select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel  from imoveis where  cod_imovel = "&rs("corretores_fora_id_imovel")
 
 
 
 rsVerifica.Open SQLVerifica, Conexao
				  %>
				  
				  
				  
				  
	  <%if session("permissao") <> "6" then  %>
	 
	 <%else%>
	  <%
 
 if not (rsVerifica.eof and rsVerifica.bof) then
		
		
				'While not rs444Imovel.eof  
				  %>
             
			 
			   
        <div align="right"><a href="javascript:newWindow333('visualizar_imovel33.asp?varCod_imovel=<%=rsVerifica("cod_imovel")%>')"><img src="icone_imovel01.jpg" width="26" height="22" border="0"  align="left" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rsVerifica("cod_imovel")%>',35)" onMouseOut="hide_info_popup(this,'<%=rsVerifica("cod_imovel")%>')"></a> 
          <DIV STYLE="border: #000000 0px solid;  width: 570px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 0px; right: 0px;" CLASS="smalltext" ID="<%=rsVerifica("cod_imovel")%>">
		
		
		<table width="570" border="0" cellspacing="0" cellpadding="0">
                
           
		    <tr>
                        
                <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                    de atualiza&ccedil;&atilde;o</strong></font></div></td>
                        
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsVerifica("data_atualizacao")%></strong></font></td>
              </tr>
		    <tr>
                        
                <td height="30" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Situação 
                    do imóvel</strong></font></div></td>
                        
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rsVerifica("imovel_em_negociacao") <> "" then response.write rsVerifica("imovel_em_negociacao") else response.write "não informado" end if %></strong></font></td>
              </tr>
		   
		     <tr>
                        
                <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação</strong></font></div></td>
                        
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rsVerifica("captacao") <> "" then response.write rsVerifica("captacao") else response.write "não informado" end if %></strong></font></td>
              </tr>
			 
			  <tr>
                        
                <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;ltimo 
                    email enviado</strong></font></div></td>
                        
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><% if rsVerifica("dataLastEmail") <> "" then response.write rsVerifica("dataLastEmail") else response.write "Nenhum email enviado" end if %></strong></font></td>
              </tr>
		   
			 
			 <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Código 
                    do imóvel</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario"  style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("cod_imovel")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>data 
                    de inclus&atilde;o</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("data")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			 
			 
			 
			 
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Proprietário</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("proprietario")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong> 
                    </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("telefone")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                
			  
			 
              
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email 
                    </strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("email")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
              
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("endereco")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_cidade" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("cidade")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong> 
                    </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_bairro" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("bairro")%>" size="38" maxlength="33" align="left">
                  </td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
                    <input name="txt_tipo" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%if rsVerifica("Tipo") <> "tqualquer" then response.write rsVerifica("Tipo") else response.write "qualquer um" end if  %>" size="38" maxlength="33" align="left">
                    </td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("quartos")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("vagas")%>" size="38" maxlength="33" align="left"></td>
              </tr>
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negociação</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("negociacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Placa</strong></font></div></td>

                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("placa")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ocupa&ccedil;&atilde;o</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("ocupacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Área 
                    Total</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("area_total")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Área 
                    Útil</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rsVerifica("area_construida")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Condomínio</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rsVerifica("condominio")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Suítes</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<% if rsVerifica("suites") <> "" then response.write rsVerifica("suites") else response.write "0" end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value=" <%if vValor <> "vqualquer" then%><%=FormatNumber(rsVerifica("Valor"),2)%> <%else%>não informado<%end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
			 
			                           
			 
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Descrição 
                          do imóvel</strong> </font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 98px; WIDTH: 288px; background:<%=claro%>; " onKeyPress="return limitfield(this, 800)"><%=rsVerifica("obs_imovel")%></textarea></td>
              </tr>
              
            </table>
		
</DIV>
		<%end if%>		 
				 
				 
				 
				  <%

              ' rs444Imovel.movenext
			  ' wend

             %>
                
				  
				  
				  
                </div>
	  
	  
	  
	  
	  <% end if%>
	  
				  
				  
				  
				  
              </td>
           
				 
				  <td width="250" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("corretores_fora_nome")%></font></div></td>
                <td width="140" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("corretores_fora_data")%></font></div></td>
              </tr>
			   <%
'-----------------------------------------------








rs.movenext
If RS.EOF Then Exit for
Next


end if
%>
		
      </table></td>
  </tr>
  
          
             
         
          
          
        
      </table></td>
  </tr>
  
</table>
</form>

<br>

<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000"> 
        <%If cInt(intPage) > 1 Then%>
        <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
        <a href="?page=<%=intPage - 1%>" style="color:#000000"> 
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF"> <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
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
<a href="?page=<%=i%>"><%if int(intPage) = int(i) then %><font color="#FF0000"><%else%><font color="#000000"><%end if%><%=i%></font>
</a> 
<%next%>
<%end if%>

	 
	   
	   
	   
	   
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"> 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
        <a href="?page=<%=intPage + 1%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>pr&oacute;ximo</b></font></a> 
        <%End If%>
      </div></td>
        </tr>
      </table>

</center>






 <%else%>
 
 
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I>Registro não encontrado</div>
</font>            
 
           
            <%
End If

%>
        
<%
  rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   
		   set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
