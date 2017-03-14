<% response.Buffer = true %>
<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->

<%
if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if

session("permissao") = "6"

%>


<html>
<head>
<title>Compradores acessados por corretores de fora</title>

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
            
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_imoveis.asp" target="_blank" style="color:#FFFFFF">corretores externos imóveis</a></strong></font></div>
			  <%else%>
              <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>corretores externos imóveis</strong></font></div>
			
			<%end if%> 
            </td>
            
            
          <td width="135" bgcolor="<%=medio%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
            <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_compradores.asp" target="_blank" style="color:#FFFFFF">Corretores 
              externos compradores</a></strong></font></div>
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


SQL = "Select corretores_fora_compradores.id_corretores_fora,corretores_fora_compradores.corretores_fora_nome,corretores_fora_compradores.corretores_fora_id_comprador,corretores_fora_compradores.corretores_fora_data,corretores_fora_compradores.origem_franquia from corretores_fora_compradores  where origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY id_corretores_fora DESC"
if session("SearchFor") = "" then
SQL = "Select corretores_fora_compradores.id_corretores_fora,corretores_fora_compradores.corretores_fora_nome,corretores_fora_compradores.corretores_fora_id_comprador,corretores_fora_compradores.corretores_fora_data,corretores_fora_compradores.origem_franquia  from corretores_fora_compradores  where origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY id_corretores_fora DESC"
end if

if session("SearchFor") <> "" and session("SearchWhere") = "Data" then
SQL = "select corretores_fora_compradores.id_corretores_fora,corretores_fora_compradores.corretores_fora_nome,corretores_fora_compradores.corretores_fora_id_comprador,corretores_fora_compradores.corretores_fora_data,corretores_fora_compradores.origem_franquia   from corretores_fora_compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
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
SQL = "select corretores_fora_compradores.id_corretores_fora,corretores_fora_compradores.corretores_fora_nome,corretores_fora_compradores.corretores_fora_id_comprador,corretores_fora_compradores.corretores_fora_data,corretores_fora_compradores.origem_franquia   from corretores_fora_compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
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
<form action="archive_corretores_fora_compradores.asp?SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>" Method="GET" name="b2" >
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
 <form  Method="Post" name="Formulario" action="multi_excluir_corretores_fora_compradores.asp?SearchFor=<%=SearchFor%>&page=<%=intpage%>" >           
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
                  do comprador</strong></font></div></td>
           
		   
		   
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
		<% varCodCorretores_fora_compradores = rs("id_corretores_fora") %>
                <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("id_corretores_fora")%>"></td>
                    <td width="95" height="18" bgcolor="<%=color1%>"> 
                      <% if  session("permissao") = "" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6"  then%>
                     <a href="excluir_corretores_fora_compradores.asp?varCodCorretores_fora_compradores=<%=varCodCorretores_fora_compradores%>&SearchFor=<%=SearchFor%>&page=<%=intpage%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a> 
                      <%else%>
                      <img src="bt_excluir001.jpg" width="95" height="18" border="0"></img> 
                      <%end if%>
                    </td>
                  
              <td width="95" height="18" bgcolor="<%=color1%>">&nbsp; </td>
                 
				 
              <td width="20" height="20" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"></td>
				  
              <td width="140" height="20" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
                
				  
				  
                <div align="center"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
				 
                  </font> </div>
				  
				   <%if session("permissao")<> "6" then  %>
	 
	 <%else%>
	  
	   
	   
	   
	   
	   <%
		  dim rs444Comprador,SQL444Comprador
 Set rs444Comprador = Server.CreateObject("ADODB.RecordSet")
 SQL444Comprador = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,compradores.pergunta,compradores.origem_franquia  FROM compradores where cod_compradores = "&rs("corretores_fora_id_comprador")  
	
	
	rs444Comprador.CursorLocation = 3
         rs444Comprador.CursorType = 3
           rs444Comprador.ActiveConnection = Conexao
	
	
	rs444Comprador.open SQL444Comprador,Conexao,2,1  
	
	
			
	if  not rs444Comprador.eof  then
	
	
	
	
				  
				 ' while not rs444Comprador.eof
				  
				  %>
              <% if  session("permissao") <> "6" then %>
	   
	   
	   <% else %>
	  
	    <div align="right"><a href="javascript:newWindow44('visualizar_compradores33.asp?varCodCompradores=<%=rs444Comprador("cod_compradores")%>')"><img src="icone_comprador01.jpg" width="26" height="22" border="0" align="left" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rs444Comprador("cod_compradores")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs444Comprador("cod_compradores")%>')"></a> 
          <DIV STYLE="border: #000000 0px solid;  width: 570px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 0px; right: 0px;" CLASS="smalltext" ID="<%=rs444Comprador("cod_compradores")%>">
		
		
		
		<table width="580" border="0" cellspacing="0" cellpadding="0">
                
             <tr>
                        
              <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                    de atualiza&ccedil;&atilde;o</strong></font></div></td>
                        <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444Comprador("data_atualizacao")%></strong></font></td>
              </tr>
			 
			 
			 <tr>
                        <td height="30" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente</strong></font></div></td>
                        <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs444Comprador("atendimento") <> "" then response.write rs444Comprador("atendimento") else response.write "não informado" end if%></strong></font></td>
              </tr>
			  
			  
			   <tr>
                        
              <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Situa&ccedil;&atilde;o 
                  do cliente</strong></font></div></td>
                        <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs444Comprador("standby") <> "" then response.write rs444Comprador("standby") else response.write "não informado" end if%></strong></font></td>
              </tr>
			 
			 
			  
			 
			 
			  <tr>
                        <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                    do &uacute;ltimo acesso</strong></font></div></td>
                        
              <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
                <%if rs444Comprador("data_ultimo_acesso") <> "" then response.write rs444Comprador("data_ultimo_acesso") else response.write "não informado" end if %></strong>
                </font></td>
              </tr>
			  
			  
			   <tr>
                        
              <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Acessos</strong></font></div></td>
                        
              <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
                <%if rs444Comprador("acessos") <> "" then response.write rs444Comprador("acessos") else response.write "0" end if %></strong>
                </font></td>
              </tr>
			  
			  
			  
			   <tr>
                        <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;ltimo 
                    email enviado</strong></font></div></td>
                        <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6"  then %>
          <%if  UCase(rs444Comprador("atendimento")) <> UCase(Session("nome_id")) then %>
          &nbsp; 
          <%else%>
          <a href="javascript:newWindow2('visualizar_lastemail.asp?varCodCompradores=<%=varCodCompradores%>')"> 
          <%if rs444Comprador("dataLastEmail") <> "" then response.write rs444Comprador("dataLastEmail") else response.write "nenhum email enviado" end if %>
          </a> 
          <%end if%>
          <%else%>
          <a href="javascript:newWindow2('visualizar_lastemail.asp?varCodCompradores=<%=varCodCompradores%>')"> 
          <%if rs444Comprador("dataLastEmail") <> "" then response.write rs444Comprador("dataLastEmail") else response.write "nenhum email enviado" end if %>
          </a> 
          <%end if%></strong></font></td>
              </tr>
			 
			 
			 
			 <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>C&oacute;digo 
                    do comprador</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("cod_compradores")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>data 
                    de inclus&atilde;o</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("data")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			 
			 
			 
			 
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Meu 
                    nome</strong> </font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("nome")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("telefone")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                
			  
			 
              
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("email")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
              
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("endereco")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade 
                    que estou interessado</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_cidade" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("cidade")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro 
                    que estou interessado</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_bairro" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("bairro")%>" size="38" maxlength="33" align="left">
                  </td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo 
                    do im&oacute;vel desejado</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
                    <input name="txt_tipo" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%if rs444Comprador("Tipo") <> "tqualquer" then response.write rs444Comprador("Tipo") else response.write "qualquer um" end if  %>" size="38" maxlength="33" align="left">
                    </td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Número 
                    de quartos desejado</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("quartos")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;meros 
                    de vagas do im&oacute;vel desejado</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("vagas")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o 
                    que eu quero</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("negociacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Faixa 
                    de pre&ccedil;o que eu quero</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color :  <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value=" <%if vValor <> "vqualquer" then%><%=FormatNumber(rs444Comprador("Valor"),2)%> <%else%>não informado<%end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
			 
			                           
			 
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        
                      <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                          descri&ccedil;&atilde;o do im&oacute;vel que eu quero</strong></font></div></td>
                    </tr>
                    <tr> 
                        
                      <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="font-weight: bold;font-size:12;border-color :  <%=claro%>;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Comprador("descricao")%></textarea></td>
              </tr>
              
            </table>
		
</DIV>
			   
<%end if%>			   
			   
			   
			   
			   
			   
			   
			    <%
					
					'rs444Comprador.movenext
					'wend
					
					
					%>
                <%else%>
				
				
				<%
				
				
				
				
				
				
				
				
				%>
              </div>
              <%end if%>
			  
			  <%end if%>
			  
				  
				  
				  
				  
				  
				  
				  
				  
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
