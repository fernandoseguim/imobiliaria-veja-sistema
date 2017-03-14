<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->

<!--#include file="cores.asp"-->
<html>
<head>
<title>Financiamentos</title>

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
   openWindow2 = window.open(abrejanela,'openWin','width=590,height=480,resizable=yes,scrolling=yes')
   openWindow2.focus( )
   }

</SCRIPT>






</head>
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="#FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">



<table width="675" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
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
            
    <td width="136" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
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
  
</table>
<br><br>
<div align="center"><font size="5" face="Verdana, Arial, Helvetica, sans-serif"><strong>Financiamentos</strong></font><br>
</div>
<br>
<br>

<center>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"></font> 
</center>


<% if session("permissao") <> "6" then %>
<div align="center"><img src="bt_excluir001.jpg" width="95" height="18" border="0"><br>
 
 <%else%>
 <div align="center"><a href="multi_excluir_financiamentos.asp"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></a><br>
 
 <%end if%>
 
 
  <%
Dim orderBy
orderBy = request.querystring("orderby")
Dim varCod_financiamentos
dim total
dim SQL
dim SearchFor
dim SearchWhere
SearchWhere = request("SearchWhere")
SearchFor = request.querystring("SearchFor")



dim rs4VerificaAtendimento

dim strSQL4VerificaAtendimento

dim rs4Verifica

 set rs4VerificaAtendimento = Server.CreateObject("ADODB.RecordSet")

set rs4Verifica = Server.CreateObject("ADODB.RecordSet")

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio   
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   
Conexao.execute"Insert into visualiza_paginas(visualiza_pagina_nome, visualiza_pagina_pagina, visualiza_pagina_data, origem_franquia) values( '"& session("nome_id") &"','"& "Financiamentos" &"','"& now() &"','"& session("vOrigem_Franquia") &"')" 
	
       
		
session("SearchFor") = SearchFor



SQL = "Select  financiamentos.cod_financiamentos,financiamentos.nome,financiamentos.telefone,financiamentos.email,financiamentos.valor,financiamentos.cod_imovel,financiamentos.data,financiamentos.origem_franquia,financiamentos.clique from financiamentos where origem_franquia like '"&session("vOrigem_Franquia")&"'  ORDER BY cod_financiamentos DESC"




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
 <font size="1" color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><strong> <%=intRecordCount%> 
  registros foram encontrados.</strong></font> 
</center>
           
 <form  Method="Post" name="Formulario" action="multi_excluir_financiamentos01.asp?page=<%=cInt(intPage)%>&SearchFor=<%=SearchFor%>" >           
<table width="965" height="18" border="0" cellpadding="0" cellspacing="0">
        <tr>
		
           <td width="20" height="18" bgcolor="<%=claro%>"><input type="checkbox" name="selTodos" onclick="check(true);"></td>
		   
          <td width="20" height="18" bgcolor="<%=claro%>"><% if  session("permissao") = "3" or  session("permissao") = "5" or session("permissao") = "6"  then %>
        <input name="image" type="image" src="bt_mini_excluir02.jpg" alt="Excluir" width="26" height="22" border="0">
        <%else%>
        <img src="bt_mini_excluir02.jpg" alt="Excluir" width="26" height="22" border="0"></img> 
        <%end if%></td>
		
		 <td width="25" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"></td>
		
            <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <strong>Atendimento</strong></font></div></td>
          
            <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                <strong>Nome</strong></font></div></td>
            
			<td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <strong>Telefone</strong></font></div></td>
			
			
			
			<td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email</strong></font></div></td>
          <td width="95" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;">
            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
          <td width="95" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cod imóvel</strong></font></div></td>
		 <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data</strong></font></div></td>
		
		
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
          <% varCod_financiamentos = rs("cod_financiamentos") %>
         
                
              
          
                  

 <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("cod_financiamentos")%>"></td>
          <td width="20" height="18" bgcolor="<%=color1%>">&nbsp;</td>
<td width="25" height="18" bgcolor="<% if rs("clique") <> "" then response.write "green" else response.write color1 end if%>" style="border:1px solid #FFFFFF;">&nbsp;</td>
 <td width="150" height="25" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
			
			<%


strSQL4VerificaAtendimento = "select*from compradores where telefone='"&rs("telefone")&"' or telefone02='"&rs("telefone")&"' or telefone03='"&rs("telefone")&"' "


rs4VerificaAtendimento.open strSQL4VerificaAtendimento,Conexao 

if not rs4VerificaAtendimento.eof then



if rs4VerificaAtendimento("atendimento") = Session("Admin_ID") then

response.write "<font size='4'>"&rs4VerificaAtendimento("atendimento")&"</font>"

else

response.write rs4VerificaAtendimento("atendimento")

end if




rs4VerificaAtendimento.close
else

response.write "não informado"
rs4VerificaAtendimento.close
end if
%>
			
			
			
			
			
			
			
			</font></div></td>
                   
          <td width="150" height="25" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("nome")%></font></div></td>
               
			        <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
					
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
strSQL4Verifica = "select*from compradores where telefone like'%"&rs("telefone")&"%' or telefone02 like'%"&rs("telefone")&"%' or telefone03 like'%"&rs("telefone")&"%'"

rs4Verifica.open strSQL4Verifica,Conexao 

if not rs4Verifica.eof then

%>

<a href="javascript:newWindow2('visualizar_compradores33.asp?varCodCompradores=<%=rs4Verifica("cod_compradores")%>')"><%=rs("telefone")%>i</a>
<%else%>
<%=rs("telefone")%>
<%end if%>

<%
 rs4Verifica.close
 %>
<% end if%>
					
					
					
					
					
					
					</font></div></td>
			   
			      <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("email")%></font></div></td>
             
			
          <td width="95" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if rs("valor") <> "não informado" and rs("valor") <> "" then response.write FormatNumber(rs("valor"),2) else response.write rs("valor") end if%></font></div></td>
          <td width="95" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('visualizar_imovel33.asp?varCod_imovel=<%=rs("cod_imovel")%>')" style="color:#FFFFFF;"><%=rs("cod_imovel")%></a></font></div>
            
			</td> 
			  <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("data")%></font></div>
            
			</td> 
          <%
'-----------------------------------------------






rs.movenext
If RS.EOF Then Exit for
Next


end if
%>
          
          
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




 <%else%>
 <table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    
      
    </tr>
 </table>
 
 
 
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I>Não foi encontrado nenhuma busca por financiamento.</div>
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
<!--#include file="dsn2.asp"-->
</body>
</html>
