<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->

<!--#include file="cores.asp"-->

<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas3 = Conexao3.Execute ( SqlMarcas3 )

While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"
Set rsCarros3 = Conexao3.Execute ( SqlCarros3 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
While NOT rsCarros3.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 


<%


dim varCidade,stringCidade,varBairro,stringBairro,varNegociacao
dim stringNegociacao,varQuartos,stringQuartos,varCidade2

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
									
	 




varNegociacao = request.querystring("example2")
varQuartos = request.querystring("txt_quartos")

	




dim varNotFind



dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao



%>




<html>
<head>
<title></title>

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


<%  EscreveFuncaoJavaScript ( Conexao3 ) %>



</head>

<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="#FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<table width="790" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr bgcolor="<%=claro%>"> 
   
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_inicial.asp">Im&oacute;veis</a></strong></font></div></td>
    
	  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_inicial.asp">Proposta</a></strong></font></div></td>
     
    
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>" > 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_inicial.asp">Email</a></strong></font></div></td>
   
   
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp">Cidades</a></strong></font></div></td>
  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro_inicial.asp">Bairros</a></strong></font></div></td>
  
  <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_vila.asp">Vila</a></strong></font></div></td>
 
  
  
  <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_inicial.asp">Compradores</a></strong></font></div></td>
 
  
  
   <td height="18" bordercolor="#FFFFFF" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_inicial.asp">Permuta</a></strong></font></div></td>
 
  
  </tr>
</table>

<br>
<center>
<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>A sua permissão é <%=session("permissao")%></strong></font>
</center>
<br>

<table width="790" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><% if  session("permissao") = "4" or  session("permissao") = "3" or  session("permissao") = "5"   then %>
        <div align="center"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="estatistica_permuta.asp" style="color:black">Estat&iacute;stica 
          da busca de permutantes</a></strong></font></div>
        <%else%>
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estat&iacute;stica 
          da busca de permutantes</strong></font></div>
        <%end if%></td>
    </tr>
  </table>


<br>




<form action="archive_permuta.asp?" Method="GET" name="b2" >

<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td bgcolor="<%=claro%>">
<table border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="<%=claro%>">
          <tr>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>Procurar</b></font></td>
            <td bgcolor="<%=claro%>">
<input type="text" name="SearchFor" class="inputBox" value="" style="background:<%=medio%>;">
            </td>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>em</b></font></td>
            <td bgcolor="<%=claro%>">
	
	
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" selected >Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone">Telefone</option>
<option value="Data" >Data</option>
<option value="cod" >Código do imóvel</option>
<option value="atendimento" >Atendimento</option>
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
 <% response.flush%>
  <%response.clear%>


</body>
</html>
