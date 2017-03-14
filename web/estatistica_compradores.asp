<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<%


if  session("permissao") <> "4" and  session("permissao") <> "3" and  session("permissao") <> "5" then
response.redirect "senha02.asp"
end if 



%>
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







<%  EscreveFuncaoJavaScript ( Conexao3 ) %>



</head>
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
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
 
  
  
  <td height="18" bordercolor="#FFFFFF" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_inicial.asp">Compradores</a></strong></font></div></td>
 
  
  
   <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_inicial.asp">Permuta</a></strong></font></div></td>
 
  
  </tr>
</table>

<center>
<br>
  <a href="excluir_estatistica_compradores.asp"><img src="bt_excluir002.jpg" width="95" height="18" border="0"></a><br>
  <br>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lista de 
  buscas por compradores efetuadas </strong></font> 
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
   

  

SQL ="SELECT * FROM compradores_procurados ORDER BY cod_CompradoresProcurados DESC" 
	
 





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
%><br>
<center><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%rs.movefirst%><strong><%=rs("data")%></strong> até  <%rs.movelast%><strong><%=rs("data")%></strong><br><br> foram <strong><%=rs.RecordCount%></strong> acessos.</font></center>
<form  Method="Post" name="Formulario" action="multi_excluir_imovel.asp?varCod_imovel=<%=varCod_imovel%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&page=<%=cInt(intPage)%>" >
  <table width="1260" border="0" cellspacing="0" cellpadding="0">
    <tr bgcolor="#000000"> 
      <td width="115" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></div></td>
	  <td width="115" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
	  <td width="115" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
      <td width="115" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div></td>
      <td width="115" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
          <td width="115" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></div></td>
      <td width="115" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></div></td>	  
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negociac&atilde;o</strong></font></div></td>
      <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
      <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>IP</strong></font></div></td>
      <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data</strong></font></div></td>
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

vValor= rs("valor")
 session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)


%>
    <% session("page")=intPage%>
    <% varCod_imovel = rs("COD_CompradoresProcurados") %>
    <tr> 
      <td width="115" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("nome")%></strong></font></div></td>
	  <td width="115" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("telefone")%></strong></font></div></td>
	  <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <% if rs("cidade") = "cqualquer" then response.write "qualquer uma" else response.write rs("cidade") end if%>
          </font></div></td>
      <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <% if rs("bairro") = "bqualquer" then response.write "qualquer um" else response.write rs("bairro") end if%>
          </font></div></td>
      <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <% if rs("tipo") = "tqualquer" then response.write "qualquer um" else response.write rs("tipo") end if%>
          </font></div></td>
      <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <% if rs("quartos") = "qqualquer" then response.write "qualquer uma" else response.write rs("quartos") end if%>
          </font></div></td>
		  
		   <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <% if rs("vagas") = "gqualquer" then response.write "qualquer um" else response.write rs("vagas") end if%>
          </font></div></td>
      <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <% if rs("negociacao") = "nqualquer" then response.write "qualquer um" else response.write rs("negociacao") end if%>
          </font></div></td>
      
	  
	  <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%if rs("valor") = "vqualquer" then response.write "qualquer um" else response.write formatnumber(session("vValor1"),2)&" até "&formatnumber(session("vValor2"),2) end if%>
          </font></div></td>
      
	  
	  <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("enderecoIP")%></font></div></td>
      <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Data")%></font></div></td>
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
        <font color="#000000">Página</font> <%=cInt(intPage)%> <font color="#000000">de</font> 
        <%=cInt(intPageCount)%> </font> 
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
        <a href="?page=<%=intPage + 1%>" style="#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Próximo</b></font> 
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
<div align="center"I><font color="<%=escuro%>">Compradores</font><font color="<%=escuro%>"> 
  não encontrados</font></div>
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
<div align="center"I><font color="<%=escuro%>">Compradores n&atilde;o encontrados</font></div>
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
