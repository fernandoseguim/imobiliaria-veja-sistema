<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%response.Buffer = true %>
<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas3.ActiveConnection = Conexao3
	
	
	rsMarcas3.Open SqlMarcas3, Conexao3



While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros3.ActiveConnection = Conexao3
	
	
	rsCarros3.Open SqlCarros3, Conexao3




'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1
While NOT rsCarros3.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
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
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1   FROM combo1 ORDER BY nome_combo1" 


Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open Sql3, Conexao3

%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")


 






   vValor=request.form("stage22")
   session("vValor")=vValor
   
   if session("vValor") = "" then
session("vValor") = request.querystring("vValor")
end if
   
   session("vValor1")=left(session("vValor"),10)
   session("vValor2")=right(session("vValor"),10)
   
   






dim vnegociacao

 vnegociacao=request.form("example2")
 session("vnegociacao") = vnegociacao
 
 if session("vnegociacao") = "" then
 session("vnegociacao") = request.querystring("vnegociacao")
 end if
 

dim vOcupacao
 
 vOcupacao = request.Form("txt_ocupacao")
 session("vOcupacao") = vOcupacao
 
 if session("vOcupacao") = "" then
session("vOcupacao") = request.querystring("vOcupacao")
end if



dim vVagas
 
 vVagas = request.Form("txt_garagem")
 session("vVagas") = vVagas
 
 if session("vVagas") = "" then
session("vVagas") = request.querystring("vVagas")
end if


dim vQuartos
 
 vQuartos = request.Form("txt_quartos")
 session("vQuartos") = vQuartos
 
 if session("vQuartos") = "" then
session("vQuartos") = request.querystring("vQuartos")
end if

 vTipo=request.form("txt_tipo")
  session("vTipo") = vTipo
  
  if session("vTipo") = "" then
  session("vTipo") = request.querystring("vTipo")
  end if
  
  

dim vCidade2

 vCidade2=request.form("combo1")
	
	
	
	session("vCidade2") = vCidade2
	 if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")

end if


dim vBairro2
	 vBairro2=request.form("combo2")
	 session("vBairro2") = vBairro2
	 if session("vBairro2") = "" then
session("vBairro2") = request.querystring("vBairro2")

end if


dim vvila2,vvila
vvila2=request.form("combo5")
	
	
	
	session("vvila2") = vvila2
	 if session("vvila2") = "" then
session("vvila2") = request.querystring("vvila2")

end if


dim Nome
 dim Telefone
 dim email
 
 Nome = request.form("txt_nome")
 
 
 session("nome") = Nome
 
 
 Telefone = request.form("txt_telefone")
 

 session("telefone") = Telefone


email = request.form("txt_email")
 

 session("email") = email



if session("nome") = "" then
session("nome") = request.querystring("nome")

end if


if session("telefone") = "" then
session("telefone") = request.querystring("telefone")

end if

if session("email") = "" then
session("email") = request.querystring("email")

end if




dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	if session("vCidade2") <> "cqualquer" then
	
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 ="&session("vCidade2")&"  ORDER BY nome_combo2" 
	else
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	end if
	
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao


dim rs44,strSQL44
   
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	
	if session("vBairro2") <> "bqualquer" then
	strSQL44 = "SELECT  combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where id_combo2 ="&session("vBairro2")&"  ORDER BY nome_combo3" 
	else
	strSQL44 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3   ORDER BY nome_combo3"
	end if 
	
	
	
	
	rs44.Open strSQL44, Conexao




%>



<%
Function EscreveFuncaoJavaScript222 ( Conexao333 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros222 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo2.options[doublecombo.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 


Set rsMarcas333 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas333.ActiveConnection = Conexao333
	
	
	rsMarcas333.Open SqlMarcas333, Conexao333

While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 FROM combo3 where id_combo2 =" & rsMarcas333("id_combo2")&""


Set rsCarros333 = Server.CreateObject("ADODB.RecordSet")

	rsCarros333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros333.ActiveConnection = Conexao333
	
	
	rsCarros333.Open SqlCarros333, Conexao333




'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros333.EoF

Response.Write "doublecombo.combo5.options[" & i & "] = new Option('" & rsCarros333("nome_combo3") & "','" & rsCarros333("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros333.MoveNext
Wend


Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas333.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%

'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open dsn

'

Sql333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 



Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs333.ActiveConnection = Conexao333
	
	
	rs333.Open Sql333, Conexao333


 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
 
 
 dim vCadastrado
 
 vCadastrado = "nao"
 
 
 
 
 
 
 
 
'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo   FROM tipo  ORDER BY tipo ASC" 
	 
	

	rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao3
	
	

	 
	 
	 
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao333









'-------------------------------------------------------------------------------------------------


 
%> 





<html>

<!--#include file="style4_imoveis.asp"-->
<head>

<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (doublecombo) 



{




{


if (doublecombo.txt_nome.value == "Seu nome:") {
		alert("Por favor,deixe seu nome na busca , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}

if (doublecombo.txt_nome.value == "") {
		alert("Por favor,deixe seu nome na busca , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}




if (doublecombo.stage22.value == "vqualquer") {
		alert("Por favor, escolha um faixa de valor na busca , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}





if (doublecombo.txt_telefone.value == "Seu telefone:") {
		alert("Por favor, coloque seu telefone , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_telefone.value == "") {
		alert("Por favor, coloque seu telefone , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "Seu email:") {
		alert("Por favor, coloque seu email , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "") {
		alert("Por favor, coloque seu email , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}





var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}



if (doublecombo.combo1.value == "cqualquer") {
		alert("Você precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}



if (doublecombo.example2.value == "nqualquer") {
		alert("Por favor, escolha um tipo de negociação , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.example2.focus();
		
		return false;
}












}
}


</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=520,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>

<script language="javascript">
function funScroll()
{
window.scrollTo(0,300)

}		
</script>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>

</head>












<body onLoad="funScroll()" bgcolor="EAA813" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><form name="doublecombo" onSubmit="return isValidDigitNumber(this);"  method="post" action="listar_imoveis.asp">

<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="755" height="78"><table width="755" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="755" height="51"><a href="primeira.asp"><img src="top_page001.jpg" width="755" height="51" border="0"></a></td>
          </tr>
          <tr>
            <td width="755" height="14" bgcolor="#000000"><div align="center"><table width="600" xmlns=""><tr><td style="width:600; color:#000000;)"><marquee width="100%" scrolldelay="10" scrollamount="2">
                      <font face="Verdana" size="1" color="#FFFFFF"><B>Imobiliária 
                      Veja: Av.Antártico 315 - Jardim do Mar - SBC - CEP 09726-150. 
                      Tel: 4123-72-44. CRECI: 11.676-J. Atuando no mercado imobiliário do grande ABC desde fevereiro de 1991.</B></font>
</marquee></td></tr></table></div></td>
          </tr>
          <tr>
            <td width="755" height="13"><img src="top_page002.jpg" width="755" height="13"></td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td width="755" height="243"><table width="755" height="243" border="0" cellpadding="0" cellspacing="0">
        <tr>
                  <td width="176" height="243" align="center" background="fundo_primeira.jpg" bgcolor="#000000"><div align="center"><table width="149" border="0" cellspacing="0" cellpadding="0" height="170">
                            <tr>
							<td width="149" height="10"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Busca 
                                  de im&oacute;veis</strong> </font> </div></td>
							</tr>
							<tr>
                                  <td height="11"><input type="text" name="txt_nome" class="inputBox" value="<%=session("nome")%>"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;"></td>
                            </tr>
							<tr>
                                  <td height="11"><input type="text" name="txt_telefone" class="inputBox" value="<%=session("telefone")%>"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;"></td>
                            </tr>
							
							<tr>
                                  <td height="11"><input type="text" name="txt_email" class="inputBox" value="<%=session("email")%>"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;"></td>
                            </tr>
							
							
							<tr>
                                  <td>
								  <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 11px; WIDTH: 149px; font-size : 10px; background: FFFFFF; color:000000; ">
                  <option value="cqualquer" selected>Cidade</option>
				  <% if not rs3.eof then %>
                  <% While NOT Rs3.EoF %>
                  <option value="<% = Rs3("id_combo1") %>"<%if session("vCidade2")<> "cqualquer" then%><%if int(rs3("id_combo1")) = int(session("vCidade2")) then response.write "selected" else response.write "" end if %><%end if%>> 
                  <% = Rs3("nome_combo1") %>
                  </option>
                  <% Rs3.MoveNext %>
                  <% Wend %>
				  <option value="cqualquer">qualquer uma</option>
                  <%else%>
                  <option value="cqualquer">qualquer uma</option>
                  <%end if%>
                </select>
								  
								   </td>
                            </tr>
                            <tr>
                                  <td><select name="combo2" class="inputBox" onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px; background: FFFFFF; color:000000;">
                   <option value="bqualquer" selected>Bairro/Região</option>
				  <% if not rs4.eof then%>
                  <% While NOT Rs4.EoF %>
                  <option value="<% = Rs4("id_combo2") %>" <% if session("vBairro2") <> "bqualquer" then if Rs4("id_combo2") = int(session("vBairro2"))  then response.write "selected" else response.write "" end if end if %>> 
                  <% = Rs4("nome_combo2") %>
                  </option>
                  <% Rs4.MoveNext %>
				  
                  <% Wend %>
				   <option value="bqualquer">qualquer um</option>
				  
                  <% else %>
                  <option value="bqualquer">qualquer um</option>
                  <% end if %>
                </select> </td>
                            </tr>
							
							 <tr>
                                  <td><select name="combo5" class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px; background: FFFFFF; color:000000;">
                   <option value="vlqualquer" selected>Vila</option>
				   <% if not rs44.eof and session("vBairro2") <> "bqualquer" then%>
                  <% While NOT Rs44.EoF %>
                  <option value="<% = Rs44("id_combo3") %>" <% if session("vVila2") <> "vlqualquer" then if Rs44("id_combo3") = int(session("vVila2"))  then response.write "selected" else response.write "" end if end if %>> 
                  <% = Rs44("nome_combo3") %>
                  </option>
                  <% Rs44.MoveNext %>
				  
                  <% Wend %>
				   <option value="vlqualquer">qualquer um</option>
				   <%else%>
                  <option value="vlqualquer">qualquer um</option>
                  <%end if%>
				
                </select> </td>
                            </tr>
                            <tr>
                                  <td><select name="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px;background: FFFFFF; color:000000;">
                  <option value="<%=session("vTipo")%>" selected><%if session("vTipo") <> "tqualquer" then  response.write session("vTipo") else response.write "Tipo" end if%></option>
				  <option value="tqualquer">Tipo</option>
				   <option value="tqualquer">Qualquer um</option>
                  	<% if not rs444Tipo22.eof then%>
					<% While NOT rs444Tipo22.EoF %>
                    <option value="<% = rs444Tipo22("tipo") %>">
                    <% =rs444Tipo22("tipo") %>
                    </option>
                    <% rs444Tipo22.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
                  
                 
                </select></td>
                            </tr>
							<tr>
                                  <td><select name="txt_Quartos" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px;background: FFFFFF; color:000000;">
                 <option value="<%=session("vQuartos")%>"><% if session("vQuartos") <> "qqualquer" then response.write session("vQuartos") else response.write "Quartos" end if%></option>
				  <option value="qqualquer">Quartos</option>
				   <option value="qqualquer">Qualquer um</option>
                  <option value="01">01</option>
				   <option value="02">02</option>
				   <option value="03">03</option>
                  <option value="04">04</option>
				  <option value="05">05</option>
                  <option value="06">06</option>
                   <option value="07">07</option>
				  <option value="08">08</option>
                  <option value="09">09</option>
                  
                 
                </select></td>
                            </tr>
							
							<tr>
                                  <td><select name="txt_garagem" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px; background: FFFFFF; color:000000;">
                 <option value="<%=session("vVagas")%>"><% if session("vVagas") <> "gqualquer" then response.write session("vVagas") else response.write "Vagas" end if%></option>
				  <option value="gqualquer">Vagas na Garagem</option>
				   <option value="gqualquer">Qualquer um</option>
                  <option value="01">01</option>
				   <option value="02">02</option>
				   <option value="03">03</option>
                  <option value="04">04</option>
				  <option value="05">05</option>
                  <option value="06">06</option>
                   <option value="07">07</option>
				  <option value="08">08</option>
                  <option value="09">09</option>
                  
                 
                </select></td>
                            </tr>
							
							
							
                            <tr>
                                  <td><select name="example2" size="1" class="inputBox" id="select7" onChange="redirect2(this.options.selectedIndex)" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px; background: FFFFFF; color:000000;">
                
				  <option value="nqualquer">Negociação </option>
                  <option value="nqualquer" >Qualquer um </option>
				  <option  value="Aluguel">Aluguel </option>
                  <option value="Venda">Venda </option>
				   <option value="<%=session("vnegociacao")%>" selected><% if session("vnegociacao") <> "nqualquer" then response.write session("vnegociacao") else response.write "Negociação" end if%></option>
                  
                </select></td>
                            </tr>
                            <tr>
                                  <td><select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 11px; WIDTH: 149px ; font-size : 10px; background: FFFFFF; color:000000;">
                  <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
            <% if session("vnegociacao") = "Aluguel" then %>
			 <option value="<%=session("vValor")%>" selected><% if session("vValor") <> "vqualquer" then response.write FormatNumber(session("vValor1"),2)&" até "&FormatNumber(session("vValor2"),2) else response.write "Valor" end if%></option>
			<option value="0000000000 0000000200">Até 200,00</option>
                  <option value="0000000201 0000000500">201,00 até 500,00</option>
                  <option value="0000000501 0000000750">501,00 até 750,00</option>
                  <option value="0000000751 0000001000">751,00 até 1000,00</option>
                  <option value="0000001001 0000001500">1001,00 até 1500,00</option>
                  <option value="0000001501 0000002000">1501,00 até 2000,00</option>
                  <option value="0000002001 0000002500">2001,00 até 2500,00</option>
                  <option value="0000002501 0000003000">2501,00 até 3000,00</option>
                  <option value="0000003001 0000003500">3001,00 até 3500,00</option>
                  <option value="0000003501 0000004000">3501,00 até 4000,00</option>
                  <option value="0000004001 1000000000">Acima de 4000,00</option>
               <%else%>
			   <option value="<%=session("vValor")%>" selected><% if session("vValor") <> "vqualquer" then response.write FormatNumber(session("vValor1"),2)&" até "&FormatNumber(session("vValor2"),2) else response.write "Valor" end if%></option>
			   <option value="0000000000 0000020000">Até 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 até 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 até 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 até 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 até 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 até 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 até 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 até 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 até 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 até 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
			   <%end if%>
			    </select></td>
                            </tr>
                            <tr>
                              <td><input name="image" type="image"  src="bt_procurar002.jpg" width="149" height="15" border="0"></td>
                            </tr>
                            
                          </table></div></td>
            <td width="579" height="243"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="579" height="243">
                <param name="movie" value="front_page.swf">
                <param name="quality" value="high">
                <embed src="front_page.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="579" height="243"></embed></object></td>
        </tr>
      </table></td>
  </tr>
  <tr>
  <td width="755" height="10" bgcolor="863F15"><table width="755" height="10" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="136"> <div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="quem_somos.asp" style="color:#FFCC00">Quem somos</a></strong></font></div></td>
            <td width="116"> <div align="right"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="onde_estamos.asp" style="color:#FFCC00">Onde 
                      estamos </a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="servicos.asp" style="color:#FFCC00">Servi&ccedil;os</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="financiamento.asp" style="color:#FFCC00">Financiamento/FGTS</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="dicas.asp" style="color:#FFCC00">Dicas</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('form_enviar_email.asp')" style="color:#FFCC00">Contato</a></strong></font></div></td>
          </tr>
        </table></td>
  </tr>
  
</table></form>





<%

%>





<%

dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where (telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%') and atendimento <>'"&"internet"&"' and atendimento <>'"&"não informado"&"' " 
	
	
	
	rs444VerificaConta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta.ActiveConnection = Conexao
	
	
	
	
	
	 rs444VerificaConta.Open strSQL444VerificaConta, Conexao
	

if  not rs444VerificaConta.eof then






vCadastrado = "sim"

%>
<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="568" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
          <td width="708" height="11"><img src="top_display2.jpg" width="708" height="11"></td>
  </tr>
  <tr> 
          <td width="708" height="153">
<table width="708" border="0" cellpadding="0" cellspacing="0" bgcolor="FE9225">
                    <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
                      <td><div align="center">
                          <table width="400" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="5"><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:#FFFFFF;" target="_blank">Ol&aacute; 
                                  sr(a) <%=session("nome")%></a></font><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:#FFFFFF;" target="_blank"><br>
                                  Obrigado por retornar ao nosso site, voc&ecirc; 
                                  est&aacute; conosco desde o dia <%=rs444VerificaConta("data")%> 
                                  e tem sua conta gratuita de comprador de im&oacute;vel.Se 
                                  quiser verificar se novos im&oacute;veis foram 
                                  indicados para o sr(a) clique aqui.Se o 
                                  sr(a) quiser visitar algum dos im&oacute;veis 
                                  abaixo ligue para 4123-72-44 e fale com o seu 
                                  atendente o sr(a) <%=rs444VerificaConta("atendimento")%>.</a></strong></font></div></td>
                            </tr>
                          </table>
                        </div></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
          <td width="708" height="11"><img src="bottom_display2.jpg" width="708" height="11"></td>
  </tr>
</table></td>
  </tr>
</table>

<br>
<%else%>




<%end if%>




  <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2




 
  
   
  
   
   if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")
end if
   
   
   














	  Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	if session("vCidade2") <> "cqualquer" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade2")
 
 
 rs2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs2.ActiveConnection = Conexao
 
 
 
 
 rs2.open SQL2,Conexao,2,1
 
 vCidade = rs2("nome_combo1")
 
 '-----------------------------
           rs2.Close           
		   
           Set rs2 = Nothing
		   
'---------------------------------
 
 else
 vCidade = vCidade2
 end if

	session("vCidade")= vCidade
	
	
	
	
	 
	 if session("vBairro2") <> "bqualquer" then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 from combo2 where id_combo2 ="& session("vBairro2")
 
 
 rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao
 
 
 
 
 
 rs3.open SQL3,Conexao,2,1

 vBairro = rs3("nome_combo2")
 
 '-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------
 
 
 else
 vBairro = vBairro2
	end if                                      
									
	 
	 
	 
	 session("vBairro")= vBairro
	 
	 
	 
	 
	 
	 
	 
	 if session("vvila2") <> "vlqualquer" then
	
	dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 from combo3 where id_combo3 ="&session("vvila2")
 
 rs22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs22.ActiveConnection = Conexao
 
 
 
 
 
 rs22.open SQL22,Conexao,2,1
 
 vvila = rs22("nome_combo3")
 
 '-----------------------------
           rs22.Close           
		   
           Set rs22 = Nothing
		   
'---------------------------------
 
 else
 vvila = vvila2
 end if

	session("vvila")= vvila
	 
	 
	 
	 
	 
	 
	 
	  
	  
	 if session("vCidade")="sao bernardo" then
	  session("vCidade")="São Bernardo"
	  end if
	 
	   if session("vCidade")="santo andre" then
	  session("vCidade")="Santo André"
	  end if
	  
	   if session("vCidade")="sao caetano" then
	  session("vCidade")="São Caetano"
	  end if
	  
	  
	  if session("vBairro")="bairro assuncao" then
	  session("vBairro")="Bairro Assunção"
	  end if
	  
	   if session("vBairro")="ceramica" then
	  session("vBairro")="Cerâmica"
	  end if
	  
	  
	  if session("vBairro")="jd sao caetano" then
	  session("vBairro")="JD São Caetano"
	  end if 
	  'Acima as variáveis recebem os valores dos formulários para fazer busca.
	
'Codificar para receber "qualquer um" vcidade +2 pois vBairro está conectado a vcidade
'----------------------------------------------------------
if session("vCidade") = "" then
session("vCidade") = request.querystring("vCidade")
end if

if session("vBairro") = "" then
session("vBairro") = request.querystring("vBairro")
end if



if session("vTipo") = "" then
session("vTipo") = request.querystring("vTipo")
end if

if session("vNegociacao") = "" then
session("vNegociacao") = request.querystring("vNegociacao")
end if

 '---------------------------------------------------  
 
 
 
 '------------------pegar número de quartos------------------
 
 
 '--------------------------------------------------------
 
 
 '--------------------------Pegar número de Vagas na Garagem-------------
 
 
 
 
 
 
'----------------------------------------------------------------









'----------------------Ocupação do imóvel--------------------------

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 '------------------------------------------------------------------
 
 
 
 
 '----------------------Cidade--------------------------------
 dim stringIndex
 
stringIndex = " where cod_imovel<>"&"0"&"" 

if  session("vCidade") <> "cqualquer" then
stringCidade = "and cidade='"& session("vCidade")&"'"
else
stringCidade = ""
end if
 '--------------------------Bairro----------------------------

if session("vBairro") <> "bqualquer"  then
stringBairro = " and Bairro='"&session("vBairro")&"'"
else
stringBairro = "" 
end if


'--------------------------Vila----------------------------
dim stringVila

if session("vVila") = "" then
session("vVila") = request.QueryString("vVila")
end if

if session("vVila") <> "vlqualquer"  then
stringVila = " and Vila='"&session("vVila")&"'"
else
stringVila = "" 
end if










'------------------------------------------------------------------------------------------

'------------------Tipo-----------------------------

dim StringTipo

if session("vTipo") <> "tqualquer" then
stringTipo = " and Tipo='"&session("vTipo")&"'"
else
stringtipo = ""
end if







'------------------------------------------------------




'----------------------------Vagas na garagem----------------


dim StringVagas

if session("vVagas") <> "gqualquer" then
stringVagas = " and vagas >="&session("vVagas")&""
else
stringVagas = ""
end if

'---------------------------------------------------------------









'---------------------------Ocupação------------------------



dim StringOcupacao

if session("vOcupacao") = "" then
session("vOcupacao") = request.querystring("vOcupacao")
end if

if session("vOcupacao") <> "oqualquer" then
stringOcupacao = " and ocupacao='"&session("vOcupacao")&"'"
else
stringOcupacao = ""
end if








'--------------------------------------------------------------





'-------------------Negociação---------------------------

if  session("vNegociacao") <> "nqualquer" then
stringNegociacao = " and negociacao='"&session("vNegociacao")&"'"
else
stringNegociacao = ""

end if

'------------------------------------------------------------------------------


'---------------------------Quartos------------------------------


if  session("vQuartos") <> "qqualquer" then
stringQuartos = " and quartos >="&int(session("vQuartos"))&""
else
stringQuartos = ""

end if


'-------------------------------------------------------------


'---------------------------------Valor-----------------------------------



dim stringValor



if  session("vValor") <> "vqualquer"  then
stringValor = " and Valor >="& session("vValor1") &" and Valor <="& session("vValor2") &""
else
stringValor = ""
end if







strSQL ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM Imoveis"&stringIndex&stringCidade&stringBairro&stringVila&stringTipo&stringVagas&stringNegociacao&stringQuartos&stringVagas&stringValor&" ORDER BY cod_imovel DESC" 
	


'----------------------------------------------------Fim da instrução SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
' if  vTipo <> ""  then
 
 
 ' Conexao.execute"Insert into imoveis_procurados(nome,telefone,email,Cidade, bairro ,tipo,negociacao,valor,enderecoIP,data,quartos,vagas) values('"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& vCidade &"','"& vBairro &"','"& vTipo &"','"& vNegociacao &"','"& vValor &"','"& EnderecoIP &"','"& vData &"','"&session("vQuartos")&"','"&session("vVagas")&"')"
  
 ' end if
  
  
  '------------------Verifica se o internauta já tem conta---------------------------
  
  dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao
	
	
	
	'----------------------primeira verificação----------------------
'dim rsVerificacao007
'dim strSQLVerificacao007

'dim rsVerificacao008
'dim strSQLVerificacao008

' Set rsVerificacao007 = Server.CreateObject("ADODB.RecordSet")
    
	'strSQLVerificacao007 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.quartos,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.foto_grande1,imoveis.StandBy,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.placa,imoveis.dataLastEmail,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.imovel_em_negociacao,imoveis.origem_captacao  FROM imoveis Where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%'or telefone02 like '%"&session("telefone")&"%'"
	 
   
   
'rsVerificacao007.CursorLocation = 3
'rsVerificacao007.CursorType = 3

       ' rsVerificacao007.Open strSQLVerificacao007, Conexao 


'------------------segunda verificação---------------------------

 Set rsVerificacao008 = Server.CreateObject("ADODB.RecordSet")
    
	strSQLVerificacao008 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores Where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%'or telefone02 like '%"&session("telefone")&"%'"
	 
   
   
rsVerificacao008.CursorLocation = 3
rsVerificacao008.CursorType = 3

        rsVerificacao008.Open strSQLVerificacao008, Conexao 




	
	
	
	
	

if   rs444VerificaConta2.eof and rsVerificacao008.eof  and vTipo <> "" then
 
 dim vValorConta
 dim vQuartosConta
 dim vVagasConta
 dim vValorMedio
 
 if session("vValor") = "vqualquer" then
 vValorMedio = "0"
 else
 vValorMedio = (int(session("vValor1")) + int(session("vValor2")))/2
 end if
 
  
  
  
  if session("vVagas") = "gqualquer" then
 vVagasConta = "0"
 else
 vVagasConta = session("vVagas")
 end if
 
  if session("vQuartos") = "qqualquer" then
 vQuartosConta = "0"
 else
 vQuartosConta = session("vQuartos")
 end if
 

 session("vValorMedio") = vValorMedio
 
 
 session("vQuartosConta") = vQuartosConta
 session("vVagasConta") = vVagasConta
 session("vValorConta") = vValorConta
 
 
 dim vNegociacaoConta
 
 
 
 
 
 
 if session("vNegociacao") = "Venda" then
 vNegociacaoConta = "compra"
 
else
 
 
 vNegociacaoConta = session("vNegociacao")
 end if
 
 
 if session("vNegociacao") = "Aluguel" then
 vNegociacaoConta = "aluguel"
 end if
 
 
 
 session("vNegociacaoConta") = vNegociacaoConta
  
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vQuartosConta") &"','"& session("vNegociacaoConta") &"','"& int(session("vValorMedio")) &"','"& now() &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& session("vVila") &"','"& session("vVagasConta") &"','"& "não informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"')"
	
	
	
	else
	
	
	 While NOT rs444VerificaConta2.EoF 
                      
   if vTipo <> "" then              
		if rs444VerificaConta2("acessos") <> "" then
		
		 
	 'Conexao.execute"update compradores set acessos='"&int(rs444VerificaConta2("acessos"))+1&"' where cod_compradores="&rs444VerificaConta2("cod_compradores")
	 
			else
			
			 	 
	 'Conexao.execute"update compradores set acessos='"&"1"&"' where cod_compradores="&rs444VerificaConta2("cod_compradores")
	 
		end if
	
	end if		   
                    
                      rs444VerificaConta2.MoveNext 
                      Wend 
  
  
  end if
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  

'--------------------------Atualizar acesso-----------------------------

dim rs444VerificaConta002,strSQL444VerificaConta002
   
    Set rs444VerificaConta002 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta002 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	
	rs444VerificaConta002.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta002.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta002.ActiveConnection = Conexao
	
	
	
	
	 rs444VerificaConta002.Open strSQL444VerificaConta002, Conexao
	

if  not rs444VerificaConta002.eof then




	' Conexao.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta002("cod_compradores")
	end if 





'----------------------------------------------------------------------------


  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 



Set RS = Server.CreateObject("ADODB.Recordset")
'um objeto recordset é instânciado.

Dim LinkTemp
'essa variável vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
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
	
RS.Open strSQL, Conn, 1, 3
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

    
	

<table width="708" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  
  

  
  
    <td width="518" height="155">
	
	  
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
<% varCodimovel = rs("COD_Imovel") %>
  
  
  
  
	
	
	
	
	
	
	  <table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="708" height="11"><img src="top_display2.jpg" width="708" height="11"></td>
  </tr>
  <tr> 
          <td width="708" height="153">
<table width="708" border="0" cellpadding="0" cellspacing="0" bgcolor="FE9225">
                    <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
          <td><table width="692" height="153" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                <td width="552" height="16" bgcolor="FE9225"><table width="692" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="130"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
                            <td width="140"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro/Região</strong></font></div></td>
                           <td width="140"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div></td>
						    <td width="99"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
                            <td width="80"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o</strong></font></div></td>
                            <td width="103"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
                          </tr>
                        </table></td>
              </tr>
              <tr> 
                <td width="552" height="16" bgcolor="E17508">
				<table width="692" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="130"><div align="center"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')"><font face="Verdana, arial" size="1" color="white"><%=RS("Cidade")%></font></a></div></td>
                            <td width="140"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Bairro")%></font></div></td>
                            <td width="140"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Vila")%></font></div></td>
						    <td width="99"><div align="center"><font face="Verdana, arial" size="1" color="white"><%if rs("tipo") = "terreno" then response.Write"Terreno/Área" else response.Write RS("Tipo") end if%></font></div></td>
                            <td width="80"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Negociacao")%></font></div></td>
                            <td width="103"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=FormatNumber(RS("Valor"),2)%></font></div></td>
                          </tr>
                        </table>
				
				
				
				</td>
              </tr>
              <tr> 
                <td><table width="692" height="115" border="0" cellpadding="0" cellspacing="0" bgcolor="FE9225">
                                <tr> 
                      <td width="173" bgcolor="FE9225"> 
                        <center>
                                      <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                                        <tr>
                              <td bgcolor="<%=escuro%>" width="160" align="right"><% If objFSO.FileExists(Server.MapPath(rs("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')"><img src="<%=rs("foto_pequena")%>" width="158" height="91" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="color:#FFFFFF"><strong>Foto não disponível</strong></a></font></div><%end if%></td>
                            </tr>
                          </table>
                        </center>
					  
					  </td>
                                  <td bgcolor="FE9225" width="532"><table width="520" height="93" border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td width="431">
										<% if vCadastrado = "sim" then %>
<div align="center"><font face="Verdana, arial" size="1" color="#FFFFFF"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="color:#FFFFFF"><%=RS("obs_imovel")%><br>
                                            <br>
                                            <strong>Veja mais detalhes.<br>Código de referência<%=RS("cod_imovel")%> </strong></a></font></div>
											<%else%>
											
                                          <div align="center"><font face="Verdana, arial" size="1" color="#FFFFFF"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="color:#FFFFFF">Olá 
                                            , meu nome é <strong><%=rs("proprietario")%></strong>, 
                                            o sitema veja analizou os seus dados 
                                            e os dados do meu imóvel, e detectou 
                                            a possibilidade de negócio entre nós. 
                                            Lique já para <strong>4123-72-44</strong> 
                                            e fale com o meu atendente o sr(a) 
                                            <strong>
                                            <%if rs("captacao") = "Spirity" or rs("captacao") = "internet" then response.write "Wanderlei" else response.write rs("captacao") end if%>
                                            </strong>, para que o mesmo agende 
                                            uma visita sua ao meu imóvel, <strong>clique 
                                            aqui</strong> e saiba mais sobre meus 
                                            interesses e condições de pagamento. 
                                            Muito Obrigado</a></font></div>
											<%end if%>
											
											
											
											</td>
                                        
										
										
										
										<td width="89"> 
                                          
										   
										    <%if   (rs("imovel_em_negociacao")= "Vendido pela Veja"  ) then%>
										   <img src="gifvejavendeu.gif" width="89" height="89"></img> 
										   
										   
										   
										   
										    <%elseif (rs("imovel_em_negociacao")= "Vendido por outros") then%>
										    <img src="gifstandby03.gif" width="89" height="89"></img> 
                                         
										    <%elseif (rs("imovel_em_negociacao")= "Suspenso") then%>
										    <img src="gifsuspenso.gif" width="89" height="89"></img> 
										   
										   
										    <% elseif (rs("qualidade")= "bom negócio")  then %>
                                          <img src="gifoferta.gif" width="89" height="89"></img> 
										   
										   
										   
										   
										   
										   <%else%>
										   
										  <%end if%>
                                        </td>
  </tr>
</table>
</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
          <td width="708" height="11"><img src="bottom_display2.jpg" width="708" height="11"></td>
  </tr>
</table>
<br>
	
	
	 <%
RS.MoveNext


	  





 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
		
	
	
	
  </tr>
</table>
	
	</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vValor=<%=session("vValor")%>&vValor1=<%=session("vValor1")%>&vValor2=<%=session("vValor2")%>&vTipo=<%=session("vTipo")%>&vNegociacao=<%=session("vNegociacao")%>&vOcupacao=<%=session("vOcupacao")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
            <strong> 
            <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
            <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
            Página <%=cInt(intPage)%> de <%=cInt(intPageCount)%> </strong></font> 
            <strong><font color="#000000"> 
            <%End If%></font>
            </font></strong></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%>
            <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
            <strong><a href="?page=<%=intPage + 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vValor=<%=session("vValor")%>&vValor1=<%=session("vValor1")%>&vValor2=<%=session("vValor2")%>&vTipo=<%=session("vTipo")%>&vNegociacao=<%=session("vNegociacao")%>&vOcupacao=<%=session("vOcupacao")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="color:#000000">Próximo</a></strong> 
            </a> 
            <%End If%>
            </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>

<table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="100"> 
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
        <br>
        O sr(a) buscou im&oacute;veis para compra e obteve o resultado acima, 
        contate o atendente deste im&oacute;vel em nosso escrit&oacute;rio, para 
        que o mesmo, agende sua visita aos im&oacute;veis que lhe interessam,n&atilde;o 
        havendo interesse por nenhuma das op&ccedil;&otilde;es, a partir de agora, 
        consulte constantemente a p&aacute;gina <a href="contas.asp" style="color:red" target="_blank">veja sua 
        conta cadastro gratuita</a>, pois v&aacute;rios im&oacute;veis s&atilde;o 
        constantemente cadastrados em nosso site, e um destes novos im&oacute;veis 
        pode ser o seu. Obrigado por sua consulta e boa sorte na sua procura.<br>
        <br>
        <br>
        </strong><strong>
		<br>
		<br>
        </strong></font></div></td>
  </tr>
</table>

<%End If


Else

%>



  <table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="250">
<div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="2">O 
        sr(a) buscou im&oacute;veis para compra e n&atilde;o encontrou, isto &eacute; 
        dif&iacute;cil de ocorrer, principalmente se a procura for em S&atilde;o 
        Bernardo do Campo, talvez voc&ecirc; deva corrigir a sua busca sendo menos 
        exigente, querendo mant&ecirc;-la, a partir de agora consulte constantemente 
        a p&aacute;gina <a href="contas.asp" style="color:red" target="_blank">veja sua conta gratuita</a>, pois 
        v&aacute;rios im&oacute;veis s&atilde;o constantemente cadastrados em 
        nosso site, e um destes novos im&oacute;veis, pode ser o seu.Obrigado 
        por sua consulta e boa sorte na sua procura.</font><br>
        <br>
        </strong></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><br><br><br><br>
        </strong></font></div></td>
  </tr>
</table>

  <% end if %>
  
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

'----------------------------------

rs.close
set rs = nothing

'--------------------------------------



'-----------------------------
           rs444VerificaConta.Close           
		   
           Set rs444VerificaConta = Nothing
		   
'---------------------------------


'-----------------------------
           rs444VerificaConta2.Close           
		   
           Set rs444VerificaConta2 = Nothing
		   
'---------------------------------




'-----------------------------
           rs444VerificaConta002.Close           
		   
           Set rs444VerificaConta002 = Nothing
		   
'---------------------------------



'-----------------------------
          ' rs444VerificaConta02.Close           
		   
          ' Set rs444VerificaConta02 = Nothing
		   
'---------------------------------


'-----------------------------
                      
		   
           Set objfso = Nothing
		   
'---------------------------------

'-----------------------------
           rs44.Close           
		   
           Set rs44 = Nothing
		   
'---------------------------------


'-----------------------------
           rs4.Close           
		   
           Set rs4 = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Tipo22.Close           
		   
           Set rs444Tipo22 = Nothing
		   
'---------------------------------










%>






  
  <% response.flush%>
  <%response.clear%>
  

<br>
<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50" bgcolor="<%=escuro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Os 
        dados dispon&iacute;veis neste site s&atilde;o de inteira responsabilidade 
        dos internautas</strong></font></div></td>
  </tr>
</table></td>
  </tr>
</table>
<%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao333 ) %>


<%

conexao.close
set conexao = nothing


conexao3.close
set conexao3 = nothing



conexao333.close
set conexao333 = nothing










%>




<!--#include file="dsn2.asp"-->
</body>
</html>
