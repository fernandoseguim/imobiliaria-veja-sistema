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
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



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
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"

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
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1

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


rsMarcas3.close

set rsMarcas3 = nothing


rsCarros3.close

set rsMarcas3 = nothing




End Function
%> 


<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 




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

dim rs4,strSQL4,Conexao
  
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
	
	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao3
	
	
	
	
	rs4.Open strSQL4, Conexao3



%>


<%
Function EscreveFuncaoJavaScript222 ( Conexao3 )
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
SqlMarcas333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 

Set rsMarcas333 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas333.ActiveConnection = Conexao3
	
	
	rsMarcas333.Open SqlMarcas333, Conexao3



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

rsCarros333.ActiveConnection = Conexao3
	
	
	rsCarros333.Open SqlCarros333, Conexao3
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


rsMarcas333.close

set rsMarcas333 = nothing

rsCarros333.close

set rsCarros333 = nothing



End Function
%> 


<%

'Criando conexão com o banco de dados! 


'

Sql333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 

Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs333.ActiveConnection = Conexao3
	
	
	rs333.Open Sql333, Conexao3




dim vvNome
  dim vvTelefone
  
  vvNome = request.querystring("txt_nome")
  
  if vvNome = "" then
  vvNome = "não informado"
  end if
  
  session("nome") = vvNome
  
  vvTelefone = request.querystring("txt_telefone")
  if vvTelefone = "" then
  vvTelefone = "não informado"
  end if

session("telefone") = vvTelefone


dim vvEmail
  vvEmail = request.querystring("txt_email")
  if vvEmail = "" then
  vvEmail = "não informado"
  end if

session("email") = vvEmail





'--------------------------Atualizar acesso-----------------------------

dim rs444VerificaConta002,strSQL444VerificaConta002
   
    Set rs444VerificaConta002 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta002 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	rs444VerificaConta002.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta002.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta002.ActiveConnection = Conexao3
	
	
	
	 rs444VerificaConta002.Open strSQL444VerificaConta002, Conexao3
	

if  not rs444VerificaConta002.eof then




	 Conexao3.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta002("cod_compradores")
	end if 





'----------------------------------------------------------------------------










'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo FROM tipo  ORDER BY tipo ASC" 
	 
	 rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao3
	 
	 
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao3










'-------------------------------------------------------------------------------------------------



%> 




<html>

<!--#include file="style4_imoveis.asp"-->
<head>
<title>Listas Permutantes</title>


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
   openWindow = window.open(abrejanela,'openWin','width=600,height=510,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>

<script language="javascript">
function funScroll()
{
window.scrollTo(0,321)

}		
</script>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>
</head>












<body onLoad="funScroll()" bgcolor="E17508" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<table width="755" border="0" cellspacing="0" cellpadding="0" bgcolor="EAA813">
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
                  <td width="176" height="243" background="fundo_primeira.jpg" align="center" bgcolor="#000000"> 
                    <div align="center"><table width="149" border="0"  cellspacing="0" cellpadding="0" height="170">
                            <tr>
							<td width="149" height="10"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Busca 
                                  de im&oacute;veis</strong> </font> </div></td>
							</tr>
							<tr>
                                  <td height="11"><input name="txt_nome" onfocus="doublecombo.txt_nome.value=''"  type="text" class="inputBox" id="txt_nome"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;" value="Seu nome:"></td>
                            </tr>
							<tr>
                                  <td><input name="txt_telefone" onfocus="doublecombo.txt_telefone.value=''" type="text" class="inputBox" id="txt_telefone"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;" value="Seu telefone:"></td>
                            </tr>
							<tr>
                                  <td><input name="txt_email" onfocus="doublecombo.txt_email.value=''" type="text" class="inputBox" id="txt_email"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;" value="Seu email:"></td>
                            </tr>
							
							
							<tr>
                                  <td>
								  <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 11px; WIDTH: 149px; background:white;color:black;">
                  <option value="cqualquer" selected>Cidade</option>
				  <% if not rs3.eof then %>
                  <% While NOT Rs3.EoF %>
                  <option value="<% = Rs3("id_combo1") %>" > 
                  <% = Rs3("nome_combo1") %>
                  </option>
                  <% Rs3.MoveNext %>
                  <% Wend %>
				  <option value="cqualquer">qualquer uma</option>
                  <%else%>
                  <option value=""></option>
                  <%end if%>
                </select>
								  
								   </td>
                            </tr>
                            <tr>
                                  <td><select name="combo2" class="inputBox" onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 11px; WIDTH: 149px; background:white;color:black;">
                   <option value="bqualquer" selected>Bairro/Região</option>
				  <% if not rs4.eof then%>
                  <% While NOT Rs4.EoF %>
                  <option value="<% = Rs4("id_combo2") %>"> 
                  <% = Rs4("nome_combo2") %>
                  </option>
                  <% Rs4.MoveNext %>
				  
                  <% Wend %>
				   <option value="bqualquer">qualquer um</option>
				  
                  <% else %>
                  <option value=""></option>
                  <% end if %>
                </select> </td>
                            </tr>
							
							 <tr>
                                  <td><select name="combo5" class="inputBox" style="HEIGHT: 11px; WIDTH: 149px; background:white;color:black;">
                   <option value="vlqualquer" selected>Vila</option>
				 <option value="vlqualquer">qualquer um</option>
                </select> </td>
                            </tr>
                            <tr>
                                  <td><select name="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
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
                                  <td><select name="txt_Quartos" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ;  background:#FFFFFF; color:#000000;">
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
                                  <td><select name="txt_garagem" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ;  background:#FFFFFF; color:#000000;">
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
                                  <td><select name="example2" size="1" class="inputBox" id="select7" onChange="redirect2(this.options.selectedIndex)" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="nqualquer">Negociação </option>
                  <option value="nqualquer" >Qualquer um </option>
				  <option  value="Aluguel">Aluguel </option>
                  <option value="Venda">Venda </option>
                  
                </select></td>
                            </tr>
                            <tr>
                                  <td><select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
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

      <br>
	  
	 
	  

      <br>
	  <%

dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais FROM permuta where telefone ='"&session("telefone")&"' and atendimento <>'"&"internet"&"' and atendimento <>'"&"não informado"&"' " 
	
	
	rs444VerificaConta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta.ActiveConnection = Conexao3
	
	
	
	
	
	
	 rs444VerificaConta.Open strSQL444VerificaConta, Conexao3
	

if  not rs444VerificaConta.eof then
%>
<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="518" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="708" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
                <td width="568" height="153"> 
                  <table width="568" border="0" cellpadding="0" cellspacing="0" bgcolor="FE9225">
                    <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
                      <td><div align="center">
                          <table width="400" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="5"><a href="acessoLink01.asp?varTelefone=<%=session("telefone")%>" style="color:#FFFFFF;" target="_blank">Ol&aacute; 
                                  <%=session("nome")%></a></font><a href="acessoLink01.asp?varTelefone=<%=session("telefone")%>" style="color:#FFFFFF;" target="_blank"><br>
                                  voc&ecirc; est&aacute; cadastrado no nosso sistema 
                                  e tem uma conta gratuita de permutante de im&oacute;veis, 
                                  se quiser verificar se novos permutantes foram 
                                  indicados para voc&ecirc;, clique aqui.</a></strong></font></div></td>
                            </tr>
                          </table>
                        </div></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
        </tr>
      </table>
                </td>
  </tr>
  <tr> 
                <td width="708" height="11"><img src="bottom_display2.jpg" width="568" height="11"></td>
  </tr>
</table></td>
  </tr>
</table>

<br>
<%else%>




<%end if%>

	  <br>
	  
	  
<center>
  <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2
dim vValor_vend,vValor_vend1,vValor_vend2
dim vValor_comp,vValor_comp1,vValor_comp2
dim vCidade_vend,vCidade_comp
dim stringIndex
 
  vValor_vend=request.querystring("txt_valor_vend")
  
  session("vValor_vend") = vValor_vend
  
  session("vValor_vend1")=left(vValor_vend,10)
   session("vValor_vend2")=right(vValor_vend,10)
 
  
  
  vValor_comp=request.querystring("txt_valor_comp")
  
  if vValor_comp = "vqualquer" then
  vValor_comp = "0000000000 0000000000"
  end if
  
  
  session("vValor_comp") = vValor_comp
  
   session("vValor_comp1")=left(vValor_comp,10)
   session("vValor_comp2")=right(vValor_comp,10)
   
   dim vNome,vTelefone
   
   vNome = request.querystring("txt_nome")
   vTelefone = request.querystring("txt_telefone")
   
  
 
  
  
 
 '---------------------------Buacar Cidades-------------------------------------
 

 
 vCidade_vend2 = request.querystring("combo1")
 
session("vCidade_vend2") = vCidade_vend2
  
   
   if session("vCidade_vend2") = "" then
session("vCidade_vend2") = request.querystring("vCidade_vend2")
end if
   
   
    
	
	
	
	
	 

	  
	
	if session("vCidade_vend2") = "" then
	session("vCidade_vend2") = request.QueryString("vCidade_vend2")
	end if
	
	
	if session("vCidade_vend2") <> "cqualquer" then
	
	dim rs222,SQL222
 Set rs222 = Server.CreateObject("ADODB.RecordSet")
 SQL222 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade_vend2")
 
 rs222.open SQL222,Conexao3,2,1
 
 vCidade_vend = rs222("nome_combo1")
 
 rs222.close
 
 set rs222 = nothing
 
 else
 vCidade_vend = vCidade_vend2
 end if

	session("vCidade_vend")= vCidade_vend
	
	if session("vCidade_vend") = "" then
	session("vCidade_vend") = request.querystring("vCidade_vend")
	end if
	
	
	
	dim vBairro_vend2
	 vBairro_vend2=request.Querystring("combo2")
	 session("vBairro_vend2") = vBairro_vend2
	 if session("vBairro_vend2") = "" then
session("vBairro_vend2") = request.querystring("vBairro_vend2")

end if
	 
	 if session("vBairro_vend2") <> "bqualquer" then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& session("vBairro_vend2")
 
 rs3.open SQL3,Conexao3,2,1

 vBairro_vend = rs3("nome_combo2")
 
 rs3.close
 
 set rs3 = nothing
 
 
 
 
 else
 vBairro_vend = vBairro_vend2
	end if                                      
									
	 
	 
	 
	 session("vBairro_vend")= vBairro_vend
	 
	 if session("vBairro_vend") = "" then
	session("vBairro_vend") = request.querystring("vBairro_vend")
	end if
	 
	 
	 
	 
	dim vVila_vend2
	 vVila_vend2=request.Querystring("combo5")
	 session("vVila_vend2") = vVila_vend2
	 
	 if session("vVila_vend2") = "" then
session("vVila_vend2") = request.querystring("vVila_vend2")

end if
	 
	 if session("vVila_vend2") <> "vlqualquer" then
	  dim rs88,SQL88
 Set rs88 = Server.CreateObject("ADODB.RecordSet")
 SQL88 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="& session("vVila_vend2")
 
 rs88.open SQL88,Conexao3,2,1

 vVila_vend = rs88("nome_combo3")
 
 rs88.close
 
 set rs88 = nothing
 
 else
 vVila_vend = vVila_vend2
	end if                                      
									
	 
	 
	 
	 session("vVila_vend")= vVila_vend


if session("vVila_vend") = "" then
	session("vVila_vend") = request.querystring("vVila_vend")
	end if
 '----------------------------Cidade e bairro comp-------------------------
 
 
 
 
  
  dim vCidade_comp2
 
   
   vCidade_comp2=request.querystring("combo3")
   
   session("vCidade_comp2") = vCidade_comp2
   
   if session("vCidade_comp2") = "" then
session("vCidade_comp2") = request.querystring("vCidade_comp2")
end if
   
   
    
	
	
	
	
	 

	 
	
	if session("vCidade_comp2") <> "cqualquer" then
	
	dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade_comp2")
 
 rs22.open SQL22,Conexao3,2,1
 
 vCidade_comp = rs22("nome_combo1")
 
 rs22.close
 
 set rs22 = nothing
 
 else
 vCidade_comp = vCidade_comp2
 end if

	session("vCidade_comp")= vCidade_comp
	
	if session("vCidade_comp") = "" then
	session("vCidade_comp") = request.querystring("vCidade_comp")
	end if
	
	
	dim vBairro_comp2
	 vBairro_comp2=request.Querystring("combo4")
	 session("vBairro_comp2") = vBairro_comp2
	 if session("vBairro_comp2") = "" then
session("vBairro_comp2") = request.querystring("vBairro_comp2")

end if
	 
	 if session("vBairro_comp2") <> "bqualquer" then
	  dim rs33,SQL33
 Set rs33 = Server.CreateObject("ADODB.RecordSet")
 SQL33 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   from combo2 where id_combo2 ="& session("vBairro_comp2")
 
 rs33.open SQL33,Conexao3,2,1

 vBairro_comp = rs33("nome_combo2")
 
 rs33.close
 
 set rs33 = nothing
 
 else
 vBairro_comp = vBairro_comp2
	end if                                      
									
	 
	 
	 
	 session("vBairro_comp")= vBairro_comp
	 
	 
	 if session("vBairro_comp") = "" then
	session("vBairro_comp") = request.querystring("vBairro_comp")
	end if
	 
	 dim vVila_comp2
	 vVila_comp2=request.Querystring("combo7")
	 session("vVila_comp2") = vVila_comp2
	 if session("vVila_comp2") = "" then
session("vVila_comp2") = request.querystring("vVila_comp2")

end if
	 
	 if session("vVila_comp2") <> "vlqualquer" then
	  dim rs99,SQL99
 Set rs99 = Server.CreateObject("ADODB.RecordSet")
 SQL99 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="& session("vVila_comp2")
 
 rs99.open SQL99,Conexao3,2,1

 vVila_comp = rs99("nome_combo3")
 
 
 rs99.close
 
 set rs99 = nothing
 
 
 else
 vVila_comp = vVila_comp2
	end if                                      
									
	 
	 
	 
	 session("vVila_comp")= vVila_comp

	 if session("vVila_comp") = "" then
	session("vVila_comp") = request.querystring("vVila_comp")
	end if
	 
	 
	 

 '-----------------------------Buscando os tipos de imóveis---------------------
 
 
 
 dim vTipo_vend,vTipo_comp
 
 vTipo_vend = request.Querystring("txt_Tipo_vend")
 session("vTipo_vend") = vTipo_vend
 
 if session("vTipo_vend") = "" then
 
 session("vTipo_vend") = request.querystring("vTipo_vend")
 
 end if
 
 
 
 
 vTipo_comp = request.querystring("txt_Tipo_comp")
 session("vTipo_comp") = vTipo_comp
 
 if session("vTipo_comp") = "" then
 
 session("vTipo_comp") = request.querystring("vTipo_comp")
 
 end if
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 '------------------------------Números de quartos--------------------------------

 
 
 dim vQuartos_vend,vQuartos_comp
 
 vQuartos_vend = request.querystring("txt_Quartos_vend")
 session("vQuartos_vend") = vQuartos_vend
 
 if session("vQuartos_vend") = "" then
 
 session("vQuartos_vend") = request.querystring("vQuartos_vend")
 
 end if
 
 
 
 
 vQuartos_comp = request.querystring("txt_Quartos_comp")
 session("vQuartos_comp") = vQuartos_comp
 
 if session("vQuartos_comp") = "" then
 
 session("vQuartos_comp") = request.querystring("vQuartos_comp")
 
 end if
 
 
 
 
 '------------------------------Números de Vagas--------------------------------

 
 
 dim vVagas_vend,vVagas_comp
 
 vVagas_vend = request.querystring("txt_vagas_vend")
 session("vVagas_vend") = vVagas_vend
 
 if session("vVagas_vend") = "" then
 
 session("vVagas_vend") = request.querystring("vVagas_vend")
 
 end if
 
 
 
 
 vVagas_comp = request.querystring("txt_vagas_comp")
 session("vVagas_comp") = vVagas_comp
 
 if session("vVagas_comp") = "" then
 
 session("vVagas_comp") = request.querystring("vVagas_comp")
 
 end if
 
 
 
 
 
 
 
 
 
 
 
 '------------------------Sua Cidade--------------------------

stringIndex = " where cod_permuta<>"&"0"&"" 

if  session("vCidade_vend") <> "cqualquer" and session("vCidade_vend") <> "qualquer um" and session("vCidade_vend") <> "não informado" then
stringCidadeVend = " and (cidade_comp='"& session("vCidade_vend")&"' or cidade_comp='"& "não informado" &"' or cidade_comp='"& "cqualquer" &"' or cidade_comp='"&"qualquer um"&"')"
else
stringCidadeVend = ""
end if
 
 
 
 	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend

 if   session("vBairro_vend") <> "bqualquer" and session("vBairro_vend") <> "não informado" and session("vBairro_vend") <> "qualquer um" then
	stringBairroVend = " and (Bairro_comp like '%"&session("vBairro_vend")&"%' or Bairro_comp like '%"&"não informado"&"%' or Bairro_comp like'%"&"bqualquer"&"%' or Bairro_comp like'%"&"qualquer um"&"%')"
 else

stringBairroVend = ""

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend

 if   session("vVila_vend") <> "vlqualquer" and session("vVila_vend") <> "não informado" and session("vVila_vend") <> "qualquer um" then
	stringVilaVend = " and (Vila_comp='"&session("vVila_vend")&"' or Vila_comp='"&"não informado"&"' or Vila_comp='"&"qualquer um"&"')"
 else

stringVilaVend = ""

end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend
 
 
 if session("vTipo_vend") <> "tqualquer" then

stringTipoVend = " and Tipo_comp like '%"&session("vTipo_vend")&"%'"

else
stringTipoVend = ""
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend
 
 
 if session("vQuartos_vend") <> "qqualquer" then

stringQuartosVend = " and Quartos_comp <="&session("vQuartos_vend")&""
else
stringQuartosVend = ""
 end if
 


 
 
 '-----------------------Número de Vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend
 
 
 if session("vVagas_vend") <> "vgqualquer" then

stringVagasVend = " and vagas_comp <="&session("vVagas_vend")&""
else

stringVagasVend = ""
 end if
 


 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
	 dim stringValorVend
	
	
	 if session("vValor_vend") = "" then
	 session("vValor_vend")= request.QueryString("vValor_vend")
	 end if
	 
	  
	   if session("vValor_vend1") = "" then
	 session("vValor_vend1")= request.QueryString("vValor_vend1")
	 end if
	 
	  if session("vValor_vend2") = "" then
	 session("vValor_vend2")= request.QueryString("vValor_vend2")
	 end if
	 
  if session("vValor_vend")<>"vqualquer" then
	stringValorVend = " and Valor_comp >="& session("vValor_vend1") &" and Valor_comp <="& session("vValor_vend2") &""
	else	
	stringValorVend = ""
  end if
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp
  if session("vCidade_comp")<>"cqualquer" and session("vCidade_comp")<>"não informado" and session("vCidade_comp")<>"qualquer um" then
	stringCidadeComp = " and Cidade_vend ='"& session("vCidade_comp") &"'"
	else
	
	stringCidadeComp = ""
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp

	if session("vBairro_comp") <> "bqualquer" and session("vBairro_comp") <> "não informado" and session("vBairro_comp") <> "qualquer um" then
	stringBairroComp = " and Bairro_vend ='"& session("vBairro_comp") &"'"
	else
	
	stringBairroComp = ""
	end if
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 
	 dim stringVilaComp

	if session("vVila_comp") <> "vlqualquer" and session("vVila_comp") <> "qualquer um" and session("vVila_comp") <> "não informado" then
	stringVilaComp = " and Vila_vend ='"& session("vVila_comp") &"'"
	else
	
	stringVilaComp = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	 dim stringTipoComp
  if session("vTipo_comp")<>"tqualquer" then
	stringTipoComp = " and Tipo_vend ='"& session("vTipo_comp") &"'"
	else
	
	
	stringTipoComp = ""
	end if
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp
  if session("vQuartos_comp")<>"qqualquer" then
	stringQuartosComp = " and Quartos_vend >="& session("vQuartos_comp") &""
	else
	
	stringQuartosComp = ""
	end if
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp
  if session("vVagas_comp") <> "vgqualquer" then
	stringVagasComp = " and vagas_vend >="& session("vVagas_comp") &""
	else	
	stringVagasComp = ""
	end if
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 '----------------------------Valor pretendido----------------------------



	 dim stringValorComp
	 
	 
	 if session("vValor_comp") = "" then
	 session("vValor_comp")= request.QueryString("vValor_comp")
	 end if
	 
	 if session("vValor_comp1") = "" then
	 session("vValor_comp1")= request.QueryString("vValor_comp1")
	 end if
	 
	 if session("vValor_comp2") = "" then
	 session("vValor_comp2")= request.QueryString("vValor_comp2")
	 end if
	 
  if session("vValor_comp")<>"vqualquer" then
	stringValorComp = " and Valor_vend >="& session("vValor_comp1") &" and Valor_vend <="& session("vValor_comp2") &""
	else
	
	
	stringValorComp = ""
	end if
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringVilaVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringVilaComp&stringTipoComp&stringQuartosComp&stringVagasComp&stringValorComp
	
	
	
	
	
	
	
	
	
	
	
	if vNome = "" then
	vNome = "não informado"
	end if
	
	if vTelefone = "" then
	vTelefone = "não informado"
	end if
	
	
	 dim vEnderecoIP , vData
  vData = now()
  
 
 vEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	
	
	
dim rs444VerificaConta2,strSQL444VerificaConta2
 dim rs444VerificaConta3,strSQL444VerificaConta3
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where telefone like '%"&session("telefone")&"%' or telefone02 like'%"&session("telefone")&"%' or telefone03 like'%"&session("telefone")&"%'" 
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao3
	
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao3	
	
	
	 Set rs444VerificaConta3 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta3 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou FROM compradores where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	 rs444VerificaConta3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

 rs444VerificaConta3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

 rs444VerificaConta3.ActiveConnection = Conexao3
	
	
	
	
	
	 rs444VerificaConta3.Open strSQL444VerificaConta3, Conexao3	
	
	
	
	
	
	
	if rs444VerificaConta2.eof and vTipo <> "" and rs444VerificaConta3.eof then
	
	'Conexao.execute"Insert into permuta_procurados(Nome,Telefone,cidade_vend, bairro_vend,tipo_vend,quartos_vend,valor_vend,cidade_comp,bairro_comp,tipo_comp,quartos_comp,valor_comp,enderecoIP,data,vagas_comp,vagas_vend)values( '"& vNome &"','"& vTelefone &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& vQuartos_vend &"','"& vValor_vend &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& vValor_comp &"','"& vEnderecoIP &"','"& vData &"','"& session("vVagas_comp") &"','"& session("vVagas_vend") &"')" 
	
	
	 if session("vVagas_comp") <> "vgqualquer" then
	session("vVagas_comp") = session("vVagas_comp")
	else	
	session("vVagas_comp") = "0"
	end if
	
	if vQuartos_comp <> "qqualquer" then
	vQuartos_comp = vQuartos_comp
	else
	vQuartos_comp = "0"
	end if
	
	if vValor_comp <> "vqualquer" then
	vValor_comp = vValor_comp
	else
	vValor_comp = "0"
	end if
	
	
	
	
	 if vVagas_vend <> "vgqualquer" then
	vVagas_vend = vVagas_vend
	else	
	vVagas_vend = "0"
	end if
	
	if vQuartos_vend <> "qqualquer" then
	vQuartos_vend = vQuartos_vend
	else
	vQuartos_vend = "0"
	end if
	
	if vValor_vend <> "vqualquer" then
	vValor_vend = vValor_vend
	else
	vValor_vend = "0"
	end if
	
	
	
	if vCidade_comp <> "cqualquer" then
	vCidade_comp = vCidade_comp
	else
	vCidade_comp = "não informado"
	end if
	
	
	
	if vBairro_comp <> "bqualquer" then
	vBairro_comp = vCidade_comp
	else
	vBairro_comp = "não informado"
	end if
	
	dim varValorMedioComp
	
	varValorMedioComp= (int(session("vValor_comp1")) + int(session("vValor_comp2")))/2
	
	
	dim varValorMedioVend
	
	varValorMedioVend= (int(session("vValor_vend1")) + int(session("vValor_vend2")))/2
	
	
	
	
	'Conexao3.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& int(varValorMedioComp) &"','"& now() &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& "não informado" &"','"& session("vVagas_comp") &"','"& "não informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"')"
	
	
				
'Conexao3.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade,indexador_indicacoes) values( '"& session("nome") &"','"& "não informado" &"','"& session("telefone") &"','"& session("email") &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "icon_foto2.gif" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "0" &"','"& "0" &"','"& vQuartos_vend &"','"& "não informado" &"','"& vVagas_vend&"','"& "venda" &"','"& int(varValorMedioVend) &"','"& now() &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& "não informado" &"','"& "Sem Placa"&"','"& "0" &"','"& "0" &"','"& "0" &"','"& "negócio comum" &"','"&"0"&"')"	 

	
	
	
  end if
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
   dim rs444VerificaConta02,strSQL444VerificaConta02
   
    Set rs444VerificaConta02 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta02 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone ='"&session("telefone")&"'" 
	
	
	rs444VerificaConta02.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta02.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta02.ActiveConnection = Conexao3
	
	
	
	 rs444VerificaConta02.Open strSQL444VerificaConta02, Conexao3
	

if   rs444VerificaConta02.eof and vTipo_vend <> "" then
 
 dim vValorConta
 dim vQuartosConta
 dim vVagasConta
 dim vValorMedio
 
 if session("vValor") = "vqualquer" then
 vValorMedio = "0"
 else
 vValorMedio = int(session("vValor1")) + int(session("vValor2"))/2
 end if
 
  
  
  
  if session("vVagas") = "vgqualquer" then
 vVagasConta = "0"
 else
 vVagasConta = session("vVagas")
 end if
 '-----------------quartos-------------------------
  
  dim vQuartosConta_comp
  dim vQuartosConta_vend
  
  if session("vQuartos_comp") = "qqualquer" then
 vQuartosConta_comp = "0"
 else
vQuartosConta_comp = session("vQuartos_comp")
 end if
 
 
  if session("vQuartos_vend") = "qqualquer" then
 vQuartosConta_vend = "0"
 else
vQuartosConta_vend = session("vQuartos_vend")
 end if
 
 
 
 
 
 '------------------------------------
 
 
 '-----------------Vagas-------------------------
  
  dim vVagasConta_comp
  dim vVagasConta_vend
  
  if session("vVagas_comp") = "vgqualquer" then
 vVagasConta_comp = "0"
 else
vVagasConta_comp = session("vVagas_comp")
 end if
 
 
  if session("vVagas_vend") = "vgqualquer" then
 vVagasConta_vend = "0"
 else
vVagasConta_vend = session("vVagas_vend")
 end if
 
  
 
 '------------------------------------

 '-------------------Valor-----------------------------
 
 dim vValorMedio_vend
 dim vValorMedio_comp
 
 
 if session("vValor_vend") = "vqualquer" then
 vValorMedio_vend = "0"
 else
 vValorMedio_vend = (int(session("vValor_vend1")) + int(session("vValor_vend2")))/2
 end if
 
 
  
 if session("vValor_comp") = "vqualquer" then
 vValorMedio_comp = "0"
 else
 vValorMedio_comp = (int(session("vValor_comp1")) + int(session("vValor_comp2")))/2
 end if
 
 session("vValorMedio_comp") = vValorMedio_comp
 
 

 session("vValorMedio") = vValorMedio
 
 
 session("vQuartosConta") = vQuartosConta
 session("vVagasConta") = vVagasConta
 session("vValorConta") = vValorConta
 
 
 dim vNegociacaoConta
 
 if session("vNegociacao") = "compra" or session("vNegociacao") = "Compra" then
 vNegociacaoConta = "venda"
 else
 vNegociacaoConta = session("vNegociacao")
 end if
 
 session("vNegociacaoConta") = vNegociacaoConta
  
	'Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& "não informado" &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vQuartosConta") &"','"& session("vNegociacaoConta") &"','"& session("vValorMedio") &"','"& now() &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& session("vVila") &"','"& session("vVagasConta") &"','"& session("vOcupacao") &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"')"
	
	if vVila_vend = "vlqualquer" then
	vVila_vend = "não informado"
	end if
	
		if vVila_comp = "vlqualquer" then
	vVila_comp = "não informado"
	end if	
    	
	'Conexao.execute"Insert into permuta(Foto_imovel,Nome,Email,Telefone,endereco_vend,cidade_vend,bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby)values( '"& "imovel00000.jpg" &"','"& session("nome") &"','"& session("email") &"','"& session("telefone") &"','"& "não informado" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "não informado" &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& "não informado" &"','"& "0" &"','"& "não informado" &"','"& now() &"','"& vQuartosConta_vend &"','"& vQuartosConta_comp &"','"& vValorMedio_vend &"','"& vValorMedio_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"&vVila_comp&"','"& vVagasConta_vend &"','"& vVagasConta_comp &"','"& "excluido" &"')" 
	
	
	
	
	
	
	else
	
	
  
  
  end if
  
  '------------------------------------------------------
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 
if vTipo_vend <> "" then
	
	Conexao3.execute"Insert into permuta_procurados(Nome,Telefone,cidade_vend, bairro_vend,tipo_vend,quartos_vend,valor_vend,cidade_comp,bairro_comp,tipo_comp,quartos_comp,valor_comp,enderecoIP,data,vagas_comp,vagas_vend)values( '"& vNome &"','"& vTelefone &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& vQuartos_vend &"','"& vValor_vend &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& vValor_comp &"','"& vEnderecoIP &"','"& vData &"','"& session("vVagas_comp") &"','"& session("vVagas_vend") &"')" 
	
  end if


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


  <table width="537" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="537" height="36"><table width="537" border="0" cellspacing="0" cellpadding="0">
        
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

dim varCodPermuta



dim Conexao2,rs7
 
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL7
	
	
	


 dim Conexao22,rs77

	Set rs77 = Server.CreateObject("ADODB.RecordSet")

	dim strSQL77
	
	
	dim vimagem
	if rs("cod_imovel") <> "não informado" then
	 strSQL77 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou FROM imoveis where cod_imovel="&rs("cod_imovel")
	
	 rs77.CursorLocation = 3
      rs77.CursorType = 3
	 rs77.Open strSQL77, Conexao3
	 
	 
	 
   if not rs77.eof then
   vimagem = rs77("foto_grande")
   else
   vimagem = "imovel00000.jpg"
  end if
	
	else
	vimagem = "imovel00000.jpg"
	end if
	
	
	 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	


%>
<% varCodPermuta =RS("cod_permuta") %>
 <tr>
            <td><table bgcolor="FE9225" width="568" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="568" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
    <td width="568" height="153"><table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
                              <td width="8" height="153"><img src="left_display2.jpg" width="8" height="173"></td>
          <td><table width="552" height="153" border="0" cellpadding="0" cellspacing="0" bgcolor="FE9225">
                            <tr> 
                                    <td width="552" height="30" bgcolor="FE9225">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Estou 
                                        interessado em im&oacute;vel na cidade 
                                        de<strong> <a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><font face="Verdana, arial" size="1" color="white"><%=RS("Cidade_comp")%></font></a> 
                                        </strong><strong> </strong></font><font face="Verdana, arial" size="1" color="white"><strong> 
                                        <%if rs("bairro_comp") = "bqualquer" or rs("bairro_comp")= "não informado"  then response.Write "" else response.write "no bairro de "&RS("Bairro_comp") end if %>
                                        </strong></font></div></td>
              </tr>
              <tr> 
                              <td width="552" height="16" bgcolor="E17508"><div align="center"><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')">
                                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
                                    mais detalhes</strong></font></div>
                                  </a></div></td>
              </tr>
              <tr> 
                <td><table width="552" height="115" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="173" bgcolor="FE9225"> 
                        <center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td bgcolor="<%=escuro%>"><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><%If objFSO.FileExists(Server.MapPath(vimagem)) = True Then%><img src="<%=vimagem%>" width="158" height="90" border=0></img><% else %><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Foto não disponível</strong></font></div><% end if %></a></td>
                            </tr>
                          </table>
                        </center>
					  
					  </td>
                      <td bgcolor="FE9225"><div align="center"><font face="Verdana, arial" size="1" color="FFFFFF"> 
                                              <a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')" style="color:#FFFFFF">Olá, 
                                              meu nome &eacute;<strong> <%=rs("nome")%></strong> 
                                              ,o sitema VEJA analizou os dados 
                                              do seu e do meu imóvel e dectetou 
                                              a possibilidade de efetuarmos uma 
                                              permuta entre nossos imóveis. Lique 
                                              já para <strong>4123-72-44</strong> 
                                              e fale com meu atendente sr(a) <strong><%=rs("atendimento")%></strong>. 
                                              para que cada um de nós visitemos 
                                              os imóveis de um e de outro, para 
                                              ver mais detalhes do meu imóvel 
                                              clique aqui, muito obrigado.</a> 
                                              </font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
                              <td width="8" height="153"><img src="right_display2.jpg" width="8" height="173"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="568" height="11"><img src="bottom_display2.jpg" width="568" height="11"></td>
  </tr>
</table></td>
 
 </tr>
 <tr>
          <td height="18"> </td>
 </tr>
 
       
		
		 <%
RS.MoveNext


	  




If colorchanger = 1 Then
	colorchanger = 0
	color1 = "#537497"
	color2 = "#94ADC8"
Else
	colorchanger = 1
	color1 = "#94ADC8"
	color2 = "#537497"
End If

if corfonte = "black" then
 corfonte = "white"
 
 else
 
 corfonte = "white"
 end if
 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
		
		
		
		
      </table></td>
  </tr>
  
  <tr>
    <td width="537" height="18"><table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><font color="#000000" size="1" face="Verdana, arial"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
               <a href="?page=<%=intPage - 1%>&vCidade_vend=<%=session("vCidade_vend")%>&vCidade_vend2=<%=session("vCidade_vend2")%>&vBairro_vend=<%=session("vBairro_vend")%>&vBairro_vend2=<%=session("vBairro_vend2")%>&vVila_vend=<%=session("vVila_vend")%>&vVila_vend2=<%=session("vVila_vend2")%>&vTipo_vend=<%=session("vTipo_vend")%>&vNegociacao_vend=<%=session("vNegociacao_vend")%>&vQuartos_vend=<%=session("vQuartos_vend")%>&vVagas_vend=<%=session("vVagas_vend")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_vend1=<%=session("vValor_vend1")%>&vValor_vend2=<%=session("vValor_vend2")%>&vCidade_comp=<%=session("vCidade_comp")%>&vCidade_comp2=<%=session("vCidade_comp2")%>&vBairro_comp=<%=session("vBairro_comp")%>&vBairro_comp2=<%=session("vBairro_comp2")%>&vVila_comp=<%=session("vVila_comp")%>&vVila_comp2=<%=session("vVila_comp2")%>&vTipo_comp=<%=session("vTipo_comp")%>&vNegociacao_comp=<%=session("vNegociacao_comp")%>&vQuartos_comp=<%=session("vQuartos_comp")%>&vVagas_comp=<%=session("vVagas_comp")%>&vValor_comp1=<%=session("vValor_comp1")%>&vValor_comp2=<%=session("vValor_comp2")%>&vValor_comp=<%=session("vValor_comp")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>" style="color:#000000">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td>
              
			  <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
			  <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
        Página <%=cInt(intPage)%> de 
        <%=cInt(intPageCount)%> </font> 
        <%End If%></font>
        </div>
             
             </td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
                <a href="?page=<%=intPage + 1%>&vCidade_vend=<%=session("vCidade_vend")%>&vCidade_vend2=<%=session("vCidade_vend2")%>&vBairro_vend=<%=session("vBairro_vend")%>&vBairro_vend2=<%=session("vBairro_vend2")%>&vVila_vend=<%=session("vVila_vend")%>&vVila_vend2=<%=session("vVila_vend2")%>&vTipo_vend=<%=session("vTipo_vend")%>&vNegociacao_vend=<%=session("vNegociacao_vend")%>&vQuartos_vend=<%=session("vQuartos_vend")%>&vVagas_vend=<%=session("vVagas_vend")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_vend1=<%=session("vValor_vend1")%>&vValor_vend2=<%=session("vValor_vend2")%>&vCidade_comp=<%=session("vCidade_comp")%>&vCidade_comp2=<%=session("vCidade_comp2")%>&vBairro_comp=<%=session("vBairro_comp")%>&vBairro_comp2=<%=session("vBairro_comp2")%>&vVila_comp=<%=session("vVila_comp")%>&vVila_comp2=<%=session("vVila_comp2")%>&vTipo_comp=<%=session("vTipo_comp")%>&vNegociacao_comp=<%=session("vNegociacao_comp")%>&vQuartos_comp=<%=session("vQuartos_comp")%>&vVagas_comp=<%=session("vVagas_comp")%>&vValor_comp1=<%=session("vValor_comp1")%>&vValor_comp2=<%=session("vValor_comp2")%>&vValor_comp=<%=session("vValor_comp")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>" style="color:#000000"><b>Próximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>
</center>

  <table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="250">
<div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>O 
              sr(a) buscou im&oacute;veis para permutar(trocar) e obteve o resultado 
              acima ,contate nossos atendentes destes poss&iacute;veis permutantes 
              em nosso escrit&oacute;rio, para que os mesmos agendem as visitas 
              aos im&oacute;veis tanto de um como o de outro, caso n&atilde;o 
              seja realmente o que voc&ecirc; queira, a partir de agora consulte 
              constantemente a p&aacute;gina <a href="contas.asp" style="color:red" target="_blank">veja 
              sua conta cadastro gratuita</a>, pois assim, como voc&ecirc;, diariamente 
              um novo interessado em permuta se cadastra em nosso site. Obrigado 
              por sua consulta e boa sorte na sua procura.<br>
              <br>
              </strong><strong><br>
              <a href="procurar_permuta002.asp" style="color:red;">nova busca.</a></strong></font></div></td>
  </tr>
</table>
      <%End If


Else

%>

      <table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="250">
<div align="center">
              <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>O 
                sr(a) buscou im&oacute;veis para permutar(trocar) e n&atilde;o 
                obteve resultado, tente ser mais flex&iacute;vel em seus interesses, 
                pois a permuta de im&oacute;veis realmente &eacute; dif&iacute;cil 
                de ser realizada,de qualquer forma, a partir de agora, consulte 
                constantemente a p&aacute;gina <a href="contas.asp" style="color:red" target="_blank">veja 
                sua conta cadastro gratuita</a>, pois assim , como voc&ecirc;, 
                diarimente um novo interessado em permuta se cadastra em nosso 
                site. Obrigado por sua consulta e boa sorte na sua procura.</strong></font></p>
              <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><br><br>
                <a href="procurar_permuta002.asp" style="color:red;">nova busca.</a></strong></font></p>
            </div></td>
  </tr>
</table>


  <% end if %>
  <%


RS.close
Set RS = Nothing

'-------------------------------

rs444Verificaconta2.close

set rs444Verificaconta2 = nothing



'-------------------------------------------


'-------------------------------

rs444Verificaconta02.close

set rs444Verificaconta02 = nothing



'-------------------------------------------




'-------------------------------

rs444Verificaconta3.close

set rs444Verificaconta3 = nothing



'-------------------------------------------


'-------------------------------

rs444Verificaconta002.close

set rs444Verificaconta002 = nothing



'-------------------------------------------

set objfso = nothing

'---------------------------------




'-------------------------------

rs444Tipo22.close

set rs444Tipo22 = nothing



'-------------------------------------------


'-------------------------------



set rs77 = nothing



'-------------------------------------------



'-------------------------------

rs4.close

set rs4 = nothing



'-------------------------------------------




'-------------------------------



set rs7 = nothing



'-------------------------------------------















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
</table>
</td>
  </tr>
</table>

<%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>


<%

conexao3.close

set conexao3 = nothing

%>
<!--#include file="dsn2.asp"-->
</body>
</html>
