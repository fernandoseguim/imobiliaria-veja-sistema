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
'indica o tipo de cursor utilizão

rsMarcas3.ActiveConnection = Conexao3


rsMarcas3.Open SqlMarcas3, Conexao3




While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"




Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizão

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

set rsCarros3 = nothing




End Function
%> 


<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 






Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizão

rs3.ActiveConnection = Conexao3


rs3.Open Sql3, Conexao3



%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs4,strSQL4,Conexao
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
	
	
	

	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizão

rs4.ActiveConnection = Conexao3


	
	
	
	
	
	rs4.Open strSQL4, Conexao3



%>







<%
Function EscreveFuncaoJavaScript222 ( Conexao3)
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
'indica o tipo de cursor utilizão

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
'indica o tipo de cursor utilizão

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

Sql333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 





Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizão

rs333.ActiveConnection = Conexao3


rs333.Open Sql333, Conexao3




 dim vNome
  dim vTelefone
  dim vEmail
  
  vNome = request.form("txt_nome")
  
  if vNome = "" then
  vNome = "não informado"
  end if
  
  session("nome") = vNome
  
  vTelefone = request.form("txt_telefone")
  if vTelefone = "" then
  vTelefone = "não informado"
  end if

session("telefone") = vTelefone



 vEmail = request.form("txt_email")


 if vEmail = "" then
  vEmail = "não informado"
  end if
  
  session("email") = vEmail




'--------------------------Atualizar acesso-----------------------------

dim rs444VerificaConta002,strSQL444VerificaConta002
   
    Set rs444VerificaConta002 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta002 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou FROM compradores where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	
	
	rs444VerificaConta002.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta002.CursorType = 3
'indica o tipo de cursor utilizão

rs444VerificaConta002.ActiveConnection = Conexao3
	
	
	
	
	 rs444VerificaConta002.Open strSQL444VerificaConta002, Conexao3
	
	
	
	






	
	
	
	
	

if  not rs444VerificaConta002.eof then




	 Conexao3.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta002("cod_compradores")
	end if 





'----------------------------------------------------------------------------










'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 
	rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizão

rs444Tipo22.ActiveConnection = Conexao3
	 
	 
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao3







 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
		
	rs444Tipo23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo23.CursorType = 3
'indica o tipo de cursor utilizão

rs444Tipo23.ActiveConnection = Conexao3
	
	
	
	
	 rs444Tipo23.Open strSQL444Tipo23, Conexao3




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




if (doublecombo.stage22.value == "vqualquer") {
		alert("Por favor, escolha um faixa de valor na busca , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}








}
}


</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=510,resizable=no,scrollbars=yes')
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
    <td><form name="doublecombo"  method="post" onSubmit="return isValidDigitNumber(this);" action="listar_imoveis.asp">

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
                  <td width="176" height="243" align="center" background="fundo_primeira.jpg" bgcolor="#000000"> 
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


<%

dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.captacao,imoveis.cod_imovel,imoveis.data FROM imoveis where (telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%') and captacao <>'"&"internet"&"' and captacao <>'"&"não informado"&"' " 
	 
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
                              <td><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="5"><a href="acessoLink03.asp?varTelefone=<%=session("telefone")%>" style="color:#FFFFFF;" target="_blank">Ol&aacute; 
                                  sr(a) <%=session("nome")%></a></font><a href="acessoLink03.asp?varTelefone=<%=session("telefone")%>" style="color:#FFFFFF;" target="_blank"><br>
                                  Obrigado por retornar ao nosso site, voc&ecirc; 
                                  est&aacute; conosco desde o dia <%=rs444VerificaConta("data")%> 
                                  e tem sua conta gratuita de vendedor de im&oacute;vel.Se 
                                  quiser verificar se novos compradores foram 
                                  indicados para o sr(a) clique aqui.</a></strong></font></div></td>
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





      <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2
dim vQuartos
dim vVagas


 vnegociacao=request.form("example22")
 session("vnegociacao") = vnegociacao
  vTipo=request.form("txt_tipo")
  session("vTipo") = vTipo
  
  vVagas=request.form("txt_vagas")
  session("vVagas")=vVagas
  
  if session("vVagas") = "" then
  session("vVagas") = request.querystring("vVagas")
  end if
  
  vQuartos=request.form("txt_quartos")
  session("vQuartos")=vQuartos
  
  if session("vQuartos") = "" then
  session("vQuartos") = request.querystring("vQuartos")
  end if
  
  
  
   vValor=request.form("stage222")
   
   if vValor = "" then
   vValor = request.querystring("vValor")
   end if
   
   session("vValor")=vValor
   session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)
   
   
      
   
   
   
   dim vCidade2
   
   if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")
end if
   
   
    vCidade2=request.form("combo3")
	
	
	
	session("vCidade2") = vCidade2
	 if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")

end if
	  
	
	if session("vCidade2") <> "cqualquer" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade2")
 
 rs2.open SQL2,Conexao3,2,1
 
 vCidade = rs2("nome_combo1")
 
 rs2.close
 
 set rs2 = nothing
 
 else
 vCidade = vCidade2
 end if

	session("vCidade")= vCidade
	
	
	
	dim vBairro2
	 vBairro2=request.form("combo4")
	 session("vBairro2") = vBairro2
	 if session("vBairro2") = "" then
session("vBairro2") = request.querystring("vBairro2")

end if
	 
	 if session("vBairro2") <> "bqualquer" then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& session("vBairro2")
 
 rs3.open SQL3,Conexao3,2,1

 vBairro = rs3("nome_combo2")
 
 rs3.close
 
 set rs3 = nothing
 
 else
 vBairro = vBairro2
	end if                                      
									
	 
	 
	 
	 session("vBairro")= vBairro
	 
	 
	 
dim vvila2,vvila
vvila2=request.form("combo6")
	
	
	
	session("vvila2") = vvila2
	 if session("vvila2") = "" then
session("vvila2") = request.querystring("vvila2")

end if



if session("vvila2") <> "vlqualquer" then
	
	dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 from combo3 where id_combo3 ="&session("vvila2")
 
 rs22.open SQL22,Conexao3,2,1
 
 vvila = rs22("nome_combo3")
 
 rs22.close
 
 set rs22 = nothing
 
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

if session("vVila") = "" then
session("vVila") = request.querystring("vVila")
end if



if session("vValor") = "" then
session("vValor") = request.querystring("vValor")
end if

if session("vTipo") = "" then
session("vTipo") = request.querystring("vTipo")
end if

if session("vNegociacao") = "" then
session("vNegociacao") = request.querystring("vNegociacao")
end if

 '---------------------------------------------------  
 
 '-------------------------Cidade-----------------------------------
 dim stringCidade
 dim stringIndex
 stringIndex = " where cod_compradores <>"&"0"&"" 

if  session("vCidade") <> "cqualquer" then
stringCidade = " and (cidade='"& session("vCidade")&"' or cidade='"& "não informado"&"') "
else
stringCidade = ""
end if
'-----------------------------------------------------------------------


'------------------------------------Bairro-----------------------------

dim stringBairro
 
  if session("vBairro") = "bqualquer"  then
  stringBairro = ""
  
  else
  stringBairro = " and (bairro like '%"&session("vBairro")&"%' or bairro like '%"&"não informado"&"%') "
  
  
  end if
  
  '------------------------------------Vila-----------------------------

dim stringVila
 
  if session("vVila")="vlqualquer"  then
  stringVila = ""
  
  else
  stringVila = " and (vila='"&session("vVila")&"' or vila='"&"não informado"&"')"
  
  
  end if
  

'--------------------------------------Tipo-------------------------------


	dim stringTipo
 
  if session("vTipo")<>"tqualquer"  then
  stringTipo = " and Tipo like'%"&session("vTipo")&"%'"
  else
  stringTipo = ""
  end if
 
 
  
  

	
	'--------------------------------------------------------------------------------
	
	
	
	
	'-------------------------------Negociacao-------------------------------
	
	
	
	dim stringNegociacao
 
  if session("vNegociacao")<>"nqualquer"  then
  stringNegociacao = " and Negociacao ='"&session("vNegociacao")&"'"
  else  
  stringNegociacao = ""
  end if
  	
	
	
	'-------------------------------------------------------------------
	'---------------------------Quartos------------------------------


if  session("vQuartos") <> "qqualquer" then
stringQuartos = " and quartos >="&int(session("vQuartos"))&""
else
stringQuartos = ""
end if

'---------------------------------------------------------------------------

'---------------------------Vagas------------------------------


if  session("vVagas") <> "vgqualquer" then
stringVagas = " and vagas >="&int(session("vVagas"))&""
else

stringVagas = ""
end if





	
	
	 '----------------------------------Valor--------------------------------
	 
	
	 
	 dim stringValor
 
  if session("vValor")<>"vqualquer"  then
  stringValor = " and Valor >="& session("vValor1") &" and Valor <= "& session("vValor2") &""
  else
  stringValor = "" 
  end if
   
	 
	 '----------------------------------------------------------------------
	 
	 
	
	

	strSQL = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex&stringCidade&stringBairro&stringVila&stringTipo&stringNegociacao&stringQuartos&stringVagas&stringValor&" ORDER  BY Cod_compradores DESC"

  
  'Aqui a variável strSQL é defenida para depois ser usada no record set.
  
  
   dim EnderecoIP , vData
  vData = now()
  
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 if  vTipo <> ""  then
 
  Conexao3.execute"Insert into compradores_procurados(nome,telefone,Cidade, bairro ,tipo,negociacao,valor,enderecoIP,data,quartos,vagas) values( '"& vNome &"','"& vTelefone &"','"& vCidade &"','"& vBairro &"','"& vTipo &"','"& vNegociacao &"','"& vValor &"','"& EnderecoIP &"','"& vData &"','"& session("vQuartos") &"','"& session("vVagas") &"')"
  
  end if
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  '----------------------Verifica conta----------------------------------------
  
   dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT imoveis.cod_imovel,imoveis.captacao,imoveis.telefone,imoveis.telefone02,imoveis.telefone03 FROM imoveis where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	 
	 
	 rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao3
	 
	 
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao3
	
	
	'----------------------primeira verificação----------------------
dim rsVerificacao007
dim strSQLVerificacao007

dim rsVerificacao008
dim strSQLVerificacao008

 Set rsVerificacao007 = Server.CreateObject("ADODB.RecordSet")
    
	strSQLVerificacao007 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.quartos,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.foto_grande1,imoveis.StandBy,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.placa,imoveis.dataLastEmail,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.imovel_em_negociacao,imoveis.origem_captacao  FROM imoveis Where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%'or telefone02 like '%"&session("telefone")&"%'"
	 
   
   
rsVerificacao007.CursorLocation = 3
rsVerificacao007.CursorType = 3

        rsVerificacao007.Open strSQLVerificacao007, Conexao3 


'------------------segunda verificação---------------------------

' Set rsVerificacao008 = Server.CreateObject("ADODB.RecordSet")
    
'	strSQLVerificacao008 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores Where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%'or telefone02 like '%"&session("telefone")&"%'"
	 
   
   
'rsVerificacao008.CursorLocation = 3
'rsVerificacao008.CursorType = 3

     '   rsVerificacao008.Open strSQLVerificacao008, Conexao3 


	
	
	
	
	

if   rs444VerificaConta2.eof  and  rsVerificacao007.eof and vTipo <> "" then
 
 dim vValorConta
 dim vQuartosConta
 dim vVagasConta
 dim vValorMedio
 
 if session("vValor") = "vqualquer" then
 vValorMedio = "0"
 else
 vValorMedio = (int(session("vValor1")) + int(session("vValor2")))/2
 end if
 
  
  
  
  if session("vVagas") = "vgqualquer" then
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
 
 if session("vNegociacao") = "compra" or session("vNegociacao") = "Compra" then
 vNegociacaoConta = "venda"
 else
 vNegociacaoConta = session("vNegociacao")
 end if
 
 session("vNegociacaoConta") = vNegociacaoConta
  
	'Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& "não informado" &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vQuartosConta") &"','"& session("vNegociacaoConta") &"','"& session("vValorMedio") &"','"& now() &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& session("vVila") &"','"& session("vVagasConta") &"','"& session("vOcupacao") &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"')"
	
	
	
	
	
	
	
	
			
Conexao3.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade,indexador_indicacoes,origem_captacao,data_captacao) values( '"& session("nome") &"','"& "não informado" &"','"& session("telefone") &"','"& session("email") &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "icon_foto2.gif" &"','"& vCidade &"','"& vBairro &"','"& vTipo &"','"& "00" &"','"& "00" &"','"& session("vQuartosConta") &"','"& "não informado" &"','"& session("vVagasConta")&"','"& session("vNegociacaoConta") &"','"& int(session("vValorMedio")) &"','"& now() &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& "não informado" &"','"& "Sem Placa"&"','"& "0" &"','"& "0" &"','"& "0" &"','"& "negócio comum" &"','"&"0"&"','"&"Busca de compradores"&"','"& now()&"')"	 

	
	
	
	
	
	else
	
	
  
  
  end if
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  '------------------------------------------------------------------------------
  
  
  
  
  
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


      <table width="518" border="0" align="center" cellpadding="0" cellspacing="0">
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

  
  
<% 
dim varCodCompradores

varCodCompradores=RS("cod_Compradores") %>  
  
	
	
	
	
	
	
	<table width="568" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="568" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
    <td width="568" height="153"><table bgcolor="FE9225" width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
                      <td width="8" height="153"><img src="left_display2.jpg" width="8" height="173"></td>
          <td><table width="552" height="153" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                            <td width="552" height="30" bgcolor="FE9225">
<div align="center"><font face="Verdana, arial" size="1" color="white">Eu quero 
                                <strong> 
                                <%
						  if  RS("Negociacao") = "não informado" or RS("Negociacao") = "qualquer um"  then
						   response.write ""
						    end if 
							
							if  RS("Negociacao") = "compra"  then
						   response.write "comprar"
						    end if 
						  
						  if  RS("Negociacao") = "aluguel"  then
						   response.write "alugar"
						    end if 
						  
						  		  
						  
						  
						     %>
                                </strong> 
                                <%if RS("Tipo") <> "tqualquer" then response.write RS("Tipo") else response.write "" end if %>
                                na cidade de <a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')"><font face="Verdana, arial" size="1" color="white"><strong><%=RS("Cidade")%></strong></font></a> 
                                <a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')"><font face="Verdana, arial" size="1" color="white"><strong>
                                <% if RS("bairro") = "bqualquer" then response.write "" else response.write "</strong>no bairro de <strong>"&rs("bairro") end if%>
                                </strong></font></a> </font></div></td>
              </tr>
              <tr> 
                      <td width="552" height="16" bgcolor="E17508"> <a href="javascript:newWindow2('visualizar_comprador01.asp?varCodCompradores=<%=varCodCompradores%>')"><div align="center"><font face="Verdana, arial" size="1" color="white"><strong>Veja 
                          mais detalhes</strong></font></div></a></td>
              </tr>
              <tr> 
                <td bgcolor="FE9225"><table width="552" height="115" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                     
                      <td bgcolor="FE9225"><div align="center">
                                      <table width="552" height="115" border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                          <td width="463"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('visualizar_comprador01.asp?varCodCompradores=<%=varCodCompradores%>')" style="color:#FFFFFF"> 
                                              Olá , meu nome é <strong><%=rs("nome")%></strong>, 
                                              o sitema veja analizou os dados 
                                              do seu imóvel e o que eu desejo 
                                              comprar, e detectou a possibilidade 
                                              de negócio entre nós. Lique já para 
                                              <strong>4123-72-44</strong> e fale 
                                              com o meu atendente o sr(a) <strong><%=rs("atendimento")%></strong>, 
                                              para que o mesmo agende uma visita 
                                              minha ao seu imóvel, <strong>clique 
                                              aqui</strong> e saiba mais sobre 
                                              meus interesses e condições de pagamento. 
                                              Muito Obrigado.</a></font></div></td>
                                          <td width="89"><% if  rs("standby")<> "excluido" then %>
                                          <img src="gifstandby02.gif" width="89" height="89"></img> 
                                          
										   <%else%>
										   
										  <%end if%></td>
                                        </tr>
                                      </table>
                                    </div></td>
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
  <tr> 
    <td width="568" height="11"></td>
  </tr>
</table>
	
	
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
            <a href="?page=<%=intPage - 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vValor=<%=session("vValor")%>&vTipo=<%=session("vTipo")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vNegociacao=<%=session("vNegociacao")%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          
        <td> 
          <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
            <strong> 
            <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
            <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
            <font color="#000000">Página</font> <%=cInt(intPage)%> <font color="#000000">de</font> 
            <%=cInt(intPageCount)%> </strong></font> <strong> 
            <%End If%>
            </strong></font> </div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%> 
            <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
            <a href="?page=<%=intPage + 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vValor=<%=session("vValor")%>&vTipo=<%=session("vTipo")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vNegociacao=<%=session("vNegociacao")%>"><b><font color="#000000" face="Verdana, arial" size="1">Próximo</font></b></a> 
            <%end if%> 
            
            </font></div></td>
        </tr>
      </table>
	  
	  
	  <br>
	  
	  
  <table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="100"> 
      <div align="center">
              <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                <br>
                O sr(a) buscou interessados em comprar o seu imovel e obteve o 
                resultado acima, contate os atendentes destes poss&iacute;veis 
                compradores em nosso escrit&oacute;rio para que os mesmos agendem 
                a visita destes interessados ao seu im&oacute;vel, caso as condi&ccedil;&otilde;es 
                de pagamento ou a proposta destes n&atilde;o lhe agrade, a partir 
                de agora consulte constantemente a p&aacute;gina <a href="contas.asp" style="color:red" target="_blank">veja 
                sua conta cadastro gratuita</a> pois v&aacute;rios compradores 
                / locat&aacute;rios s&atilde;o constatemente cadastrados em nosso 
                site, e um destes novos interessados pode ser o comprador do seu 
                im&oacute;vel, obrigado por sua consulta e boa sorte. <br>
				<br>
				<br>
			  <a href="procurar_compradores01.asp" style="color:red;">nova 
              busca.</a></strong></font> </div></td>
  </tr>
</table>
  
	  
	  <br>
	  
	  
	
   
  
<%end if%>


  


<%Else%>

<table width="750" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="300"> 
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <br>
              <br>
              O sr(a) buscou interessados em comprar o seu im&oacute;vel e n&atilde;o 
              encontrou, reveja a avalia&ccedil;&atilde;o do seu im&oacute;vel, 
              pois temos muitos clientes compradores cadastrados para todo tipo 
              de im&oacute;vel,caso contr&aacute;rio, a partir de agora consulte 
              constantemente a p&aacute;gina <a href="contas.asp" target="_blank" style="color:red">veja 
              sua conta cadastro gratuita</a> , pois v&aacute;rios compradores/locat&aacute;rios 
              s&atilde;o constantemente cadastrados em nosso site, e um destes 
              novos interessados, pode ser o comprador do seu im&oacute;vel. Obrigado 
              por sua consulta e boa sorte na procura.</strong><strong><br>
              <br>
              <br>
              <a href="procurar_compradores01.asp" style="color:red;">nova busca.</a></strong></font></div>
		<br>
        <br>
		<br>
		
        <br>
		<br>
        <br>
        </td>
  </tr>
</table>

<%end if%>
  
  </table>
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
  
  <%
  
  '------------------------------
  
  rs333.close
  
  set rs333 = nothing
  
  '-----------------------------------
  
  
  '------------------------------
  
  rs444Verificaconta2.close
  
  set rs444Verificaconta2 = nothing
  
  '----------------------------------- 
  
  
   '------------------------------
  
  rs444Tipo22.close
  
  set rs444Tipo22 = nothing
  
  '-----------------------------------
  
  
   '------------------------------
  
  rs444Tipo23.close
  
  set rs444Tipo23 = nothing
  
  '-----------------------------------
  
  
   '------------------------------
  
  rs444Verificaconta.close
  
  set rs444Verificaconta = nothing
  
  '-----------------------------------
  
  
  
   '------------------------------
  
  rs444Verificaconta002.close
  
  set rs444Verificaconta002 = nothing
  
  '-----------------------------------
  
  
  
   '------------------------------
  
  rs.close
  
  set rs = nothing
  
  '-----------------------------------
  
  
  
  
  
  
  
  
  
  
  
  
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
  <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>
  <!--#include file="dsn2.asp"-->




</body>
</html>
