<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%response.Buffer = true %>
<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utiliz�o

rsMarcas3.ActiveConnection = Conexao3

rsMarcas3.Open SqlMarcas3, Conexao3


While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utiliz�o

rsCarros3.ActiveConnection = Conexao3

rsCarros3.Open SqlCarros3, Conexao3


'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1
While NOT rsCarros3.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



  rsMarcas3.Close           
		   
           Set rsMarcas3 = Nothing
		   
		   
		   rsCarros3.Close           
		   
           Set rsCarros3 = Nothing






End Function


        




%> 


<%
'Criando conex�o com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1 ASC" 

Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open sql3, Conexao3







%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs4,strSQL4,Conexao
  
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
	
	rs4.Open strSQL4, Conexao3



%>



<%
Function EscreveFuncaoJavaScript222 ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros222 (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo2.options[doublecombo.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 




Set rsMarcas333 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsMarcas333.ActiveConnection = Conexao3
	
	
	rsMarcas333.Open SqlMarcas333, Conexao3
	



While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas333("id_combo2")&""



Set rsCarros333 = Server.CreateObject("ADODB.RecordSet")

	rsCarros333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsCarros333.ActiveConnection = Conexao3
	
	
	rsCarros333.Open SqlCarros333, Conexao3
	
'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros333.EoF

Response.Write "doublecombo.combo5.options[" & i & "] = new Option('" & rsCarros333("nome_combo3") & "','" & rsCarros333("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros333.MoveNext
Wend


Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas333.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 


rsMarcas333.close

set rsMarcas333 = nothing

rsCarros333.close

set rsCarros333 = nothing



End Function
%> 


<%

'Criando conex�o com o banco de dados! 

'

Sql333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 


Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs333.ActiveConnection = Conexao3
	
	
	rs333.Open Sql333, Conexao3
	





'------------------------------selecionar os tipos de im�vel para o formul�rio-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo   FROM tipo  ORDER BY tipo ASC" 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao3











'-------------------------------------------------------------------------------------------------





%> 



<html>

<!--#include file="style4_imoveis.asp"-->
<head>

<script>
function isValidDigitNumber2 (doublecombo2){




var strValidNumber1_4="1234567890,";
for (nCount=0; nCount < doublecombo2.ref.value.length; nCount++) 
		{
strTempChar1_4=doublecombo2.ref.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Este formul�rio s� pode conter n�meros!");
doublecombo2.ref.focus();
doublecombo2.ref.select();
return false;
}
}

if (doublecombo2.ref.value == "") {
        alert("Este formul�rio est� vazio!");
        doublecombo2.ref.focus();
		doublecombo2.ref.select();
        return false;
    }


}
</script>



<script>

// Verifica se somente n�meros foram digitados no campo
function isValidDigitNumber (doublecombo) 



{




{


if (doublecombo.txt_nome.value == "Seu nome:") {
		alert("Por favor,deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}

if (doublecombo.txt_nome.value == "") {
		alert("Por favor,deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}









if (doublecombo.txt_telefone.value == "Seu telefone:") {
		alert("Por favor, coloque seu telefone , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_telefone.value == "") {
		alert("Por favor, coloque seu telefone , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "Seu email:") {
		alert("Por favor, coloque seu email , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "") {
		alert("Por favor, coloque seu email , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}





var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas n�meros!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}



if (doublecombo.combo1.value == "cqualquer") {
		alert("Voc� precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}



if (doublecombo.example2.value == "nqualquer") {
		alert("Por favor, escolha um tipo de negocia��o , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.example2.focus();
		
		return false;
}



if (doublecombo.stage22.value == "vqualquer") {
		alert("Por favor, escolha um faixa de valor na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
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
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>
</head>












<body bgcolor="E17508" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
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
                      <font face="Verdana" size="1" color="#FFFFFF"><B>Imobili�ria 
                      Veja: Av.Ant�rtico 315 - Jardim do Mar - SBC - CEP 09726-150. 
                      Tel: 4123-72-44. CRECI: 11.676-J. Atuando no mercado imobili�rio do grande ABC desde fevereiro de 1991.</B></font>
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
                   <option value="bqualquer" selected>Bairro/Regi�o</option>
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
                  <option value="nqualquer">Negocia��o </option>
                  <option value="nqualquer" >Qualquer um </option>
				  <option  value="Aluguel">Aluguel </option>
                  <option value="Venda">Venda </option>
                  
                </select></td>
                            </tr>
                            <tr>
                                  <td><select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000020000">At� 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 at� 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 at� 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 at� 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 at� 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 at� 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 at� 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 at� 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 at� 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 at� 400.000,00</option>
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
  <td width="755" height="10" bgcolor="863F15">
  <table width="755" height="10" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="136"> <div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="quem_somos.asp" style="color:#FFCC00">Quem somos</a></strong></font></div></td>
            <td width="116"> <div align="right"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="onde_estamos.asp" style="color:#FFCC00">Onde 
                      estamos </a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="servicos.asp" style="color:#FFCC00">Servi&ccedil;os</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="financiamento.asp" style="color:#FFCC00">Financiamento/FGTS</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="dicas.asp" style="color:#FFCC00">Dicas</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('form_enviar_email.asp')" style="color:#FFCC00">Contato</a></strong></font></div></td>
          </tr>
        </table>
  </td>
  </tr>
  <tr>
            <td width="755" height="50"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Aten&ccedil;&atilde;o!!</font></strong> 
                <strong><font size="1">Na busca acima deixe seu nome e telefone 
                para podermos lhe atender com mais rapidez e agilidade, os clientes 
                que fornecerem essas informa&ccedil;&otilde;es ser&atilde;o clientes 
                com atendimento preferencial.</font></strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
  </tr>
</table></form>

<form name="doublecombo2" onSubmit="return isValidDigitNumber2(this);" method="post" action="listar_referencia.asp">
  <table width="566" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
    <td width="243"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Procurar 
          im&oacute;vel por c&oacute;digo de refer&ecirc;ncia:<font color="EAA813"> 
          :</font> </strong></font></div> </td>
    <td width="149"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
      <input type="text" name="ref"  style="HEIGHT: 18px; WIDTH: 149px; ; font-size : 10px; background: FFFFFF; color:000000;">
      </strong></font></td>
    <td width="149"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
      <input name="image2" type="image"  src="bt_procurar001.jpg" width="149" height="18" border="0">
      </strong></font></td>
    <td>&nbsp;</td>
  </tr>
</table>
</form>
<center>
<table width="568" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="568" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
    <td width="568" height="153"><table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
            <td width="8" height="153"><img src="left_display2.jpg" width="8" height="2080"></td>
          <td bgcolor="FE9225"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Use 
                      os links abaixo para simular o seu finaciamento</strong></font><br><br>
				
				<table width="530" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td><div align="center"><a href="http://www.bradesco.com.br"><img src="link_bradesco.jpg" width="104" height="30" border="0"></a></div></td>
                    <td><div align="center"><a href="http://www.unibanco.com.br"><img src="link_unibanco.gif" height="30" border="0"></a></div></td>
                    <td><div align="center"><a href="http://www.santander.com.br"><img src="link_satander.jpg" width="90" height="30" border="0"></a></div></td>
                  </tr>
                  <tr> 
                    <td height="18">
<div align="center"></div></td>
                    <td height="18">
<div align="center"></div></td>
                    <td height="18">
<div align="center"></div></td>
                  </tr>
                  <tr> 
                    <td><div align="center"><a href="http://www.caixa.gov.br"><img src="link_caixa.jpg" width="77" height="26" border="0"></a></div></td>
                    <td><div align="center"><a href="http://www.itau.com.br"><img src="link_itau.jpg" width="34" height="30" border="0"></a></div></td>
                    <td><div align="center"><a href="http://www.bancoreal.com.br"><img src="link_real.gif" width="136" height="30" border="0"></a></div></td>
                  </tr>
                </table>

				<div align="left"><br>
                        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                        <strong><font size="2">Veja abaixo como e quando utilizar 
                        o seu FGTS na aquisi��o do seu im�vel</font><br><br>
				  1)Quais 
                  os pr&eacute; requisitos para a utiliza&ccedil;&atilde;o dos 
                  recursos do FGTS na casa pr&oacute;pria?</strong><br><br>
                  Pode ser utilizado por proponente(s) que:<br>
                  1. N�o seja(m) promitente(s) comprador(es) ou propriet�rio(s) 
                  de im�vel residencial financiado pelo SFH, em qualquer parte 
                  do territ�rio nacional.<br>
                  2. N�o seja(m) promitente(s) comprador(es) ou propriet�rio(s) 
                  de im�vel residencial conclu�do ou em constru��o:<br>
                  &nbsp;- no atual munic�pio de resid�ncia;<br>
                  &nbsp;- no munic�pio onde exer�a sua ocupa��o principal, nos 
                  munic�pios lim�trofes e na regi�o metropolitana. </font><font size="1"> 
                  </font></div>
                <font size="1">
<p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>2) 
                  Pode o propriet�rio que possua fra��o de im�vel residencial 
                  quitado ou financiado, conclu�do ou em constru��o, utilizar 
                  o FGTS para adquirir outro im�vel?<br>
                  </b>Sim, desde que detenha fra��o ideal igual ou inferior a 
                  40%.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>3) 
                  Pode o c�njuge separado, propriet�rio de im�vel residencial, 
                  conclu�do ou em constru��o, utilizar o FGTS na compra de outro 
                  im�vel?<br>
                  </b>Sim, desde que tenha perdido o direito de nele residir e 
                  atenda as demais condi��es necess�rias para utiliza��o do FGTS 
                  na compra do novo im�vel.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>4) 
                  Pode o propriet�rio, que possui uma fra��o de im�vel residencial 
                  quitado ou financiado, comprar a fra��o remanescente do mesmo 
                  im�vel, com recursos do FGTS?<br>
                  </b>Sim, desde que figure na mesma escritura aquisitiva do im�vel 
                  como co-propriet�rio ou no mesmo contrato de financiamento.<br>
                  Neste caso particular, a deten��o de fra��o ideal pode ultrapassar 
                  os 40%.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>5) 
                  Pode ser utilizado o FGTS para compra de im�vel residencial 
                  quem for propriet�rio de lotes ou terrenos?<br>
                  </b>Sim, desde que comprovada a inexist�ncia de edifica��o, 
                  atrav�s da apresenta��o do carn� do Imposto Predial Territorial 
                  Urbano - IPTU e matr�cula atualizada do im�vel.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>6) 
                  Pode o detentor de im�vel residencial recebido por doa��o ou 
                  heran�a utilizar o FGTS na compra de outro im�vel?<br>
                  </b>Sim, desde que o im�vel recebido por doa��o ou heran�a esteja 
                  gravado com cl�usula de usufruto vital�cio em favor de terceiros.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>7) 
                  Posso utilizar meu FGTS para constru��o?<br>
                  </b>Sim, desde que a constru��o seja feita em regime de cooperativa 
                  ou cons�rcio de im�veis, ou que haja um financiamento com um 
                  Agente Financeiro, ou com um construtor (pessoa f�sica ou jur�dica). 
                  O construtor dever� apresentar cronograma de obra.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>8) 
                  � permitida a utiliza��o do FGTS na aquisi��o e constru��o de 
                  im�vel misto, ou seja, aquele destinado � resid�ncia e instala��o 
                  de atividades comerciais?<br>
                  </b>Somente para a fra��o correspondente � unidade residencial.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>9) 
                  Onde o im�vel a ser adquirido deve estar localizado?</b><br>
                  O im�vel a ser adquirido deve estar localizado:<br>
                  1. No munic�pio onde o(s) adquirente(s) exer�a(m) a sua ocupa��o 
                  principal, salvo quando se tratar de munic�pio lim�trofe ou 
                  integrante da regi�o metropolitana; ou<br>
                  2. No munic�pio em que o(s) adquirente(s) comprovar(em) que 
                  j� reside(m) h� pelo menos 01 ano, cuja comprova��o � feita 
                  mediante a apresenta��o de, no m�nimo, 02 documentos simult�neos, 
                  tais como: contrato de aluguel; contas de �gua, luz, telefone 
                  ou g�s; recibos de condom�nio; ou declara��o do empregador ou 
                  de institui��o banc�ria.<br>
                  O atendimento dos requisitos � exigido, tamb�m, em rela��o ao 
                  coadquirente, exceto ao c�njuge. Tratando-se de concubinas, 
                  a comprova��o de endere�o de um deles pode ser substitu�da pela 
                  declara��o de ambos de que a identidade de endere�o decorre 
                  de uni�o n�o conjugal de natureza familiar, est�vel e duradoura, 
                  de conhecimento p�blico.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>10) 
                  O FGTS pode ser utilizado pelos c�njuges ou companheiros independente 
                  do regime de casamento?<br>
                  </b>Sim, desde que aquele que n�o � adquirente principal compare�a 
                  no contrato como coadquirente.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>11) 
                  � permitida a utiliza��o do FGTS por companheiros que vivem 
                  em regime de concubinato?<br>
                  </b>Sim, desde que o(a) companheiro(a) compare�a no contrato 
                  como coadquirente.<o:p> </o:p> </font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>12) 
                  Para comprar o im�vel com recursos do FGTS � necess�rio ter 
                  quanto tempo de Fundo?<br>
                  </b>O Adquirente dever� comprovar o tempo m�nimo de tr�s anos 
                  de trabalho sob o regime do FGTS.<br>
                  A comprova��o ser� atrav�s dos dados constantes no extrato da 
                  conta vinculada, quando este for suficiente, ou na Carteira 
                  de Trabalho.<br>
                  <br>
                  Para c�mputo desse tempo � considerada a soma de todos os per�odos, 
                  consecutivos ou n�o, trabalhados sob o regime do FGTS, em uma 
                  ou mais empresas.<br>
                  Tratando-se de trabalhador avulso, a efetiva presta��o de servi�os 
                  � considerada de acordo com declara��o fornecida pelo sindicato 
                  da respectiva categoria profissional.<br>
                  <br>
                  Tratando-se de utiliza��o por mais de um adquirente, � exigido 
                  de cada um deles o tempo m�nimo de trabalho sob o regime do 
                  FGTS, podendo ser utilizadas todas as contas das quais sejam 
                  titulares.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>13) 
                  Im�vel comprado com recursos do FGTS, atrav�s das modalidades 
                  aquisi��o ou constru��o, pode ser objeto de outra transa��o 
                  de compra e venda com recursos do FGTS?<br>
                  </b>Somente ap�s decorridos, no m�nimo, 3 anos, contados da 
                  data da �ltima negocia��o realizada ou da libera��o da �ltima 
                  parcela para constru��o.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>14) 
                  � permitida a utiliza��o dos recursos do FGTS para a aquisi��o 
                  de lotes/terrenos, amplia��o, reforma, melhoria de im�vel residencial/comercial 
                  ou realiza��o de infra-estrutura?<br>
                  </b>N�o. � vedada a utiliza��o dos recursos da conta vinculada 
                  para tais fins.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>15) 
                  � permitida a utiliza��o do FGTS para aquisi��o de im�vel destinado 
                  exclusivamente � moradia de familiares, dependentes do adquirente 
                  ou de terceiros?<br>
                  </b>N�o. Os recursos do FGTS s� poder�o ser utilizados para 
                  moradia pr�pria. </font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>16) 
                  Qual o valor m�ximo de avalia��o do im�vel estabelecido para 
                  aquisi��o com recursos do FGTS?<br>
                  </b>A aquisi��o com recursos do FGTS est� limitada �queles im�veis 
                  avaliados em, no m�ximo, R$ 300.000,00.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>17) 
                  De que forma pode ser utilizado o FGTS para Pagamento Parcial 
                  do Pre�o de Aquisi��o de Im�vel Residencial Conclu�do, financiado 
                  fora do SFH?<br>
                  </b>Os recursos do FGTS podem ser utilizados:<br>
                  1. Na aquisi��o de im�vel residencial conclu�do, vinculado a 
                  financiamento com agentes n�o integrantes do SFH, tais como 
                  PREVI, Clube Imobili�rio - FUNCEF, entre outros;<br>
                  2. Quando parte do pre�o do im�vel for financiado pelo vendedor, 
                  pessoa jur�dica;<br>
                  3. Em opera��es realizadas no Sistema Hipotec�rio - SH e Carta 
                  de Cr�dito CAIXA;<br>
                  4. Para complementar o valor do im�vel conclu�do que esteja 
                  sendo adquirido atrav�s de carta de cr�dito concedida por administradora 
                  de cons�rcio de im�veis, devidamente credenciada pelo Banco 
                  Central do Brasil;<br>
                  5. Aquisi��o de fra��o ideal remanescente por proponente(s) 
                  participante(s) no mesmo contrato de financiamento ou escritura 
                  aquisitiva.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>18) 
                  Quais s�o os limites de utiliza��o do FGTS nesta modalidade?</b><br>
                  O valor do FGTS a ser utilizado, somado ao valor financiado/parcelado, 
                  n�o pode exceder ao menor dos valores: de compra e venda ou 
                  de avalia��o efetuada pela CAIXA. O valor m�ximo de avalia��o 
                  est� limitado a R$ 300.000,00.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>19) 
                  De que forma pode ser utilizado o FGTS para Pagamento total 
                  do pre�o de aquisi��o de im�vel residencial conclu�do?<br>
                  </b>Os recursos do FGTS podem ser utilizados:<br>
                  1. Na aquisi��o de im�vel residencial conclu�do, � vista, havendo 
                  ou n�o complementa��o com recursos pr�prios;<br>
                  2. Para complementar o valor do im�vel que esteja sendo adquirido 
                  atrav�s de carta de cr�dito concedida por administradora de 
                  cons�rcio de im�veis, credenciada pelo Banco Central do Brasil. 
                  A d�vida deve estar quitada junto ao cons�rcio, inexistindo 
                  outro financiamento complementar;<br>
                  3. Aquisi��o de fra��o ideal remanescente por proponente(s) 
                  participante(s) na mesma escritura aquisitiva.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>20) 
                  Os recursos do FGTS podem ser utilizados para pagamento da poupan�a 
                  e/ou redu��o de financiamento de im�vel financiado atrav�s do 
                  SFH junto � CAIXA?<br>
                  </b>Sim.</font></p>
                <p align="justify"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>21) 
                  Quais as modalidades de utiliza��o do FGTS na Constru��o de 
                  Im�vel Residencial?<br>
                  </b>1. Constru��o de Im�vel, financiado no SFH;<br>
                  2. Constru��o de Im�vel residencial, financiado fora do SFH;<br>
                  3. Aquisi��o parcelada de Im�vel Residencial em Constru��o, 
                  fora do SFH;<br>
                  4. Constru��o de Im�vel residencial, atrav�s de financiamento 
                  de um construtor (pessoa f�sica ou jur�dica) ou autofinanciamento 
                  (Cooperativas ou Cons�rcios habitacionais).</font></p>
                <p align="justify" ><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><b>22) 
                  Quais s�o os limites de utiliza��o do FGTS nas modalidades de 
                  Constru��o de Im�vel Residencial?<br>
                  </b>O valor do FGTS a ser utilizado, somado ao valor do financiamento, 
                  quando houver, n�o pode exceder ao menor dos valores:<br>
                  1. Limite m�ximo de valor do im�vel estabelecido para as opera��es 
                  no SFH;<br>
                  2. Custo total da obra, em caso de constru��o em terreno pr�prio;<br>
                  3. Custo total da obra, acrescido do valor do terreno, no caso 
                  de aquisi��o de terreno associada � constru��o;<br>
                  4. Valor da avalia��o efetuada pela CEF;<br>
                  5. Valor de compra e venda.</font></p>
                </font> 
                <p align="justify"></div></td>
            <td width="8" height="153"><img src="right_display2.jpg" width="8" height="2080"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="568" height="11"><img src="bottom_display2.jpg" width="568" height="11"></td>
  </tr>
</table>
</center>

  <%
 rs4.close
 
 set rs4 = nothing
 
 
 
 
 set rs4Tipo22 = nothing 
  
  
  
  
  
  %>
 <script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui � criada uma vari�vel "groups" que receber� os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a vari�vel group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero at� o n�mero de elementos do array "groups" */

group2[i2]=new Array()
/* aqui � criado o array "group" que receber� valores conforme o n�mero de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receber� valores de op��es. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("201,00 at� 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 at� 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 at� 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 at� 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 at� 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 at� 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 at� 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 at� 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 at� 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("At�  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 at� 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 at� 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 at� 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 at� 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 at� 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 at� 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 at� 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 at� 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 at� 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("Acima de 400.000,00","0000400001 1000000000")









/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


var temp2=document.doublecombo.stage22
/* aqui a vari�vel "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui � criada a fun��o "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que d� um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que � escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" ser� o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a vari�vel "location" recebe os valores de "stage2" que corresponde ao endere�o de
link para o carregamento de p�gina. */


//-->
</script>
  <%





%>
  
    
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



<% response.flush%>
  <%response.clear%>
  
  
  <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>
  
  
 <%

conexao3.close

set conexao3 = nothing

%> 
  
  
  
  <!--#include file="dsn2.asp"-->


</body>
</html>
