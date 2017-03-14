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
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1
While NOT rsCarros3.EoF

 Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf
i=i+1


rsCarros3.MoveNext

Wend




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
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao

dim rs55
dim strSQL55

Set rs55 = Server.CreateObject("ADODB.RecordSet")
	strSQL55 = "SELECT * FROM imoveis" 

rs55.open strSQL55, Conexao


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
SqlMarcas333 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas333 = Conexao3.Execute ( SqlMarcas333 )

While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas333("id_combo2")&""

Set rsCarros333 = Conexao3.Execute ( SqlCarros333 )

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


'

Sql333 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs333 = Conexao3.Execute ( Sql333 ) 



Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
%> 





<html>

<!--#include file="style4_imoveis.asp"-->
<head>


<script>
function isValidDigitNumber (doublecombo2){
var strValidNumber1_4="1234567890,";
for (nCount=0; nCount < doublecombo2.ref.value.length; nCount++) 
		{
strTempChar1_4=doublecombo2.ref.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Este formulário só pode conter números!");
doublecombo2.ref.focus();
doublecombo2.ref.select();
return false;
}
}

if (doublecombo2.ref.value == "") {
        alert("Este formulário está vazio!");
        doublecombo2.ref.focus();
		doublecombo2.ref.select();
        return false;
    }








}
</script>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=590,height=510,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>

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












}
}


</script>

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
                  <option value="Apartamento">Apartamento </option>
				   <option value="Térrea/Sobrado">Térrea/Sobrado</option>
				   <option value="Chácara">Chácara</option>
                  <option value="Flat">Flat</option>
				  <option value="Fazenda">Fazenda</option>
                  <option value="Prédio Comercial">Prédio Comercial</option>
                  <option value="Galpões">Galpões</option>
                  <option value="Sala Comercial">Sala Comercial</option>
				  <option value="Salão Comercial">Salão Comercial</option>
                  <option value="Terreno/Área">Terreno/Área</option>
                  <option value="Ponto Comercial">Ponto Comercial</option>
				  <option value="Cobertura">Cobertura</option>
                  
                 
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
</table>
</form>

<form name="doublecombo2" onSubmit="return isValidDigitNumber(this);" method="post" action="listar_referencia.asp">
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















<%

varCodImovel = request.Form("ref")


 
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis Where  cod_imovel = "&varCodImovel
	 
   Conexao.Open dsn
   
   dim vData, vEnderecoIP
   vData = now()
   vEnderecoIP = request.ServerVariables("REMOTE_ADDR")
   
    if  varCodImovel <> ""  then
 
  Conexao.execute"Insert into referencia_procurados(referencia,enderecoIP,data) values( '"& varCodImovel &"','"& vEnderecoIP &"','"& vData &"')"
  
  end if
   
   
   
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
		
		
		
	if not(rs.eof) and not(rs.bof) or (rs.recordcount >= 6) then
		
          
 %>


<table width="518" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  
  

  
  
    <td width="518" height="155">
	
	  
 
<% varCodimovel = rs("COD_Imovel") %>
  
  
  
  
	
	
	
	
	
	
	
	      <table width="568" border="0" align="center" cellpadding="0" cellspacing="0">
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
                              <td bgcolor="<%=escuro%>"><% If objFSO.FileExists(Server.MapPath(rs("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')"><img src="<%=rs("foto_pequena")%>" width="158" height="91" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="color:#FFFFFF"><strong>Foto não disponível</strong></a></font></div><%end if%></td>
                            </tr>
                          </table>
                        </center>
					  
					  </td>
                                  <td bgcolor="FE9225" width="532"><table width="530" height="93" border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td width="441"><div align="center"><font face="Verdana, arial" size="1" color="#FFFFFF"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="color:#FFFFFF"><%=RS("obs_imovel")%><br><br>
                                            <strong>Veja mais detalhes.<br>Código de referência<%=RS("cod_imovel")%> </strong></a></font></div></td>
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
<% end if %>

 <%
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   Set objFSO = Nothing
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
<%  EscreveFuncaoJavaScript222 ( Conexao3) %>
</body>
</html>
