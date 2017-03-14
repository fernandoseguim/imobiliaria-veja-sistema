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
Conexao3.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

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


dim rsFrontPage,SQLFrontPage,objFSO 

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Set rsFrontPage = Server.CreateObject("ADODB.RecordSet")

SQLFrontPage = "SELECT * FROM imoveis where presenca_primeira like '"&"incluido"&"' ORDER BY cod_imovel DESC"

rsFrontPage.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsFrontPage.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsFrontPage.ActiveConnection = Conexao


rsFrontPage.open SQLFrontPage,Conexao

dim intRecordCount 

intRecordCount = rsFrontPage.RecordCount



%> 






<html>

<!--#include file="style4_imoveis.asp"-->
<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3) %>


<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (doublecombo) 



{




{
if (doublecombo.combo1.value == "cqualquer") {
		alert("Você precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}
}
}


</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>


<script language="javascript">
function funScroll()
{
window.scrollTo(0,0)

}		
</script>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>
</head>












<body onLoad="funScroll()" bgcolor="E17508" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">

<table width="755" border="0" cellspacing="0" cellpadding="0" bgcolor="EAA813">
  <tr>
    <td>
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);" method="post" action="listar_imoveis.asp">
  
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
    <td width="755" height="243"><table width="755" height="243" border="0"  cellpadding="0" cellspacing="0">
        <tr>
                  <td width="176" height="243" align="center" background="fundo_primeira.jpg">
<div align="center"><table width="149" border="0"  cellspacing="0" cellpadding="0" height="170">
                            <tr>
							<td width="149" height="10"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Busca 
                                  de im&oacute;veis</strong> </font> </div></td>
							</tr>
							<tr>
                                  <td height="11"><input type="text" name="ref" class="inputBox" value="Seu nome:"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;"></td>
                            </tr>
							<tr>
                                  <td><input type="text" name="ref" class="inputBox" value="Seu fone ou email:"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;"></td>
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
                                  <td><select name="txt_ocupacao" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="oqualquer">Ocupação</option>
				   <option value="oqualquer">Qualquer um</option>
                  <option value="vago">Vago</option>
				   <option value="ocupado">Ocupado</option>
                  
                 
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
                <embed src="front_page.swf" width="579" height="243" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash"></embed></object></td>
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
<center>
<%
 if int(intRecordCount) >= 6 then


%>
        <table width="755" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="370" height="150"><table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsFrontPage.movefirst%> 
                    <%=rsFrontPage("titulo_anuncio")%></strong></font></div></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                                            <td><%If objFSO.FileExists(Server.MapPath(rsFrontPage("Foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="158" height="90" border="0"></img></a><% else %><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="mini_imovel00000.jpg" width="158" height="90" border="0"></img></a><% end if %></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rsFrontPage("texto_anuncio")%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7" height="116"><img src="right_display.jpg" width="7" height="116"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="360" height="11"><img src="bottom_display.jpg" width="360" height="11"></td>
  </tr>
</table>
              <div align="center"></div></td>
            <td width="370" height="150"><table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsFrontPage.movenext%><%=rsFrontPage("titulo_anuncio")%></strong></font></div></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><%If objFSO.FileExists(Server.MapPath(rsFrontPage("Foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="158" height="90" border="0"></img></a><% else %><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="mini_imovel00000.jpg" width="158" height="90" border="0"></img></a><% end if %></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rsFrontPage("texto_anuncio")%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7" height="116"><img src="right_display.jpg" width="7" height="116"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="360" height="11"><img src="bottom_display.jpg" width="360" height="11"></td>
  </tr>
</table></td>
  </tr>
  <tr>
            <td width="370" height="150"><table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsFrontPage.movenext%><%=rsFrontPage("titulo_anuncio")%></strong></font></div></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><%If objFSO.FileExists(Server.MapPath(rsFrontPage("Foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="158" height="90" border="0"></img></a><% else %><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="mini_imovel00000.jpg" width="158" height="90" border="0"></img></a><% end if %></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rsFrontPage("texto_anuncio")%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7" height="116"><img src="right_display.jpg" width="7" height="116"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="360" height="11"><img src="bottom_display.jpg" width="360" height="11"></td>
  </tr>
</table></td>
            <td width="370" height="150"><table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsFrontPage.movenext%><%=rsFrontPage("titulo_anuncio")%></strong></font></div></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><%If objFSO.FileExists(Server.MapPath(rsFrontPage("Foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="158" height="90" border="0"></img></a><% else %><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="mini_imovel00000.jpg" width="158" height="90" border="0"></img></a><% end if %></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rsFrontPage("texto_anuncio")%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7" height="116"><img src="right_display.jpg" width="7" height="116"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="360" height="11"><img src="bottom_display.jpg" width="360" height="11"></td>
  </tr>
</table></td>
  </tr>
  <tr>
            <td width="370" height="150"><table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsFrontPage.movenext%><%=rsFrontPage("titulo_anuncio")%></strong></font></div></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                                            <td><%If objFSO.FileExists(Server.MapPath(rsFrontPage("Foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="158" height="90" border="0"></img></a><% else %><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="mini_imovel00000.jpg" width="158" height="90" border="0"></img></a><% end if %></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rsFrontPage("texto_anuncio")%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7" height="116"><img src="right_display.jpg" width="7" height="116"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="360" height="11"><img src="bottom_display.jpg" width="360" height="11"></td>
  </tr>
</table></td>
            <td width="370" height="150"><table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rsFrontPage.movenext%><%=rsFrontPage("titulo_anuncio")%></strong></font></div></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><%If objFSO.FileExists(Server.MapPath(rsFrontPage("Foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="158" height="90" border="0"></img></a><% else %><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="mini_imovel00000.jpg" width="158" height="90" border="0"></img></a><% end if %></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rsFrontPage("texto_anuncio")%>.</font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="7" height="116"><img src="right_display.jpg" width="7" height="116"></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="360" height="11"><img src="bottom_display.jpg" width="360" height="11"></td>
  </tr>
</table></td>
  </tr>
</table>

  <%else%>
 <table width="568" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="568" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
    <td width="568" height="153"><table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
          <td bgcolor="FE9225"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">IMOBILIÁRIA 
                      VEJA ATUA NO MERCADO IMOBILIÁRIO DESDE FEVEREIRO DE 1991, 
                      E TEM COMO PRINCÍPIO BÁSICO REALIZAR TRANSAÇÕES IMOBILIÁRIAS 
                      COM O MÁXIMO DE CLAREZA E HONESTIDADE, SEJA NA HORA DE VOCÊ 
                      COMPRAR OU DE VENDER O SEU IMÓVEL, PROPORCIONANDO ASSIM, 
                      AOS NOSSOS CLIENTES, TRANSAÇÕES TOTALMENTE SEGURAS E TRANSPARENTES, 
                      E DANDO TODO APOIO NECESSÁRIO DURANTE E APÓS A REALIZAÇÃO 
                      DO NEGÓCIO IMOBILIÁRIO QUE TENHA SIDO INTERMEDIADO POR NÓS.</font></div></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="568" height="11"><img src="bottom_display2.jpg" width="568" height="11"></td>
  </tr>
</table>
  <%end if%>
</center>
<br>
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
  <!--#include file="dsn2.asp"-->
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

</body>
</html>
