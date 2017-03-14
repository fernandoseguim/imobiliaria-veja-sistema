<!--#include file="dsn.asp"-->
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
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao



%>





<html>

<!--#include file="style4_imoveis.asp"-->
<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>


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
</head>












<body bgcolor="EAA813" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo"  method="post" action="listar_imoveis.asp">

<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="755" height="78"><img src="top_page2.jpg" width="755" height="78"></td>
  </tr>
  <tr>
    <td width="755" height="243"><table width="755" height="243" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="176" height="243" align="center" bgcolor="#000000"> 
            <table width="164" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="164" height="18"><img src="top_find.jpg" width="164" height="18"></td>
                </tr>
                <tr>
                  <td width="164" height="153"><table width="164" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="7" height="153"><img src="left_find.jpg" width="7" height="153"></td>
                        <td width="149" height="153" bgcolor="E37307"><table width="149" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                  <td>
								  <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 149px; font-size : 10px; background: F1991B; color:FFFFFF; ">
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
                                  <td><select name="combo2" class="inputBox" style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px; background: F1991B; color:FFFFFF;">
                   <option value="bqualquer" selected>Bairro</option>
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
                                  <td><select name="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px; background: F1991B; color:FFFFFF;">
                  <option value="tqualquer">Tipo</option>
				   <option value="tqualquer">Qualquer um</option>
                  <option value="Apartamento">Apartamento </option>
				   <option value="Casa">Casa</option>
				   <option value="Comercial">Comercial</option>
                  <option value="Flat">Flat</option>
				  <option value="Rural">Rural</option>
                  <option value="Terreno">Terreno</option>
                 
                  
                 
                </select></td>
                            </tr>
                            <tr>
                                  <td><select name="example2" size="1" class="inputBox" id="select7" onChange="redirect2(this.options.selectedIndex)" style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px; background: F1991B; color:FFFFFF;">
                  <option value="nqualquer">Negociação </option>
                  <option value="nqualquer" >Qualquer um </option>
				  <option  value="Aluguel">Aluguel </option>
                  <option value="Venda">Venda </option>
                  
                </select></td>
                            </tr>
                            <tr>
                                  <td><select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px; background: F1991B; color:FFFFFF;">
                  <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000020000">menos de 20.000,00</option>
                  <option value="0000020000 0000050000">20.000,00 até 50.000,00</option>
                  <option value="0000050000 0000100000">50.000,00 até 100.000,00</option>
                  <option value="0000100000 0000200000">100.000,00 até 200.000,00</option>
                  <option value="0000200000 1000000000">acima de 200.000,00</option>
                </select></td>
                            </tr>
                            <tr>
                              <td><input name="image" type="image"  src="bt_procurar001.jpg" width="149" height="18" border="0"></td>
                            </tr>
                            <tr>
                                    <td>&nbsp;</td>
                            </tr>
							
                            <tr>
                              <td></td>
                            </tr>
                          </table>
                                
                       
<td width="10" height="153"><img src="right_find.jpg" width="8" height="153"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td width="164" height="46"><img src="bottom_find.jpg" width="164" height="46"></td>
                </tr>
              </table>
		  
		    <div align="center"></div></td>
            <td width="579" height="243"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="579" height="243">
                <param name="movie" value="front_page.swf">
                <param name="quality" value="high">
                <embed src="front_page.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="579" height="243"></embed></object></td>
        </tr>
      </table></td>
  </tr>
  <tr>
  <td width="755" height="10" bgcolor="863F15"></td>
  </tr>
</table></form>
<center>
  <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
  <a href="procurar_permuta01.asp">Nova busca.</a></strong></font> 
</center><br>
<center>

<%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2
dim vValor_vend,vValor_vend1,vValor_vend2
dim vValor_comp,vValor_comp1,vValor_comp2
dim vCidade_vend,vCidade_comp
 
  vValor_vend=request.form("txt_valor_vend")
  
  if vValor_vend = "" then
  vValor_vend = request.QueryString("vValor_vend")
  end if
  
  
  session("vValor_vend") = vValor_vend
  
  session("vValor_vend1")=left(vValor_vend,10)
   session("vValor_vend2")=right(vValor_vend,10)
  
  
  
  vValor_comp=request.form("txt_valor_comp")
  
  if vValor_comp = "" then
  vValor_comp = request.QueryString("vValor_comp")
  end if
  
  
  
  session("vValor_comp") = vValor_comp
  
  
   session("vValor_comp1")=left(vValor_comp,10)
   session("vValor_comp2")=right(vValor_comp,10)
  
  
  
  
  
  vCidade_vend = request.form("select")
  
   if vCidade_vend = "" then
  vCidade_vend = request.QueryString("vCidade_vend")
  end if
  
  
  session("vCidade_vend") = vCidade_vend
  
 
  
  
  
  
  vCidade_comp = request.form("select3")
  
   if vCidade_comp = "" then
  vCidade_comp = request.QueryString("vCidade_comp")
  end if
  
  
  session("vCidade_comp") = vCidade_comp


 '---------------------------------------------------  
 
 
 '------------------------Sua Cidade--------------------------

 dim stringCidadeVend
  if session("vCidade_vend")<>"cqualquer" then
	stringCidadeVend = " where cidade_comp='"& session("vCidade_vend")&"'"
	else
	stringCidadeVend = ""
	end if
	
	
	
	
	
	
	'-------------------------------Valor do seu imóvel--------------------------------
	
	 dim stringValorVend
  if session("vValor_vend")<>"vqualquer" and session("vCidade_vend")<>"cqualquer" then
	stringValorVend = " and  Valor_comp >="& session("vValor_vend1") &" and Valor_comp <="& session("vValor_vend2") &""
	
	elseif session("vValor_vend")="vqualquer" and session("vCidade_vend")="cqualquer" then
	stringValorVend = ""
	
	elseif session("vValor_vend")<>"vqualquer" and session("vCidade_vend")="cqualquer" then
	 stringValorVend = " where Valor_comp >="& session("vValor_vend1") &" and Valor_comp <="& session("vValor_vend2") &""
	
	end if
	
	
	'-------------------------------Cidade pretendida--------------------------------
	
	 dim stringCidadeComp
  if session("vCidade_comp")<>"cqualquer" and session("vValor_vend")<>"vqualquer" and session("vCidade_vend")<>"cqualquer" then
	stringCidadeComp = " and cidade_vend='"& session("vCidade_comp")&"'"
	
	elseif session("vCidade_comp")="cqualquer" and session("vValor_vend")="vqualquer" and session("vCidade_vend")="cqualquer" then
     stringCidadeComp = ""
	
	
	 elseif session("vCidade_comp")<>"cqualquer" and session("vValor_vend")="vqualquer" and session("vCidade_vend")="cqualquer" then
	 stringCidadeComp = " where cidade_vend='"& session("vCidade_comp")&"'"
	 
	 elseif session("vCidade_comp")<>"cqualquer" and session("vValor_vend")<>"vqualquer" and session("vCidade_vend")="cqualquer" then
	 stringCidadeComp = " and cidade_vend='"& session("vCidade_comp")&"'"
	 
	 
	  elseif session("vCidade_comp")<>"cqualquer" and session("vValor_vend")="vqualquer" and session("vCidade_vend")<>"cqualquer" then
	 stringCidadeComp = " and cidade_vend='"& session("vCidade_comp")&"'"
	 
	 
	 
	end if
	
	
	
	
	'-------------------------------Valor pretendido--------------------------------
	
	
	 dim stringValorComp
  if  session("vValor_comp")<>"vqualquer" and session("vCidade_comp")<>"cqualquer" then
	stringValorComp = " and  Valor_vend >="& session("vValor_comp1") &" and Valor_vend <="& session("vValor_comp2") &""
	
	
	elseif session("vValor_comp")<>"vqualquer" and session("vCidade_comp")<>"cqualquer" and session("vCidade_vend")="cqualquer" and  session("vValor_vend")<>"vqualquer" then
	stringValorComp = " and  Valor_vend >="& session("vValor_comp1") &" and Valor_vend <="& session("vValor_comp2") &""
	
	
	elseif session("vValor_comp")<>"vqualquer" and session("vCidade_comp")="cqualquer" and session("vCidade_vend")<>"cqualquer" and  session("vValor_vend")="vqualquer" then
	stringValorComp = " and  Valor_vend >="& session("vValor_comp1") &" and Valor_vend <="& session("vValor_comp2") &""
	
	
	
	
	elseif session("vValor_comp")="vqualquer" and session("vCidade_comp")="cqualquer" then
	stringValorComp = "" 
	
	elseif session("vValor_comp")<>"vqualquer" and session("vCidade_comp")="cqualquer" and session("vValor_vend")="vqualquer" and session("vCidade_vend")="cqualquer" then
	
	stringValorComp = " where Valor_vend >="& session("vValor_comp1") &" and Valor_vend <="& session("vValor_comp2") &""
	
	end if
	
	
	
	
	
	
	
	
	strSQL = "SELECT * FROM permuta"&stringCidadeVend&stringValorVend&stringCidadeComp&stringValorComp
	
	
	
	
  
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
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
 Set Conexao2 = Server.CreateObject("ADODB.Connection")
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	Conexao2.Open dsn
	dim strSQL7
	
	
	



 dim Conexao22,rs77
 Set Conexao22 = Server.CreateObject("ADODB.Connection")
	Set rs77 = Server.CreateObject("ADODB.RecordSet")
	Conexao22.Open dsn
	dim strSQL77
	
	
	 strSQL77 = "SELECT * FROM imoveis where cod_imovel="&rs("cod_imovel")
	 rs77.CursorLocation = 3
      rs77.CursorType = 3
	 rs77.Open strSQL77, Conexao22
	 
	 dim vimagem
	 
   if not rs77.eof then
   vimagem = rs77("foto_grande")
   else
   vimagem = "imovel00000.jpg"
  end if
	
	
	 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	






%>
<% varCodPermuta =RS("cod_permuta") %>
 <tr>
            <td><table width="568" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="568" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
    <td width="568" height="153"><table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
          <td><table width="552" height="153" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                              <td width="552" height="16" bgcolor="FE9225"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Estou 
                                  interessado em im&oacute;vel na cidade de<strong> 
                                  <a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><font face="Verdana, arial" size="1" color="white"><%=RS("Cidade_comp")%></font></a> 
                                  </strong>no bairro <strong> </strong></font><font face="Verdana, arial" size="1" color="white"><strong><%=RS("Bairro_comp")%></strong></font></div></td>
              </tr>
              <tr> 
                              <td width="552" height="16" bgcolor="E17508"><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
                                  mais detalhes</strong></font></div></a></td>
              </tr>
              <tr> 
                <td><table width="552" height="115" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="173" bgcolor="FE9225"> 
                        <center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><%If objFSO.FileExists(Server.MapPath(vimagem)) = True Then%><img src="<%=vimagem%>" width="158" height="90" border=0></img><% else %><img src="imovel00000.jpg" width="158" height="90" border=0></img><% end if %></a></td>
                            </tr>
                          </table>
                        </center>
					  
					  </td>
                      <td bgcolor="FE9225"><div align="center"><font face="Verdana, arial" size="1" color="FFFFFF"><%=RS("descricao_vend")%></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
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
          <td><div align="center"><font face="Verdana, arial" size="1"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
               <a href="?page=<%=intPage - 1%>&vCidade_vend=<%=session("vCidade_vend")%>&vCidade_comp=<%=session("vCidade_comp")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_comp=<%=session("vValor_comp")%>">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" >
              
			  <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
			  
             
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
              <a href="?page=<%=intPage + 1%>&vCidade_vend=<%=session("vCidade_vend")%>&vCidade_comp=<%=session("vCidade_comp")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_comp=<%=session("vValor_comp")%>"><b>Próximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>
</center>



<%End If


Else

%>
  <% 
response.write session("vCidade_comp")&"2"&session("vCidade_vend")
%>
  <% end if %>
  <%


RS.close
Set RS = Nothing



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
group2[2][3]=new Option("200,00 até 500,00","0000000200 0000000500")
group2[2][4]=new Option("500,00 até 1000,00","0000000500 0000001000")
group2[2][5]=new Option("1000,00 até 2000,00","0000001000 0000002000")
group2[2][6]=new Option("Mais de 2000,00","0000002000 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Menos de 20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.000,00 até 50.000,00","0000020000 0000050000")
group2[3][4]=new Option("50.000,00 até 100.000,00","0000050000 0000100000")
group2[3][5]=new Option("100.000,00 até 200.000,00","0000100000 0000200000")
group2[3][6]=new Option("Mais de 200.000,00","0000200000 1000000000")









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


</body>
</html>
