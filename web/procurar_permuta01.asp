<!--#include file="dsn.asp"-->
<%response.Buffer = true %>






<%
Function EscreveFuncaoJavaScript2 ( Conexao33 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (form) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (form.combo3.options[form.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas33 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas33 = Conexao33.Execute ( SqlMarcas33 )

While NOT rsMarcas33.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas33("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros33 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas33("id_combo1")&" order by nome_combo2"
Set rsCarros33 = Conexao33.Execute ( SqlCarros33 )

'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
While NOT rsCarros33.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%
'Criando conex�o com o banco de dados! 
Set Conexao33 = Server.CreateObject("ADODB.Connection")
Conexao33.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql33 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs33 = Conexao33.Execute ( Sql33 ) 
%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs44,strSQL44,Conexaoo
   Set Conexaoo = Server.CreateObject("ADODB.Connection")
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	strSQL44 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	Conexaoo.Open dsn
	
	rs44.Open strSQL44, Conexaoo



%>











<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
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

'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
While NOT rsCarros3.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%
'Criando conex�o com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 

dim Sql33 ,Rs33

Sql33 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs33 = Conexao3.Execute ( Sql33 ) 

dim Sql333 ,Rs333

Sql333 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs333 = Conexao3.Execute ( Sql333 ) 


%> 


<%

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
<%  EscreveFuncaoJavaScript2 ( Conexao33 ) %>

<script>
function isValidDigitNumber (doublecombo2){
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

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=590,height=510,resizable=no')
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
                  <option value="nqualquer">Negocia��o </option>
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
                  <option value="0000020000 0000050000">20.000,00 at� 50.000,00</option>
                  <option value="0000050000 0000100000">50.000,00 at� 100.000,00</option>
                  <option value="0000100000 0000200000">100.000,00 at� 200.000,00</option>
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
  <td width="755" height="10" bgcolor="863F15">
  <table width="755" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="176" height="10"></td>
              <td width="579" height="10"><table width="579" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="144"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="financiamento.asp" style="color:#FFCC00">Financiamento</a></strong></font></div></td>
                    <td width="144"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="dicas.asp" style="color:#FFCC00">Dicas</a></strong></font></div></td>
                    <td width="144"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('proposta_cliente.asp')" style="color:#FFCC00">Cadastre 
                        seu im&oacute;vel</a></strong></font></div></td>
                    <td width="144"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('form_enviar_email.asp')" style="color:#FFCC00">Contato</a></strong></font></div></td>
                  </tr>
                </table></td>
            </tr>
          </table>
  
  </td>
  </tr>
</table></form>


<center>
<font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Procure aqui im�veis para permuta</strong></font>

<br>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="procurar_permuta002.asp">Busca 
  detalhada</a></strong></font> 
</center>
<center>
 <form name="doublecombo"  method="post" action="listar_permuta01.asp">
  <table width="360" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="200"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Eu 
          moro na cidade de:<font color="EAA813">:</font></strong></font></div></td>
      <td><div align="center"> 
          <select name="select" class="inputBox" id="select" style="HEIGHT: 22px; WIDTH: 160px" onChange="javascript:atualizacarros2(this.form);">
            <option value="cqualquer" selected>Cidade</option>
            <% if not rs33.eof then %>
            <% While NOT Rs33.EoF %>
            <option value="<% = Rs33("nome_combo1") %>" > 
            <% = Rs33("nome_combo1") %>
            </option>
            <% Rs33.MoveNext %>
            <% Wend %>
            <option value="cqualquer">qualquer uma</option>
            <%else%>
            <option value=""></option>
            <%end if%>
          </select>
        </div></td>
    </tr>
    <tr> 
      <td width="200"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Meu 
          im&oacute;vel custa:<font color="EAA813">:</font></strong></font></div></td>
      <td><div align="center"><font color="#FFFFFF"> 
            <select name="txt_valor_vend" size="1"  class="inputBox" id="txt_valor_vend"  style="HEIGHT: 22px; WIDTH: 160px">
              <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000019999">menos de 20.000,00</option>
                  <option value="0000020000 0000050000">20.000,00 at� 50.000,00</option>
                  <option value="0000050000 0000100000">50.000,00 at� 100.000,00</option>
                  <option value="0000100000 0000150000">100.000,00 at� 150.000,00</option>
                   <option value="0000150000 0000200000">150.000,00 at� 200.000,00</option>
			        <option value="0000200000 0000250000">200.000,00 at� 250.000,00</option>
					<option value="0000250000 0000300000">250.000,00 at� 300.000,00</option>
					<option value="0000300000 0000350000">300.000,00 at� 350.000,00</option>
					<option value="0000350000 0000400000">350.000,00 at� 400.000,00</option>
			 <option value="0000400001 1000000000">acima de 400.000,00</option>
          </select>
          </font></div></td>
    </tr>
    <tr> 
      <td width="200"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Procuro 
          im&oacute;veis na cidade de:<font color="EAA813">:</font></strong></font></div></td>
      <td><div align="center"> 
          <select name="select3" class="inputBox" id="select3" style="HEIGHT: 22px; WIDTH: 160px" onChange="javascript:atualizacarros2(this.form);">
            <option value="cqualquer" selected>Cidade</option>
            <% if not rs333.eof then %>
            <% While NOT Rs333.EoF %>
            <option value="<% = Rs333("nome_combo1") %>" > 
            <% = Rs333("nome_combo1") %>
            </option>
            <% Rs333.MoveNext %>
            <% Wend %>
            <option value="cqualquer">qualquer uma</option>
            <%else%>
            <option value=""></option>
            <%end if%>
          </select>
        </div></td>
    </tr>
    <tr> 
      <td width="200"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>No 
          valor de:<font color="EAA813">:</font></strong></font></div></td>
      <td><div align="center"><font color="#FFFFFF"> 
            <select name="txt_valor_comp" size="1"  class="inputBox" id="txt_valor_comp"  style="HEIGHT: 22px; WIDTH: 160px">
              <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000019999">menos de 20.000,00</option>
                  <option value="0000020000 0000050000">20.000,00 at� 50.000,00</option>
                  <option value="0000050000 0000100000">50.000,00 at� 100.000,00</option>
                  <option value="0000100000 0000150000">100.000,00 at� 150.000,00</option>
                   <option value="0000150000 0000200000">150.000,00 at� 200.000,00</option>
			        <option value="0000200000 0000250000">200.000,00 at� 250.000,00</option>
					<option value="0000250000 0000300000">250.000,00 at� 300.000,00</option>
					<option value="0000300000 0000350000">300.000,00 at� 350.000,00</option>
					<option value="0000350000 0000400000">350.000,00 at� 400.000,00</option>
			 <option value="0000400001 1000000000">acima de 400.000,00</option>
          </select>
          </font></div></td>
    </tr>
    <tr> 
      <td width="200"> <div align="center"></div></td>
      <td><div align="center"> 
          <input name="image22" type="image"  src="bt_procurar001.jpg" width="160" height="18" border="0">
        </div></td>
    </tr>
  </table>
  </form>
</center>
<br>
<br>
  
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
group2[2][3]=new Option("200,00 at� 500,00","0000000200 0000000500")
group2[2][4]=new Option("500,00 at� 1000,00","0000000500 0000001000")

group2[2][5]=new Option("1000,00 at� 2000,00","0000001000 0000002000")
group2[2][6]=new Option("Mais de 2000,00","0000002000 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Menos de 20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.000,00 at� 50.000,00","0000020000 0000050000")
group2[3][4]=new Option("50.000,00 at� 100.000,00","0000050000 0000100000")
group2[3][5]=new Option("100.000,00 at� 200.000,00","0000100000 0000200000")
group2[3][6]=new Option("Mais de 200.000,00","0000200000 1000000000")









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
  <% response.flush%>
  <%response.clear%>
  <!--#include file="dsn2.asp"-->


</body>
</html>
