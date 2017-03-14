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


</head>












<body bgcolor="EAA813" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo"  method="post" action="listar_imoveis.asp">

<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="755" height="78"><img src="top_page.jpg" width="755" height="78"></td>
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

<form name="doublecombo2" onSubmit="return isValidDigitNumber(this);" method="post" action="listar_referencia.asp">
  <table width="566" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
    <td width="243"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Procurar 
          im&oacute;vel por c&oacute;digo de refer&ecirc;ncia:<font color="EAA813">:</font></strong></font></div></td>
    <td width="149"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
      <input type="text" name="ref"  style="HEIGHT: 18px; WIDTH: 149px; ; font-size : 10px; background: F1991B; color:FFFFFF;">
      </strong></font></td>
    <td width="149"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
      <input name="image2" type="image"  src="bt_procurar001.jpg" width="149" height="18" border="0">
      </strong></font></td>
    <td>&nbsp;</td>
  </tr>
</table>
</form>











<table width="755" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="10" height="143">&nbsp;</td>
                      <td width="366" height="143">
					  
					  <table width="360" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><img src="../coral/mini_imovel00004.jpg" width="158" height="90"></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100">&nbsp;</td>
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
					  
					  
					  </td>
                      <td width="10" height="143">&nbsp;</td>
                      <td width="366" height="143">
					  
					  <table width="360" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="360" height="11"><img src="top_display.jpg" width="360" height="11"></td>
  </tr>
  <tr>
    <td width="360" height="116"><table width="360" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="7" height="116"><img src="left_display.jpg" width="7" height="116"></td>
          <td width="346" height="116"><table width="346" height="116" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="346" height="16" bgcolor="FE9225"></td>
              </tr>
              <tr>
                <td width="346" height="100" bgcolor="E17508"><table width="346" height="100" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="173" height="100" valign="bottom"><center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td><img src="../coral/mini_imovel00004.jpg" width="158" height="90"></td>
                            </tr>
                          </table>
                        </center>
                        </td>
                      <td width="173" height="100">&nbsp;</td>
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
					  
					  </td>
                      <td width="10" height="143">&nbsp;</td>
                    </tr>
                  </table>





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
  <% response.flush%>
  <%response.clear%>
</p>
</body>
</html>


</body>
</html>
