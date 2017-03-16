<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->

<%response.Buffer = true %>

<%






%>




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
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 


Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsMarcas3.ActiveConnection = Conexao3
	
	
	rsMarcas3.Open SqlMarcas3, Conexao3



While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"



Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsCarros3.ActiveConnection = Conexao3
	
	
	rsCarros3.Open SqlCarros3, Conexao3



'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Seu Bairro" & "','" & "bqualquer" & "');"& vbcrlf
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


rsMarcas3.close

set rsMarcas3 = nothing


rsCarros3.close

set rsCarros3 = nothing



End Function
%> 


<%
'Criando conex�o com o banco de dados! 
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
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao




'------------------------------selecionar os tipos de im�vel para o formul�rio-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT * FROM tipo  ORDER BY tipo ASC" 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao









'-------------------------------------------------------------------------------------------------






%>





<html>

<!--#include file="style_imoveis02.asp"-->

<title>Avalia��o</title>

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

}
}


</script>


</head>














<body bgcolor="#FFFFFF">
<br><center>
<table width="250" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sendo 
          o sr(a) realmente propriet&aacute;rio de um im&oacute;vel e tendo o 
          interesse em avali&aacute;-lo, preencha os campos abaixo, n&atilde;o 
          tendo o interesse, clique em n&atilde;o avaliar.<br>
          <br><br>
          <font size="3">E depois fique a vontade para navegar no site.</font></strong></font> 
        </div></td>
  </tr>
</table>

<br>
<br>
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"  method="post" action="listar_avaliacao02.asp">
  <table width="250" height="120" border="0" cellspacing="0" cellpadding="0">
    <tr>
     
      <td width="250" height="120"><table width="250" border="0"  cellspacing="0" cellpadding="0" height="184">
            <tr><td height="18"><div align="center">
                  <input name="nome" onFocus="doublecombo.nome.value=''" type="Text" class="inputBox" id="nome" style="color:#9d9249;HEIGHT: 18px; WIDTH: 250px; background:<%=claro%>" value="Seu nome:" size="15" >
                </div></td></tr>
							
							
							<tr>
              <td height="20">
<div align="center">
                  <input name="telefone" onFocus="doublecombo.telefone.value=''" type="Text" class="inputBox" id="telefone" style="color:#9d9249;HEIGHT: 18px; WIDTH: 250px; background:<%=claro%>" value="Seu telefone:" size="15" >
                </div></td></tr>
							
							
							<tr><td height="18"><div align="center">
                  <input name="email" onFocus="doublecombo.email.value=''" type="Text" class="inputBox" id="email" style="color:#9d9249;HEIGHT: 18px; WIDTH: 250px; background:<%=claro%>" value="Seu email:" size="15" >
                </div></td></tr>
							
							<tr>
                                  <td>
								  <select name="combo1" onChange="javascript:atualizacarros(this.form);"  style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#9d9249;background:<%=claro%>; ">
                  <option value="cqualquer" selected>Sua Cidade</option>
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
                                  <td><select name="combo2"  onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#9d9249;background:<%=claro%>; ">
                   <option value="bqualquer" selected>Seu Bairro</option>
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
                                  <td><select name="txt_tipo" size="1"  style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#9d9249;background:<%=claro%>; ">
                  <option value="tqualquer" selected>Tipo do seu im�vel</option>
				   
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
                                  <td><select name="txt_Quartos" size="1"   style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#9d9249;background:<%=claro%>; ">
                  <option value="0" selected>N�mero de quartos do seu im�vel</option>
				   
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
                                  <td><select name="txt_garagem" size="1"   style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#9d9249;background:<%=claro%>; ">
                  <option value="0" selected>N�mero de vagas na garagem do seu im�vel</option>
				   
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
                                  
              <td><div align="right"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Metragem 
                  do seu im&oacute;vel:</strong></font> 
                  <input name="txt_area_total" type="text" class="inputBox" id="txt_area_total" style="border-color:#9d9249;HEIGHT: 18px; WIDTH: 80px;color:#9d9249;background:<%=claro%>;" value="00" size="12" maxlength="20">
                  <font color="#FFFFFF"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  m&sup2;</font> </font></div></td>
                            </tr>
							
							
							
							
							<tr>
							<td width="250" height="20"><div align="right">
                  <table width="250" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="18" width="125"><a href="primeira01.asp" target="_self"><img src="bt_avaliar002.jpg" border="0"></a></img></td>
                      <td height="18" width="125"><input name="image2" type="image"  src="bt_avaliar003.jpg" width="125" height="18" border="0"></td>
                    </tr>
                  </table>
                </div></td>
							
							</tr>
							
							
							
                           
                           
                           
                            
                          </table></td>
    </tr>
  </table>
</form>


</center>
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
  <% response.flush%>
  <%response.clear%>
</p>
<%  EscreveFuncaoJavaScript ( Conexao3 ) %>

</body>
</html>
