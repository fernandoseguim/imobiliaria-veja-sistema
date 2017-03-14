<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="loggedin.asp"-->

<%response.Buffer = true %>

<%






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
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas3 = Conexao3.Execute ( SqlMarcas3 )




While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"
Set rsCarros3 = Conexao3.Execute ( SqlCarros3 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro" & "','" & "bqualquer" & "');"& vbcrlf
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




'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT * FROM tipo  ORDER BY tipo ASC" 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao









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














<body bgcolor="<%=escuro%>">
<br><center>
<table width="250" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Selecione 
          abaixo as caracter&iacute;sticas do im&oacute;vel e clique no bot&atilde;o 
          &quot;AVALIAR&quot; para saber o valor de mercado do im&oacute;vel.</strong></font> 
        </div></td>
  </tr>
</table>

<br>
<br>
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"  method="post" action="listar_avaliacao_corretor.asp?nome=<%=session("nome")%>&email=<%=session("email")%>&telefone=<%=session("telefone")%>">
  <table width="250" height="120" border="0" cellspacing="0" cellpadding="0">
    <tr>
     
      <td width="250" height="120"><table width="250" border="0"  cellspacing="0" cellpadding="0" height="120">
          
							
							
							
							
							
							<tr>
                                  <td>
								  <select name="combo1" onChange="javascript:atualizacarros(this.form);"  style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#00000; ">
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
                                  <td><select name="combo2"  onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#00000; ">
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
                                  <td><select name="txt_tipo" size="1"  style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#00000; ">
                  <option value="tqualquer" selected>Tipo do  imóvel</option>
				   
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
                                  <td><select name="txt_Quartos" size="1"   style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#00000; ">
                  <option value="qqualquer" selected>Número de quartos do imóvel</option>
				   
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
                                  <td><select name="txt_garagem" size="1"   style="HEIGHT: 18px; WIDTH: 250px; font-size : 10px;  color:#00000; ">
                  <option value="gqualquer" selected>Número de vagas na garagem do imóvel</option>
				   
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
                                  
              <td><div align="right"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Metragem 
                  do im&oacute;vel:</strong></font> 
                  <input name="txt_area_total" type="text" class="inputBox" id="txt_area_total" style="HEIGHT: 18px; WIDTH: 80px; background:#FFFFFF" value="00" size="12" maxlength="20">
                  <font color="#FFFFFF"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  m&sup2;</font> </font></div></td>
                            </tr>
							
							
							
							
							<tr>
							<td width="250" height="20"><div align="right">
                  <input name="image2" type="image"  src="bt_avaliar004.jpg" width="149" height="18" border="0">
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
  <% response.flush%>
  <%response.clear%>
</p>
<%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<% EscreveFuncaoJavaScript222 ( Conexao3 ) %>
</body>
</html>
