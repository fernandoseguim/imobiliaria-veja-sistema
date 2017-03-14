<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->

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


dim varCidade,stringCidade,varBairro,stringBairro,varNegociacao
dim stringNegociacao,varQuartos,stringQuartos,varCidade2

 varCidade2 = request.querystring("combo1")
 if varCidade2 = "" then
 varCidade2 = "cqualquer"
 end if
 
session("varCidade2") = varCidade2
 
 if varCidade2 <> "cqualquer" then
 dim rrs2,SSQL2


Set Conexao4 = Server.CreateObject("ADODB.Connection")
Conexao4.open dsn

 Set rrs2 = Server.CreateObject("ADODB.RecordSet")
 SSQL2 = "select * from combo1 where id_combo1="&varCidade2
 
 rrs2.open SSQL2,Conexao4,2,1
 
 varCidade = rrs2("nome_combo1")
 else
 varCidade = varCidade2
 end if
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
varBairro2 = request.querystring("combo2")
if varBairro2 = "" then
varBairro2 = "bqualquer"
end if

 if varBairro2 <> "bqualquer" then
	  dim rrs3,SSQL3,conexao5
	  Set Conexao5 = Server.CreateObject("ADODB.Connection")
	  Conexao5.open dsn
 Set rrs3 = Server.CreateObject("ADODB.RecordSet")
 SSQL3 = "select * from combo2 where id_combo2 ="&varBairro2
 
 rrs3.open SSQL3,Conexao5,2,1

 varBairro = rrs3("nome_combo2")
 else
 varBairro = varBairro2
	end if                                      
									
	 




varNegociacao = request.querystring("example2")
varQuartos = request.querystring("txt_quartos")

	




dim varNotFind



dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao



%>




<html>
<head>
<title></title>

<script>

function check(acao){
if(document.Formulario.selTodos.checked){
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked = acao;
}
}
else
{
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked =! acao;
}
}



}





</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>


<%  EscreveFuncaoJavaScript ( Conexao3 ) %>



</head>

<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="#FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<table width="790" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr bgcolor="<%=claro%>"> 
   
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_inicial.asp">Im&oacute;veis</a></strong></font></div></td>
    
	  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_inicial.asp">Proposta</a></strong></font></div></td>
     
    
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>" > 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_inicial.asp">Email</a></strong></font></div></td>
   
   
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp">Cidades</a></strong></font></div></td>
  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro_inicial.asp">Bairros</a></strong></font></div></td>
  
  <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_vila.asp">Vila</a></strong></font></div></td>
 
  
  
  <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_inicial.asp">Compradores</a></strong></font></div></td>
 
  
  
   <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_inicial.asp">Permuta</a></strong></font></div></td>
 
  
  </tr>
</table>

<center>
<br>

<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>A sua permissão é <%=session("permissao")%></strong></font>
<br>

<table width="750" height="18" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="18" bgcolor="#FFFFFF"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="#FF0000"></td>
          </tr>
        </table></td>
      <td width="160"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
        <font color="#FFFFFF">s</font><font color="<%=escuro%>">Primeira página 
        e standBy</font></strong></font></td>
      <td width="18" bgcolor="#FFFFFF"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="1956C6"></td>
          </tr>
        </table></td>
      <td width="60"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FFFFFF">s</font><font color="<%=escuro%>">StandBy</font></strong></font></td>
      <td width="18" bgcolor="#FFFFFF"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="1F3955"></td>
          </tr>
        </table></td>
      <td width="100"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FFFFFF">s</font><font color="<%=escuro%>">Primeira 
        p&aacute;gina</font></strong></font></td>
		
		<td width="18" bgcolor="#FFFFFF"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="green"></td>
          </tr>
        </table></td>
      <td width="200"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FFFFFF">s</font><font color="<%=escuro%>">Primeira 
        página e bom neg&oacute;cio</font></strong></font></td>
		<td width="18" bgcolor="#FFFFFF"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="#06D915"></td>
          </tr>
        </table></td>
      <td width="80"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FFFFFF">s</font><font color="<%=escuro%>">Bom 
        neg&oacute;cio </font></strong></font></td>
  </tr>
 
</table>

<br>
<table width="750" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="370" bgcolor="#FFFFFF"><% if  session("permissao") = "4" or  session("permissao") = "3" or  session("permissao") = "5"   then %><div align="center"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="estatistica_imoveis.asp" style="color:black">Estat&iacute;stica da busca de im&oacute;veis</a></strong></font></div><%else%><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estat&iacute;stica da busca de im&oacute;veis</strong></font></div><%end if%></td>
      <td bgcolor="#FFFFFF">
        <% if  session("permissao") = "4" or  session("permissao") = "3" or  session("permissao") = "5"   then %>
        <div align="center"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="estatistica_referencia.asp" style="color:black">Estat&iacute;stica 
          da busca por refer&ecirc;ncia</a></strong></font></div>
        <%else%>
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estat&iacute;stica 
          da busca por refer&ecirc;ncia</strong></font></div>
        <%end if%></td>
    </tr>
  </table>


<br>
</center>


<center>
<form name="doublecombo"  method="GET" action="archive_imoveis.asp">
  <table width="640" height="26" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=claro%>">
    <tr>
    <td width="120" align="right"><select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 16px; WIDTH: 115px; background:<%=medio%>" >
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
      </select></td>
    <td width="120"><font color="#FFFFFF">
      <select name="combo2" class="inputBox" style="HEIGHT: 16px; WIDTH: 122px; background:<%=medio%>">
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
      </select>
      </font></td>
        <td width="120"><select name="example2" size="1" class="inputBox" id="select3" onChange="redirect2(this.options.selectedIndex)" style="HEIGHT: 16px; WIDTH: 120px; background:<%=medio%>">
            <option value="nqualquer">Negociação </option>
            <option value="nqualquer" >Qualquer um </option>
            <option  value="Aluguel">Aluguel </option>
            <option value="Venda">Venda </option>
          </select></td>
	  <td width="100"><select name="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 16px; WIDTH: 100px; background:<%=medio%>">
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
    <td width="70"><select name="txt_quartos" id="txt_quartos" class="inputBox" style="HEIGHT: 16px; WIDTH: 70px; background:<%=medio%>" >
        <option value="qqualquer" selected>Quartos</option>
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
	  <td width="70"><select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 16px; WIDTH: 70px; background:<%=medio%>" >
        <option value="vgqualquer" selected>Vagas</option>
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
	  <td width="120">
	  <select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 16px; WIDTH: 160px; background:<%=medio%>">
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
			   
			   
			    </select>
	  
	  </td>
	  <td><select name="txt_foto" size="1"  class="inputBox" style="HEIGHT: 16px; WIDTH: 100px; background:<%=medio%>">
                  
				  <option value="fqualquer">Foto</option>
				   <option value="Com Foto">Com Foto</option>
                  <option value="Sem Foto">Sem Foto</option>
				   <option value="fqualquer">qualquer um</option>
				 
                  
                 
                </select></td>
	  
        <td width="65" bgcolor="<%=claro%>"> 
          <input name="submit" type="submit" class="inputSubmit" style="background:<%=medio%>;" value="Buscar" width="80">
        </td>
  </tr>
</table>
</form>
</center>




<form action="archive_imoveis.asp?" Method="GET" name="b2" >

<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td bgcolor="#DAE3F0">
<table border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="<%=claro%>">
          <tr>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>Procurar</b></font></td>
            <td bgcolor="<%=claro%>">
<input type="text" name="SearchFor" class="inputBox" value="" style="background:<%=medio%>">
            </td>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>em</b></font></td>
            <td bgcolor="<%=claro%>">
	
	
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>">
<option value="proprietario" selected >Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone">Telefone</option>
<option value="Data" >Data</option>
<option value="Primeira" >Primeira Página</option>
<option value="StandBy" >StandBy</option>
<option value="cod" >Código do imóvel</option>
<option value="captador" selected>Captador</option>
</select>






















            </td>
            <td bgcolor="<%=claro%>"> 
              <input type="submit" value="Buscar" class="inputSubmit" style="background:<%=medio%>">
            </td>
</tr>
</table>
</td>
</tr>
</table>
</form>
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


</body>
</html>
