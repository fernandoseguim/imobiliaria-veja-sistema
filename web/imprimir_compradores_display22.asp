



<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
 
 
 dim varCodCompradores
	varCodCompradores=request.QueryString("varCodCompradores")
	
	
	dim Conexao9
	dim rs9
	dim strSQL9
	 Set Conexao9 = Server.CreateObject("ADODB.Connection")
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	
	
	Conexao9.open dsn
	 strSQL9 = "SELECT * FROM compradores where cod_compradores="&varCodCompradores
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	  
	  
	 rs9.Open strSQL9, Conexao9
 
 
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
   dim rs4,strSQL4
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		rs4.Open strSQL4, Conexao
		
	
	
dim rs666,strSQL666
   
    Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL666 = "SELECT * FROM combo1 where nome_combo1 ='"&rs9("cidade")&"'  ORDER BY nome_combo1" 
	 rs666.Open strSQL666, Conexao		

	 
	 
	 dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT * FROM combo2 where nome_combo2 ='"&rs9("bairro")&"' and cidade_combo2 ='"&rs9("cidade")&"'  ORDER BY nome_combo2" 
	 rs555.Open strSQL555, Conexao		

dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT * FROM combo3 where nome_combo3 ='"&rs9("vila")&"' and cidade_combo3 ='"&rs9("cidade")&"' and bairro_combo3 ='"&rs9("bairro")&"'  ORDER BY nome_combo3" 
	 rs444.Open strSQL444, Conexao


dim rs444Placa,strSQL444Placa
   
    Set rs444Placa = Server.CreateObject("ADODB.RecordSet")
	strSQL444Placa = "SELECT * FROM senha  ORDER BY id Desc" 
	 rs444Placa.Open strSQL444Placa, Conexao		


	 dim rs444Captacao,strSQL444Captacao
   
    Set rs444Captacao = Server.CreateObject("ADODB.RecordSet")
	strSQL444Captacao = "SELECT * FROM senha  ORDER BY id Desc" 
	 rs444Captacao.Open strSQL444Captacao, Conexao	
%>		


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>



<html>


<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>




<script>
function isValidDigitNumber (doublecombo)
{






if (doublecombo.txt_proprietario.value == "") {
        alert("Você precisa indicar o nome do comprador!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do comprador!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
    }
	
	
	
	
	
var strValidNumber1_7="1234567890,";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O telefone do comprador só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}





	var strValidNumber1_6="1234567890,.";
for (nCount=0; nCount < doublecombo.stage22.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.stage22.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.stage22.focus();
doublecombo.stage22.select();
return false;
}
}


var strValidNumber1_7="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_valor_vend.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_valor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor_vend.focus();
doublecombo.txt_valor_vend.select();
return false;
}
}




var strValidNumber1_8="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_condominio_vend.value.length; nCount++) 
		{
strTempChar1_8=doublecombo.txt_condominio_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_8.indexOf(strTempChar1_8,0)==-1) 
{
alert("O formulário Condomínio só pode conter números!");
doublecombo.txt_condominio_vend.focus();
doublecombo.txt_condominio_vend.select();
return false;
}
}





	
	if (doublecombo.stage22.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }


var strText2_4 = doublecombo.stage22.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.stage22.focus();
		
		doublecombo.stage22.select();
		
return false;

}





var elem=doublecombo.elements;

for (nCount=0; nCount < elem.length; nCount++)
  

	if(elem[nCount].type.indexOf("text")==0)	{
	var strValidNumber12_1="'";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)!=-1) {



alert("Este campo  não pode conter aspas");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}








}



</script>



</head>

<!--#include file="style_imprimir.asp"-->


<body onload=doublecombo.txt_atendimento.focus(); bgcolor="#FFFFFF" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >

<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="indicacao_display22.asp?varCodCompradores=<%=varCodCompradores%>" style="color:#000000">Voltar</a></strong></font></div></td>
  </tr>
  <tr>
      <td height="18">
<div align="center"> 
          <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi atualizado com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
              
			   <tr> 
                  <td style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      de refer&ecirc;ncia do comprador</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_referencia" type="text" class="inputBox" id="txt_referencia" style="HEIGHT: 18px; WIDTH: 290px;" value="<%=rs9("cod_compradores")%>" size="38" maxlength="35" align="left"></td>
                </tr>
			  
			    <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Acessos</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_acessos" type="text" class="inputBox" id="txt_acessos" style="HEIGHT: 18px; WIDTH: 290px;" value="<%if rs9("acessos") <> "" then response.write rs9("acessos") else response.write "0" end if%>" size="38" maxlength="35" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo atendimento</font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_data2" type="text" class="inputBox" id="txt_data2" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de inclus&atilde;o</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("data")%>" size="38" maxlength="50" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de atualiza&ccedil;&atilde;o</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("data_atualizacao")%>" size="38" maxlength="50" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do interessado</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("nome")%>" size="38" maxlength="35" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      do interessado</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_telefone" type="<%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6" then %><% if  UCase(rs9("atendimento")) <> UCase(Session("Admin_ID")) then response.write "Hidden" else response.write "text" end if %><%else%><%response.write "text" end if %>" class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 290px;" value="<%=rs9("telefone")%>" size="38" maxlength="20" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                      do interessado</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_email" type="<%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6" then %><% if  UCase(rs9("atendimento")) <> UCase(Session("Admin_ID")) then response.write "Hidden" else response.write "text" end if %><%else%><%response.write "text" end if %>" class="inputBox" id="txt_email" style="HEIGHT: 18px; WIDTH: 290px ; " value="<%if rs9("email") = "não informado" then response.write "" else response.write rs9("email") end if%>" size="38" maxlength="50" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("cidade")%>" size="38" maxlength="50" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("bairro")%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("vila")%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel desejado</font></div></td>
                  <td  style="border:1px solid #000000;"><font color="#FFFFFF"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("tipo")%>" size="38" maxlength="50" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de quartos do im&oacute;vel desejado</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("quartos")%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem do im&oacute;vel desejado</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("vagas")%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o 
                      do im&oacute;vel desejado</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("ocupacao")%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      que deseja</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("negociacao")%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy</font></div></td>
                  <td style="border:1px solid #000000;"> <input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs9("standby") <> "" then response.write rs9("standby") else response.write "excluido" end if%>" size="38" maxlength="50" align="left"> </td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Faixa 
                      de pre&ccedil;o desejada</font></div></td>
                  <td  style="border:1px solid #000000;"><input name="stage22" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; " value="<%=formatnumber(rs9("valor"),2)%>" size="12" maxlength="30"> 
                  </td>
                </tr>
                <tr> 
                  <td width="290" height="100" style="border:1px solid #000000;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="290" height="18"  style="border-bottom: 2px solid #000000;"> 
                          <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            do im&oacute;vel desejado e forma de pagamento</font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="82"  >&nbsp;</td>
                      </tr>
                    </table></td>
                  <td width="290" height="100"  style="border:1px solid #000000;" ><textarea name="txt_descricao" COLS=20 ROWS=10 class="inputBox" id="txt_descricao" style="HEIGHT: 100px; WIDTH: 290px; " onKeyPress="return limitfield(this, 600)"><%=rs9("descricao")%></textarea></td>
                </tr>
                <tr> 
                  <td width="290" height="100" style="border:1px solid #000000;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="290" height="18"  style="border-bottom: 2px solid #000000;"> 
                          <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            confidencial do interessado</font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="82"  >&nbsp;</td>
                      </tr>
                    </table></td>
                  <td width="290" height="100"  style="border:1px solid #000000;" ><textarea name="txt_descricao_confi" COLS=20 ROWS=10 class="inputBox" id="txt_descricao_confi" style="HEIGHT: 100px; WIDTH: 290px; " onKeyPress="return limitfield(this, 800)"><%=rs9("descricao_confi")%></textarea></td>
                </tr>
                <tr> 
                  <td > 
                    <div align="center"></div></td>
                  <td ><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="145">&nbsp;</td>
                        
                      <td width="145"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="" onclick="javascript:print();return false;" style="color:#000000">Imprimir</a></strong></font></div></td>
                      </tr>
                    </table></td>
                </tr>
                
              </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
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
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>

