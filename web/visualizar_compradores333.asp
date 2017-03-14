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

Response.Write "form.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
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


Set Conexao33 = Server.CreateObject("ADODB.Connection")
Conexao33.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")


'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")






'Abrindo a tabela MARCAS!
Sql33 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs33 = Conexao33.Execute ( Sql33 )

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 


<%
Function EscreveFuncaoJavaScript2 ( Conexao33 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo3.options[form.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
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

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
 
While NOT rsCarros33.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend

Response.Write "form.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




<%
Function EscreveFuncaoJavaScript888 ( Conexao333 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros888 (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo2.options[form.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas888 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas888 = Conexao333.Execute ( SqlMarcas888 )

While NOT rsMarcas888.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas888("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros888 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas888("id_combo2")&""

Set rsCarros888 = Conexao333.Execute ( SqlCarros888 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros888.EoF

Response.Write "form.combo5.options[" & i & "] = new Option('" & rsCarros888("nome_combo3") & "','" & rsCarros888("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros888.MoveNext
Wend


Response.Write "form.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas888.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%



Sql888 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs888 = Conexao333.Execute ( Sql888 ) 




%> 




<%
Function EscreveFuncaoJavaScript999 ( Conexao333 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros999 (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo4.options[form.combo4.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas999 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas999 = Conexao333.Execute ( SqlMarcas999 )

While NOT rsMarcas999.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""

Set rsCarros999 = Conexao333.Execute ( SqlCarros999 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.combo7.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros999.EoF

Response.Write "form.combo7.options[" & i & "] = new Option('" & rsCarros999("nome_combo3") & "','" & rsCarros999("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros999.MoveNext
Wend


Response.Write "form.combo7.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas999.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%



Sql999 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs999 = Conexao333.Execute ( Sql999 ) 




%> 




<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
	 dim rs4,strSQL4
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
	 dim rs44,strSQL44
   
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	strSQL44 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2"
	
	
	
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		rs4.Open strSQL4, Conexao
		rs44.Open strSQL44, Conexao
%>		




<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>





<script>
function isValidDigitNumber (doublecombo)
{
{



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



{
if (doublecombo.txt_email.value == "") {
		
	} else {
		prim = doublecombo.txt_email.value.indexOf("@")
		if(prim < 2) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("@",prim + 1) != -1) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(".") < 1) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(" ") != -1) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("zipmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("hotmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(".@") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("@.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(".com.br.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("/") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("[") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("]") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("(") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(")") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("..") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		
		
		
		}
		
		
	}

}










var strValidNumber1_7="1234567890,";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Telefone só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}









	
	if (doublecombo.txt_proprietario.value == "") {
        alert("O formulário Proprietário do Imóvel está vazio!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	if (doublecombo.combo2.value == "") {
        alert("O formulário Bairro do Imóvel está vazio!");
        doublecombo.combo1.focus();
		
        return false;
    }
	
	if (doublecombo.combo1.value == "") {
        alert("O formulário Cidade do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }
	
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("O formulário telefone está vazio!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
    }
	
	
	
	
	if (doublecombo.txt_endereco.value == "") {
        alert("O formulário Endereço do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }
	
	
	
	
	
	
	
	






	
	
	
	
	
	
		var strValidNumber1_5="1234567890,";
for (nCount=0; nCount < doublecombo.txt_a_total.value.length; nCount++) 
		{
strTempChar1_5=doublecombo.txt_a_total.value.substring(nCount,nCount+1);
if (strValidNumber1_5.indexOf(strTempChar1_5,0)==-1) 
{
alert("O formulário Área Total só pode conter números!");
doublecombo.txt_a_total.focus();
doublecombo.txt_a_total.select();
return false;
}
}
	
	

	

var strValidNumber1_4="1234567890,";
for (nCount=0; nCount < doublecombo.txt_a_constr.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_a_constr.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("O formulário Área Construída só pode conter números!");
doublecombo.txt_a_constr.focus();
doublecombo.txt_a_constr.select();
return false;
}
}



if (doublecombo.txt_valor.value == "") {
        alert("O formulário Valor está vazio!");
        doublecombo.txt_valor.focus();
		doublecombo.txt_valor.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_6="1234567890,";
for (nCount=0; nCount < doublecombo.txt_valor.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.txt_valor.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor.focus();
doublecombo.txt_valor.select();
return false;
}
}

var strText2_4 = doublecombo.txt_valor.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.txt_valor.focus();
		
		doublecombo.txt_valor.select();
		
return false;

}
//-----------


//----------------------

prim2_4 = doublecombo.txt_valor.value.indexOf(",")
if(doublecombo.txt_valor.value.indexOf(",",prim2_4 + 1) != -1) {
			alert("O formulário Valor não contêm a vírgula do valor-moeda");
			doublecombo.txt_valor.focus();
			doublecombo.txt_valor.select();
			return false;
		}







	
	
	
   
	
	

	
	
}



{







//------------- Verifica se é numérico---------------------



//-----------------------------------------------

}
}









</script>



<html>


<head><%
EscreveFuncaoJavaScript2 ( Conexao33 )
EscreveFuncaoJavaScript ( Conexao3 )
  EscreveFuncaoJavaScript888 ( Conexao333 ) 
 EscreveFuncaoJavaScript999 ( Conexao333 ) 

 %>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>
<title>Encomenda de imóvel</title>
</head>

<!--#include file="style_imoveis.asp"-->

<body onload=doublecombo.txt_telefone.focus(); bgcolor="<%=escuro%>" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="incluir_compradores03.asp">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  <tr>
    <td width="590" height="60"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Este 
                formul&aacute;rio deve ser preenchido caso voc&ecirc; queira comprar 
                um im&oacute;vel, informe abaixo todos os dados do im&oacute;vel 
                que voc&ecirc; pretende comprar e as informa&ccedil;&otilde;es 
                ficaram dispon&iacute;veis para os internautas. Al&eacute;m disso 
                ser&aacute; criada uma conta de comprador de im&oacute;vel para 
                voc&ecirc;, nessa conta, nosso sistema vai indicar v&aacute;rios 
                im&oacute;veis dispon&iacute;veis que correspondem ao seu pedido. 
                &quot;OBS: As conta s&atilde;o gratuitas&quot;</strong></font></div></td>
  </tr>
  
  <tr>
    <td width="590" height="18"><div align="center">
        <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
        <%else%>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
        foi incluido com sucesso.</font> 
        <% end if %>
      </div></td>
  </tr>
</table>

<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="5">&nbsp;</td>
    <td><table width="580" border="0" cellspacing="0" cellpadding="0">
        
        <tr>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                nome </font></div></td>
          <td style="border:1px solid #FFFFFF; background:<%=claro%>" ><input name="txt_proprietario" onfocus="doublecombo.txt_proprietario.value=''" value="Clique aqui e comece a digitar...." type="text" id="txt_proprietario" size="38" maxlength="35" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
        </tr>
        <tr>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>" >
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                telefone </font></div></td>
          <td style="border:1px solid #FFFFFF; background:<%=medio%>"><input name="txt_telefone" type="text" id="txt_telefone" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
        </tr>
		<tr>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                email </font></div></td>
          <td style="border:1px solid #FFFFFF; background:<%=claro%>"><input name="txt_email" type="text" id="txt_email" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>"></td>
        </tr>
      </table></td>
    <td width="5">&nbsp;</td>
  </tr>
</table>

<div align="center"><br>
    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Coloque 
    abaixo os dados do im&oacute;vel que voc&ecirc; pretende comprar e as formas 
    de pagamento que voc&ecirc; deseja.</font></strong></div>
<br>
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="5">&nbsp;</td>
    <td><div align="center"> 
          <table width="580" border="0" cellspacing="0" cellpadding="0">
            <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                  onde quer comprar ou alugar im&oacute;vel</font></div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"> <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
               <option value="cqualquer" selected>Cidade</option>
			    <% if not rs3.eof then %>
                <% While NOT Rs3.EoF %>
                <option value="<% = Rs3("id_combo1") %>"> 
                <% = Rs3("nome_combo1") %>
                </option>
                <% Rs3.MoveNext %>
                <% Wend %>
                <%else%>
                <option value=""></option>
                <%end if%>
              </select>
              </td>
          </tr>
          <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                  onde quer comprar ou alugar im&oacute;vel</font></div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"> <select name="combo2" class="inputBox"  onChange="javascript:atualizacarros888(this.form);" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                <option value="bqualquer" selected>Bairro/Região</option>
				<% if not rs4.eof then%>
                <% While NOT Rs4.EoF %>
                <option value="<% = Rs4("id_combo2") %>"> 
                <% = Rs4("nome_combo2") %>
                </option>
                <% Rs4.MoveNext %>
                <% Wend %>
                <% else %>
                <option value=""></option>
                <% end if %>
              </select>
              </td>
          </tr>
		  
		  
		  <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  do seu im&oacute;vel </font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                   <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                  
                </select></td>
            </tr>
			
			
		  
		  
		  
		  
		  
          <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                  do im&oacute;vel desejado</font></div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><font color="#FFFFFF">
              <select name="txt_tipo" size="1" id="txt_tipo" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
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
              </select>
              </font></td>
          
          
          <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                  de quartos do im&oacute;vel desejado</font></div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><select name="txt_quartos" id="txt_quartos" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                <option value="não informado" selected>não informado</option>
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
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="não informado" selected>não informado</option>
                    <option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
              
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="não informado" selected>não informado</option>
                    <option value="ocupado">Ocupado</option>
                    <option value="vago">Vago</option>
				  </select>
                  </td>
              </tr>
              
         
          <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                  que deseja</font></div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><select name="txt_negociacao" id="txt_negociacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                <option value="compra" selected>Compra</option>
			    <option value="aluguel">Aluguel</option>               
              </select></td>
          </tr>
         
          <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                  que voc&ecirc; pretende pagar</font></div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><input name="txt_valor" type="text" class="inputBox" id="txt_valor" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="20"></td>
          </tr>
          <tr> 
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><div align="center">
                  <table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                          do im&oacute;vel desejado e forma de pagamento</font></div></td>
                    </tr>
                    <tr>
                      <td width="290" height="82" bgcolor="<%=medio%>">&nbsp;</td>
                    </tr>
                  </table>
                </div></td>
            <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="HEIGHT: 100px; WIDTH: 290px; background: <%=claro%>" onKeyPress="return limitfield(this, 200)"></textarea></td>
          </tr>
		  
        </table>
      </div></td>
    <td width="5">&nbsp;</td>
  </tr>
</table>

<div align="center"><br>
          <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Caso 
          voc&ecirc; tenha um im&oacute;vel para dar como parte de pagamento, 
          informe abaixo:</font></strong></div>
<br>
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="5">&nbsp;</td>
    <td><div align="center">
          <table width="580" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Voc&ecirc; 
                      aceita dar im&oacute;vel como parte de pagamento?</font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><select name="txt_pergunta" id="select2" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                  <option value="sim">Sim</option>
                  <option value="nao" selected>Não</option>
                </select></td>
            </tr>
            <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      do seu im&oacute;vel </font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" onChange="javascript:atualizacarros2(this.form);">
                  <option value="cqualquer" selected>Cidade</option>
				  <% if not rs33.eof then %>
                  <% While NOT Rs33.EoF %>
                  <option value="<% = Rs33("id_combo1") %>"> 
                  <% = Rs33("nome_combo1") %>
                  </option>
                  <% Rs33.MoveNext %>
                  <% Wend %>
                  <%else%>
                  <option value=""></option>
                  <%end if%>
                </select></td>
            </tr>
            <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      do seu im&oacute;vel</font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><select name="combo4" onChange="javascript:atualizacarros999(this.form);" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                 <option value="bqualquer" selected>Bairro/Região</option>
				  <% if not rs44.eof then%>
                  <% While NOT Rs44.EoF %>
                  <option value="<% = Rs44("id_combo2") %>"> 
                  <% = Rs44("nome_combo2") %>
                  </option>
                  <% Rs44.MoveNext %>
                  <% Wend %>
                  <% else %>
                  <option value=""></option>
                  <% end if %>
                </select></td>
            </tr>
			
			
			 <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do seu im&oacute;vel</font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                  <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                  
                </select></td>
            </tr>
			
			
			
			
            <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do seu im&oacute;vel </font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><font color="#FFFFFF">
                <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
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
                </select>
                </font></td>
            </tr>
            <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                      de dormit&oacute;rios do seu im&oacute;vel </font></div></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><select name="txt_quartos_vend" id="txt_quartos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                  <option value="não informado" selected>não informado</option>
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem do im&oacute;vel </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas_vend" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                <option value="não informado" selected>não informado</option>
                    <option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
              
			
              
			
            <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do seu im&oacute;vel </font></div></td>
              
            <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="20">
            </td>
            </tr>
			
			<tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>"><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF; background:<%=medio%>"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descrição 
                            do seu imóvel</font></div></td>
                    </tr>
                    <tr>
                        <td width="290" height="82" bgcolor="<%=claro%>">&nbsp;</td>
                    </tr>
                  </table></td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><textarea name="txt_descricao_vend" class="inputBox" id="txt_descricao_vend" style="HEIGHT: 100px; WIDTH: 290px; background: <%=medio%>" onKeyPress="return limitfield(this, 200)"></textarea></td>
            </tr>
			
			<tr> 
              
            <td width="290" bgcolor="<%=escuro%>" style="border:1px solid #FFFFFF; background:<%=escuro%>"> 
            </td>
              <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><table width="290" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td><input name="image" type="image"  src="bt_enviar001.jpg" width="145" height="18" border="0"></td>
                    <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
                  </tr>
                </table></td>
          </tr>
        </table></td>
            </tr>
			
          </table></form>
     </td>
    <td width="5">&nbsp;</td>
  </tr>
</table>


</td>
  </tr>
</table>


</body>
</html>
