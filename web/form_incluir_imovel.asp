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

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 



<%
Function EscreveFuncaoJavaScript222 ( Conexao333 )
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
Set rsMarcas333 = Conexao333.Execute ( SqlMarcas333 )

While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas333("id_combo2")&""

Set rsCarros333 = Conexao333.Execute ( SqlCarros333 )

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

'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'

Sql333 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs333 = Conexao333.Execute ( Sql333 ) 




%> 










<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
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
		
		
%>		



<script>
function isValidDigitNumber (doublecombo)
{
{





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



var strValidNumber1_77="1234567890";
for (nCount=0; nCount < doublecombo.txt_a_total.value.length; nCount++) 
		{
strTempChar1_77=doublecombo.txt_a_total.value.substring(nCount,nCount+1);
if (strValidNumber1_77.indexOf(strTempChar1_77,0)==-1) 
{
alert("O formulário Área Total só pode conter números!");
doublecombo.txt_a_total.focus();
doublecombo.txt_a_total.select();
return false;
}
}







var strValidNumber1_777="1234567890";
for (nCount=0; nCount < doublecombo.txt_a_constr.value.length; nCount++) 
		{
strTempChar1_777=doublecombo.txt_a_constr.value.substring(nCount,nCount+1);
if (strValidNumber1_777.indexOf(strTempChar1_777,0)==-1) 
{
alert("O formulário Área Construída só pode conter números!");
doublecombo.txt_a_constr.focus();
doublecombo.txt_a_constr.select();
return false;
}
}





if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do comprador!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
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
	
	
	
	
	if (doublecombo.txt_endereco.value == "") {
        alert("O formulário Endereço do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }
	
	
	if (doublecombo.blob.value == "") {
        alert("O formulário Foto Grande está vazio!");
        doublecombo.blob.focus();
		doublecombo.blob.select();
        return false;
    }
	
	 vfile = doublecombo.blob.value;
    tfile = vfile.length;
    
    if (vfile.substr(tfile - 4, 4) != ".jpg" && vfile.substr(tfile - 4, 4) != ".gif" && vfile.substr(tfile - 4, 4) != ".GIF" && vfile.substr(tfile - 4, 4) != ".JPG") {
        alert("O arquivo do formulário Foto Grande deverá possuir o formato (.jpg) ou (.gif)!");
        doublecombo.blob.value == vfile.substr(tfile - 4, 4);
		doublecombo.blob.focus();
		doublecombo.blob.select();
		
		
        return false;
    }
	
	
	



var strVerif2 = doublecombo.blob.value;
var	strVerif_n2 = strVerif2.length;
if (strVerif2.substring(strVerif_n2 - 15,strVerif_n2) == "imovel00000.jpg" ){

       alert("Você escolheu o arquivo errado, imovel00000.jpg não pode ser enviado.");
       doublecombo.blob.focus();
		doublecombo.blob.select();
		
		
		
return false;

}


	

//---------------------------------configuração do combo foto_pequena---------------------


	if (doublecombo.blob2.value == "") {
        alert("O formulário Foto Pequena está vazio!");
        doublecombo.blob2.focus();
		doublecombo.blob2.select();
        return false;
    }
	
	 vfile2 = doublecombo.blob2.value;
    tfile2 = vfile2.length;
    
    if (vfile2.substr(tfile2 - 4, 4) != ".jpg" && vfile2.substr(tfile2 - 4, 4) != ".gif" && vfile2.substr(tfile2 - 4, 4) != ".GIF" && vfile2.substr(tfile2 - 4, 4) != ".JPG") {
        alert("O arquivo do formulário Foto Pequena deverá possuir o formato (.jpg) ou (.gif)!");
        doublecombo.blob2.value == vfile2.substr(tfile2 - 4, 4);
		doublecombo.blob2.focus();
		doublecombo.blob2.select();
		
		
        return false;
    }
	
	
	




var strVerif3 = doublecombo.blob2.value;
var	strVerif_n3 = strVerif3.length;
if (strVerif3.substring(strVerif_n3 - 20,strVerif_n3) == "mini_imovel00000.jpg" ){

       alert("Você escolheu o arquivo errado, mini_imovel00000.jpg não pode ser enviado.");
       doublecombo.blob2.focus();
		doublecombo.blob2.select();
		
		
		
return false;

}




//--------------------------------------------------------------------










	
	
	
	
	
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



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_6="1234567890,.";
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









//------------------------------saldo devedor--------------------------







if (doublecombo.txt_ja_pago_devedor.value == "") {
        alert("O formulário valor já pago no saldo devedor está vazio!");
        doublecombo.txt_ja_pago_devedor.focus();
		doublecombo.txt_ja_pago_devedor.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_ja="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_ja_pago_devedor.value.length; nCount++) 
		{
strTempChar1_ja=doublecombo.txt_ja_pago_devedor.value.substring(nCount,nCount+1);
if (strValidNumber1_ja.indexOf(strTempChar1_ja,0)==-1) 
{
alert("O formulário valor já pago no saldo devedor só pode conter números!");
doublecombo.txt_ja_pago_devedor.focus();
doublecombo.txt_ja_pago_devedor.select();
return false;
}
}






if (doublecombo.txt_devendo_devedor.value == "") {
        alert("O formulário valor devido no saldo devedor está vazio!");
        doublecombo.txt_devendo_devedor.focus();
		doublecombo.txt_devendo_devedor.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_devendo="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_devendo_devedor.value.length; nCount++) 
		{
strTempChar1_devendo=doublecombo.txt_devendo_devedor.value.substring(nCount,nCount+1);
if (strValidNumber1_devendo.indexOf(strTempChar1_devendo,0)==-1) 
{
alert("O formulário valor devido no saldo devedor só pode conter números!");
doublecombo.txt_devendo_devedor.focus();
doublecombo.txt_devendo_devedor.select();
return false;
}
}














//----------------------------------------------------------------------

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
//-----------------------------------------------

}
}









</script>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>



<html>


<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao333 ) %>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_captacao.focus(); bgcolor="blue" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo" ENCTYPE="multipart/form-data" onSubmit="return isValidDigitNumber(this);" method="post" action="outFile001.asp">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="590" height="48"><a href="form_incluir_imovel.asp"><img src="top_resultado.jpg" width="590" height="48" border="0"></a></td>
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
          foi incluido com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><a href="form_incluir_imovel02.asp"><img src="bt_foto0.jpg" width="290" height="18" border="0"></a></td>
                  <td><a href="form_incluir_imovel.asp"><img src="bt_foto1.jpg" width="290" height="18" border="0"></a></td>
              </tr>
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Capta&ccedil;&atilde;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_captacao" type="text" id="txt_captacao" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
			  
			  
			  
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Uacute;ltima 
                      foto inclu&iacute;da</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input  type="text" name="ultimo" value="<%if not rs.eof then%> <%=rs.movefirst%>
				
				
				  <% if rs("Foto_Grande5") <> "imovel00000.jpg" then %>
					 <%=rs("Foto_Grande5")%>
					 <%end if%>
					 
					 <% if rs("Foto_Grande5") = "imovel00000.jpg" and rs("foto_grande4") <> "imovel00000.jpg" then %>
					 <%=rs("Foto_Grande4")%>
					 <%end if%>
					 
					 <% if rs("Foto_Grande4") = "imovel00000.jpg" and  rs("foto_grande3") <> "imovel00000.jpg"  then %>
					 <%=rs("Foto_Grande3")%>
					 <%end if%>
					 
					 <% if rs("Foto_Grande3") = "imovel00000.jpg" and  rs("foto_grande2") <> "imovel00000.jpg" then %>
					 <%=rs("Foto_Grande2")%>
					 <%end if%>
					 
					 <% if rs("Foto_Grande2") = "imovel00000.jpg" and  rs("foto_grande1") <> "imovel00000.jpg" then %>
					 <%=rs("Foto_Grande1")%>
					 <%end if%>
					 
					 <% if rs("Foto_Grande1") = "imovel00000.jpg" and  rs("foto_grande") <> "imovel00000.jpg" then %>
					 <%=rs("Foto_Grande")%>
					 <%end if%>
					 
				 
				 
				 
				 
				 <%else%>Sem registro<%end if%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Propriet&aacute;rio 
                    do im&oacute;vel</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" id="txt_proprietario" size="38" maxlength="35" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                    do propriet&aacute;rio do im&oacute;vel</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" id="txt_telefone" size="38" maxlength="30" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                    do propriet&aacute;rio do im&oacute;vel</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" id="txt_email" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                    do Im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_endereco" type="text" id="txt_endereco" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background: <%=medio%>;"></td>
              </tr>
			  
			  <tr bgcolor="<%=claro%>"> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                      do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_titulo" type="text" id="txt_titulo4" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
              <tr bgcolor="<%=medio%>"> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_anuncio" type="text" id="txt_anuncio" size="38" maxlength="120" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"></td>
              </tr>
			  
			  
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                    Grande</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="blob" type="file"  class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" size="40" maxlength="80" align="left" ></td>
              </tr>
			  <tr bgcolor="<%=medio%>"> 
                  <td width="290" bgcolor="<%=medio%>"  style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                      Pequena</font></div></td>
                  <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="blob2" type="file" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>" size="40" maxlength="80" align="left" >
                  </td>
              </tr>
			  
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presen&ccedil;a 
                      na Primeira P&aacute;gina</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_presenca_primeira" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="excluido"selected>Excluído</option>
                    <option value="incluido">Incluído</option>
                  </select>
                  </td>
              </tr>
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link 
                      de visualiza&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_link_foto" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="icon_foto.gif"selected>Com Foto</option>
                    <option value="icon_foto2.gif">Sem Foto</option>
                  </select>
                  </td>
              </tr>
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
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
                    <a href="javascript:newWindow3('form_incluir_cidade.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></img></a></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro/Regi&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo2" class="inputBox" onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
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
                    <a href="javascript:newWindow3('form_incluir_bairro.asp')"><img src="bt_mais02.jpg" width="18" height="18" border="0"></a> 
                  </td>
              </tr>
			  
			   <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                     <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                    </select>
                    <a href="javascript:newWindow3('form_incluir_vila.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a> 
                  </td>
              </tr>
			  
			  
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
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
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Total</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <input name="txt_a_total" type="text" class="inputBox" id="txt_a_total" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Constru&iacute;da </font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_a_constr" type="text" class="inputBox" id="txt_a_constr" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font> </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos" id="select4" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
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
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_banheiros" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
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
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na Garagem</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
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
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_negociacao" id="select7" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                    <option value="aluguel" selected>Aluguel</option>
                    <option value="venda">Venda</option>
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_valor" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
                      devedor </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_saldo_devedor" id="txt_saldo_devedor" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="sem saldo devedor" selected>Sem saldo devedor</option>
					  <option value="com saldo devedor" >Com saldo devedor</option>
                    
                    
                  </select>
                  </td>
              </tr>
			  
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      j&aacute; pago no saldo devedor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_ja_pago_devedor" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      devido no saldo devedor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_devendo_devedor" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  
			  
			    <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="excluido" selected>Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="não informado" selected>não informado</option>
                    <option value="vago">vago</option>
                    <option value="ocupado">ocupado</option>
                    
                  </select>
                  </td>
              </tr>
			  
			  <tr>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Qualidade 
                      do neg&oacute;cio</font></div></td>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_qualidade" id="txt_qualidade" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="bom negócio" >Bom Negócio</option>
                    <option value="negócio comum" selected>Negócio Comum</option>
                    
                  </select></td>
              </tr>
			  
			  
              <tr>
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
                      sobre o im&oacute;vel</font></div></td>
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="obs_imovel" type="text" id="obs_imovel" size="38" maxlength="200" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
			   
			  
			  
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
                            sobre o propriet&aacute;rio</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=claro%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="obs_proprietario" class="inputBox" id="obs_proprietario" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><% if  session("permissao") = "3" or session("permissao") = "2" or session("permissao") = "4" or session("permissao") = "5" then %><input name="image" type="image"  src="bt_enviar001.jpg" width="145" height="18" border="0"><%else%><img  src="bt_enviar001.jpg" width="145" height="18" border="0"></img><%end if%></td>
                        <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
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
