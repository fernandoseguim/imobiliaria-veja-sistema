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



'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")






'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Sql5 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs5 = Conexao3.Execute ( Sql5 )
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 










<%
Function EscreveFuncaoJavaScript2 ( Conexao4 )
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
SqlMarcas4 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas4 = Conexao4.Execute ( SqlMarcas4 )

While NOT rsMarcas4.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas4("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros4 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas4("id_combo1")&" order by nome_combo2"
Set rsCarros4 = Conexao4.Execute ( SqlCarros4 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1   
While NOT rsCarros4.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros4("nome_combo2") & "','" & rsCarros4("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros4.MoveNext
Wend
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas4.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao4 = Server.CreateObject("ADODB.Connection")
Conexao4.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql4 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs5 = Conexao4.Execute ( Sql4 ) 
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
   dim rs4,strSQL4,strSQL6,rs6
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	Set rs6 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	strSQL6 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		rs4.Open strSQL4, Conexao
		rs6.Open strSQL6, Conexao
		
		
%>		



<script>
function isValidDigitNumber (doublecombo)
{

 vfile = doublecombo.blob.value;
    tfile = vfile.length;
    
    if (vfile.substr(tfile - 4, 4) != ".jpg" && vfile.substr(tfile - 4, 4) != ".gif") {
        alert("O arquivo do formulário Foto deverá possuir o formato (.jpg) ou (.gif)!");
        doublecombo.blob.value == vfile.substr(tfile - 4, 4);
		doublecombo.blob.focus();
		doublecombo.blob.select();
		
		
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
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("O formulário do telefone está vazio!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
    }
	
	
	
	if (doublecombo.txt_descricao_comp.value == "") {
        alert("O formulário descrição do  imóvel pretendido está vazio!");
        doublecombo.txt_descricao_comp.focus();
		doublecombo.txt_descricao_comp.select();
        return false;
    }
	
	
	if (doublecombo.txt_descricao_vend.value == "") {
        alert("O formulário descrição do imóvel seu  está vazio!");
        doublecombo.txt_descricao_vend.focus();
		doublecombo.txt_descricao_vend.select();
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
	
	
	
	if (doublecombo.txt_valor_comp.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.txt_valor_comp.focus();
		doublecombo.txt_valor_comp.select();
        return false;
    }
	
	if (doublecombo.txt_valor_vend.value == "") {
        alert("O formulário valor do seu Imóvel está vazio!");
        doublecombo.txt_valor_vend.focus();
		doublecombo.txt_valor_vend.select();
        return false;
    }
	


var strText2_4 = doublecombo.txt_valor_vend.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.txt_valor_vend.focus();
		
		doublecombo.txt_valor_vend.select();
		
return false;

}






var strText2_5 = doublecombo.txt_valor_comp.value;
var s_strText2_5 = strText2_5.length
if (strText2_5.substring((s_strText2_5 - 3), (s_strText2_5 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.txt_valor_comp.focus();
		
		doublecombo.txt_valor_comp.select();
		
return false;

}





	var strValidNumber1_6="1234567890,";
for (nCount=0; nCount < doublecombo.txt_valor_vend.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.txt_valor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor_vend.focus();
doublecombo.txt_valor_vend.select();
return false;
}
}



	var strValidNumber1_7="1234567890,";
for (nCount=0; nCount < doublecombo.txt_valor_comp.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_valor_comp.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor_comp.focus();
doublecombo.txt_valor_comp.select();
return false;
}
}











	
}
}
//-----------------------------------------------










</script>



<html>


<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript2 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript888 ( Conexao333 ) %>
<%  EscreveFuncaoJavaScript999 ( Conexao333 ) %>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>
<title>Incluir Permutante</title>
</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo" ENCTYPE="multipart/form-data" onSubmit="return isValidDigitNumber(this);" method="post" action="outFile004.asp">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  
   <tr>
      <td width="590" height="90">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Este 
          formul&aacute;rio &eacute; igual ao anterior, a &uacute;nica diferente 
          &eacute; que voc&ecirc; pode mandar a foto do seu im&oacute;vel junto 
          com as outras informa&ccedil;&otilde;es. Para enviar a foto clique no 
          bot&atilde;o &quot;procurar(ou browse)&quot; da caixa de texto abaixo, 
          e o sistema ir&aacute; buscar uma foto que esteja nos arquivos do seu 
          computador, selecione a foto e clique duas vezes para que a foto seja 
          inclu&iacute;da no formul&aacute;rio, depois basta enviar junto com 
          os outros dados das outras caixas de texto.</strong></font></div></td>
  </tr>
  
  
  <tr>
    <td height="18"><div align="center">
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
                  <td bgcolor="<%=escuro%>" >
<div align="center"><a href="form_permuta_incluir01.asp"><img src="bt_foto0.jpg" width="290" height="18" border="0"></a></div></td>
                  <td bgcolor="<%=escuro%>" ><img src="bt_foto1.jpg" width="290" height="18" border="0"></td>
              </tr>  
				
				
				
				<tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      nome </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" id="txt_proprietario" size="38" maxlength="35" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      telefone </font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" id="txt_telefone" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      email </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" id="txt_email" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                      do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_endereco" type="text" id="txt_endereco" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background: <%=medio%>"></td>
              </tr>
			  
			 
              
			  
			  
             
			  
			  
			  
                
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      Cidade do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo1" class="inputBox" id="combo1" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" onChange="javascript:atualizacarros(this.form);">
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
                    </font></font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      Bairro do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo2" class="inputBox" onChange="javascript:atualizacarros888(this.form);" id="combo2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
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
                    </select> </td>
              </tr>
			  
			   <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      Vila do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                  
                    </select> </td>
              </tr>
			  
			  
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
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
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                      de dormit&oacute;rio do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_quartos_vend" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                    <option value="não informado" selected>Não informado</option>
					<option value="01" >01</option>                    
					<option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07 </option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                    
				  
				  
				  </select>
                    </font></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_vagas_vend" size="1" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                    <option value="não informado" selected>Não informado</option>
					<option value="01" >01</option>                    
					<option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07 </option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                    
				  
				  
				  </select>
                    </font></td>
              </tr>
			  
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
                    <input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="20">
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                      do seu im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    
                    <input name="txt_descricao_vend" type="text" id="txt_descricao_vend" size="38" maxlength="200" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                    </td>
              </tr>
			  
			  
			   <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Clique 
                      em &quot;procurar&quot; para buscar a foto do im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="blob" type="file"  class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" size="40" maxlength="80" align="left" ></td>
              </tr>
			  <tr>
                  <td height="30" bgcolor="<%=medio%>" > 
                    <div align="center"></div></td>
                  <td bgcolor="<%=medio%>" > 
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></font> </td>
              </tr>
			  
			  
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" onChange="javascript:atualizacarros2(this.form);">
                     <option value="cqualquer" selected>Cidade</option>
					  <% if not rs5.eof then %>
                      <% While NOT Rs5.EoF %>
                      <option value="<% = Rs5("id_combo1") %>">
                    <% = Rs5("nome_combo1") %>
                    </option>
                    <% Rs5.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
                  </select>
                    </font></font> </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      do im&oacute;vel pretendido </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<select name="combo4" onChange="javascript:atualizacarros999(this.form);" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="bqualquer" selected>Bairro/Região</option>
					  <% if not rs6.eof then%>
                      <% While NOT Rs6.EoF %>
                      <option value="<% = Rs6("id_combo2") %>">
                    <% = Rs6("nome_combo2") %>
                    </option>
                    <% Rs6.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
					
					
                  </select></td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
					
                  </select></td>
              </tr>
			  
			  
			  
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      de im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo_comp" size="1" id="txt_tipo_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
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
                    </font> </td>
              </tr>
                
				
				
				<tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                      de dormit&oacute;rios do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_quartos_comp" size="1" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="não informado" selected>Não informado</option>
                   <option value="01" >01</option>                    
					<option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07 </option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                    </select>
                    </font> </td>
              </tr>
			  
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_vagas_comp" size="1" id="txt_vagas_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="não informado" selected>Não informado</option>
                    <option value="01" >01</option>                    
					<option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07 </option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                    </select>
                    </font> </td>
              </tr>
                
                
				
				<tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <input name="txt_valor_comp" type="text" class="inputBox" id="txt_valor_comp" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="20">
                    </font> </td>
              </tr>
               
			    <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Voc&ecirc; 
                      aceita somente vender o seu imóvel?</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_pergunta" size="1" id="txt_pergunta" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="sim" >Sim</option>
                      <option value="nao" selected>Não </option>
                     
                    </select>
                    </font> </td>
              </tr>
				
				
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            do im&oacute;vel pretendido</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=claro%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao_comp" class="inputBox" id="txt_descricao_comp" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 200)"></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><input name="image" type="image"  src="bt_enviar001.jpg" width="145" height="18" border="0"></td>
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
