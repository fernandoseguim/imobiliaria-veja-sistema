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




<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
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
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="incluir_imovel02.asp"> 
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  <tr>
    <td><div align="center">
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
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" id="txt_proprietario" size="38" maxlength="33" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      do propriet&aacute;rio do im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" id="txt_telefone" size="38" maxlength="12" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> <input name="txt_endereco" type="text" id="txt_endereco" size="38" maxlength="33" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=medio%>"></td>
                </tr>
				
				<tr bgcolor="<%=claro%>"> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                      do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_titulo" type="text" id="txt_titulo4" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
              <tr bgcolor="<%=medio%>"> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_anuncio" type="text" id="txt_anuncio" size="38" maxlength="90" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
				
				
				 <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presen&ccedil;a 
                      na Primeira P&aacute;gina</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_presenca_primeira" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="incluido">Incluído</option>
                      <option value="excluido" selected>Excluído</option>
                    </select>
                  </td>
                </tr>
				
				
				
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link 
                      de visualiza&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_link_foto" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="icon_foto.gif">Com Foto</option>
                      <option value="icon_foto2.gif" selected>Sem Foto</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                     <% if not rs3.eof then %>
				    <% While NOT Rs3.EoF %>
                    <option value="<% = Rs3("id_combo1") %>" <% if rs3("nome_combo1") = "Santo André" then%>selected<%else%><%end if%>>
                    <% = Rs3("nome_combo1") %>
                    </option>
                    <% Rs3.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
                  </select>
                    <a href="javascript:newWindow3('form_incluir_cidade.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a> 
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo2" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <% if not rs4.eof then%>
					<% While NOT Rs4.EoF %>
                    <option value="<% = Rs4("id_combo2") %>"<%if rs4("nome_combo2") = "Bairro Campestre" then%>selected<%end if%>>
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="casa" selected>Casa</option>
                      <option value="apartamento">Apartamento </option>
                      <option value="flat">Flat</option>
                      <option value="terreno">Terreno</option>
                      <option value="rural">Rural</option>
                      <option value="comercial">Comercial</option>
                    </select>
                    </font></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Total</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <input name="txt_a_total" type="text" id="txt_a_total" size="12" maxlength="12" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font></td>
                </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Constru&iacute;da </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_a_constr" type="text" id="txt_a_constr" size="12" maxlength="12" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font> </td>
                </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos" id="select4" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_banheiros" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na Garagem</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_negociacao" id="select7" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="aluguel" selected>Aluguel</option>
                      <option value="venda">Venda</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_valor" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="12">
                  </td>
                </tr>
                
				<tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="excluido" selected>Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="não informado" selected>não informado</option>
                    <option value="vago">vago</option>
                    <option value="ocupado">ocupado</option>
                    
                  </select>
                  </td>
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
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="obs_proprietario" class="inputBox" id="obs_proprietario" style="HEIGHT: 100px; WIDTH: 290px; background: <%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
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
