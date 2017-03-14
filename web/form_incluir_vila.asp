<!--#include file="dsn.asp"-->
<% response.buffer = true %>
<!--#include file="style_imoveis.asp"-->

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
Response.Write "switch (form.txt_cidade.options[form.txt_cidade.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 


Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas3.ActiveConnection = Conexao3
	
	
	rsMarcas3.Open SqlMarcas3, Conexao3
	

While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.txt_bairro.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"



Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros3.ActiveConnection = Conexao3
	
	
	rsCarros3.Open SqlCarros3, Conexao3
	
	


'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.txt_bairro.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros3.EoF

Response.Write "form.txt_bairro.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "form.txt_bairro.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 


rsMarcas3.close

set rsMarcas3 = nothing

rsCarros3.close

set rsCarros3 = nothing



End Function
%> 


<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 




Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open Sql3, Conexao3
	
	


%> 














<%

dim varSucesso_bairro,varExistente,varBairro

varSucesso_bairro = request.querystring("varSucesso_bairro")
varExistente = request.querystring("varExistente")
varCidade = request.querystring("varCidade")
%> 


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script>
function isValidDigitNumber (doublecombo)
{
if (doublecombo.txt_bairro.value == "bqualquer") {
        alert("Você precisa escolher um bairro!");
        doublecombo.txt_bairro.focus();
		
        return false;
    }
	
	if (doublecombo.txt_vila.value == "") {
        alert("Você precisa digitar a vila!");
        doublecombo.txt_vila.focus();
		doublecombo.txt_vila.select();
        return false;
    }
}	
</script>
<% EscreveFuncaoJavaScript ( Conexao3) %>
<title>Incluir Vila</title>


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body  bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="incluir_vila.asp">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center"> 
          <% if varSucesso_bairro = "" then%>
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varSucesso_bairro%> 
          foi incluída com sucesso.</font> 
          <%end if%>
          <% if varExistente = "" then%>
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varExistente%> 
          </font> 
          <%end if%>
        </div></td>
    </tr>
    <tr>
      <td><table width="345" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td><table width="335" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td width="235" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><select name="txt_cidade" onChange="javascript:atualizacarros(this.form);" class="inputBox" id="txt_cidade" style="HEIGHT: 18px; WIDTH: 235px; background:<%=medio%>">
                      <% While NOT Rs3.EoF %>
                      <option value="<% = Rs3("id_combo1") %>"<% if rs3("nome_combo1") = "Santo André" and varCidade = "" then %>selected<%end if%><% if varCidade <> "" and varCidade = rs3("nome_combo1") then %>selected<%end if%>> 
                      <% = Rs3("nome_combo1") %>
                      </option>
                      <% Rs3.MoveNext %>
                      <% Wend %>
                    </select></td>
                </tr>
				<tr>
                  <td width="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                  <td width="235" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><select name="txt_bairro" class="inputBox" id="txt_bairro" style="HEIGHT: 18px; WIDTH: 235px; background:<%=claro%>">
                     <option value="bqualquer">Bairro/Região</option>
					 <option value="bqualquer">qualquer um</option>
                    </select></td>
                </tr>
				
                <tr>
                  <td width="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" > 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                  <td width="235" bgcolor="#537497" style="border:1px solid #FFFFFF;"><input name="txt_vila" type="text" id="txt_vila" size="38" maxlength="23" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>"></td>
                </tr>
                <tr>
                  <td width="100">&nbsp;</td>
                  <td width="235"><table width="235" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="117"><%if session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %><input name="image" type="image"  src="bt_enviar003.jpg" width="117" height="18" border="0"><%else%><img src="bt_enviar003.jpg" border="0"></img><%end if%></td>
                        <td width="117"><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar003.jpg" width="118" height="18" border="0"></a></td>
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
</body>
</html>

<%

rs3.close

set rs3 = nothing

conexao3.close

set conexao3 = nothing




%>
<% response.flush%>
  <%response.clear%>
