<% response.buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style_imoveis.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->
<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT  combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open Sql3, Conexao3
	


dim varSucesso_bairro,varExistente,varBairro

varSucesso_bairro = request.querystring("varSucesso_bairro")
varExistente = request.querystring("varExistente")
varCidade = request.querystring("varCidade")
%> 


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body onload=doublecombo.txt_bairro.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo"   method="post" action="incluir_bairro.asp">
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
                  <td width="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td width="235" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><select name="txt_cidade" class="inputBox" id="txt_cidade" style="HEIGHT: 18px; WIDTH: 235px; background:<%=claro%>">
                      <% While NOT Rs3.EoF %>
                      <option value="<% = Rs3("id_combo1") %>"<% if rs3("nome_combo1") = "Santo André" and varCidade = "" then %>selected<%end if%><% if varCidade <> "" and varCidade = rs3("nome_combo1") then %>selected<%end if%>> 
                      <% = Rs3("nome_combo1") %>
                      </option>
                      <% Rs3.MoveNext %>
                      <% Wend %>
                    </select></td>
                </tr>
                <tr>
                  <td width="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" > 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                  <td width="235" bgcolor="#537497" style="border:1px solid #FFFFFF;"><input name="txt_bairro" type="text" id="txt_bairro" size="38" maxlength="23" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>"></td>
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
<!--#include file="dsn2.asp"-->
<% response.flush%>
  <%response.clear%>
