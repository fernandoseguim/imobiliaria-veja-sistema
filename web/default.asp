<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="style2_primeira.asp"-->

<%
dim Conexao3

Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

dim Sql33
dim rs33


Sql33 = "Select franquia.id_franquia,franquia.nome_franquia,franquia.data_franquia from franquia  ORDER BY id_franquia DESC"



Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs33.ActiveConnection = Conexao3
	
	
	rs33.Open Sql33, Conexao3


%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
    <td width="135" height="137"><img src="default_img02.jpg" width="135" height="137"></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="135" height="137"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><strong>Escolha 
        a cidade onde voc&ecirc; quer encontrar o seu im&oacute;vel.</strong></font></div></td>
    <td>&nbsp;</td>
  </tr>
  <form action="default_franquia.asp" target="_self" method="post">
  <tr>
    <td>&nbsp;</td>
    <td width="135" height="20"><select name="vOrigem_Franquia" class="inputBox" id="vOrigem_Franquia" style="HEIGHT: 18px; WIDTH: 135px; color: #000000; background: #FFFFFF;" >
        <% if not rs33.eof then %>
        <% While NOT Rs33.EoF %>
        <option value="<% = Rs33("nome_franquia") %>" <% if lcase(Rs33("nome_franquia")) = lcase("São Bernardo") then response.write "selected" end if %>> 
        <% = Rs33("nome_franquia") %>
        </option>
        <% Rs33.MoveNext %>
        <% Wend %>
        <%else%>
        <option value=""></option>
        <%end if%>
      </select></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="135" height="20"><input name="image" type="image" src="bt_entrar_franquia01.jpg" width="135" height="18"></td>
    <td>&nbsp;</td>
  </tr>
  </form>
</table>
</body>
</html>
<!--#include file="dsn2.asp"-->