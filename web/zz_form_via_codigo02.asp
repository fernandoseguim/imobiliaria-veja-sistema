<!--#include file="dsn.asp"-->
<!--#include file="style_imoveis.asp"-->
<!--#include file="cores.asp"-->



<%

dim varSucesso

varSucesso = request.QueryString("varSucesso")



dim Conexao




Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn

dim rs444Captacao,strSQL444Captacao
   
	strSQL444Captacao = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
			
			 
Set rs444Captacao = Server.CreateObject("ADODB.RecordSet")

	rs444Captacao.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Captacao.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Captacao.ActiveConnection = Conexao
	
	
	rs444Captacao.Open strSQL444Captacao, Conexao
	
	
	
	'-----------------------------------------------------------------
	
	
	
	dim rs555Captacao,strSQL555Captacao
   
	strSQL555Captacao = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
			
			 
Set rs555Captacao = Server.CreateObject("ADODB.RecordSet")

	rs555Captacao.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs555Captacao.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs555Captacao.ActiveConnection = Conexao
	
	
	rs555Captacao.Open strSQL555Captacao, Conexao
	
	
	
	
	 	
			%>





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Mudança de captador</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<p align="center"> <strong><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Mudan&ccedil;a 
  de captac&atilde;o:<br>
  <Br>
  </font></strong><br>
  <br>
<form action="atualizar_via_codigo02.asp" method="post">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30" bgcolor="<%=claro%>"><table width="580" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr bgcolor="<%=claro%>"> 
    <td width="200" bgcolor="<%=claro%>"><select name="txt_atendimento01" id="txt_atendimento01" class="inputBox" style="HEIGHT: 18px; WIDTH: 200px; background: <%=medio%>">
    <option value="Internet" >Internet</option>
    <% if not rs444Captacao.eof then %>
    <% While NOT rs444Captacao.EoF %>
    <option value="<% = rs444Captacao("list_name") %>"> 
    <% = rs444Captacao("list_name") %>
    </option>
    <% rs444Captacao.MoveNext %>
    <% Wend %>
    <%else%>
    <option value="Internet">Internet</option>
    <%end if%>
  </select></td>
    <td height="20" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>mudar 
        para</strong></font></div></td>
    <td width="200"><select name="txt_atendimento02" id="txt_atendimento02" class="inputBox" style="HEIGHT: 18px; WIDTH: 200px; background: <%=medio%>">
    <option value="Internet" >Internet</option>
    <% if not rs555Captacao.eof then %>
    <% While NOT rs555Captacao.EoF %>
    <option value="<% = rs555Captacao("list_name") %>"> 
    <% = rs555Captacao("list_name") %>
    </option>
    <% rs555Captacao.MoveNext %>
    <% Wend %>
    <%else%>
    <option value="Internet">Internet</option>
    <%end if%>
  </select></td>
	<td width="80"><input name="image" type="image"  src="bt_mudar01.jpg" width="80" height="20" border="0"></td>
  </tr>
</table></td>
  </tr>
</table>
</form>
<br><br><br><br><br><br>
<center><strong><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso%></font></strong></center>

<p>&nbsp; 
</body>
</html>
