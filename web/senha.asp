<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->


<%
dim Conexao3

Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

dim Sql33
dim rs33


Sql33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs33.ActiveConnection = Conexao3
	
	
	rs33.Open Sql33, Conexao3


%>




<html>
<head>
<title>Senha</title>

<!--#include file="style2_primeira.asp"-->

</head>
<body onload=document.forms.senha.Admin_ID.focus(); bgcolor="<%=escuro%>" vlink="#48576C" link="#48576C" alink="#000000">

<br><br><br><br><br>

<font face="Verdana" size="1" color="#FFFFFF">
<%
	If Request("SecondTry") = "True" Then
		If Request("WrongPW") = "True" Then
			Response.Write "<center>Senha Inválida. Tente novamente.</center>"
		Else
			Response.Write "<center>Usuário Inválido.  Tente novamente.</center>"
		End If
	End If
%>
</font>


<form action="login.asp" Method="Post" name="senha">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="<%=claro%>" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
               
			   <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Franquia</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <select name="vOrigem_Franquia" class="inputBox" id="vOrigem_Franquia" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" >
                         <% if not rs33.eof then %>
				    <% While NOT Rs33.EoF %>
                    <option value="<% = Rs33("nome_combo1") %>" <% if lcase(Rs33("nome_combo1")) = lcase("São Bernardo") or lcase(Rs33("nome_combo1")) = lcase("Sao Bernardo") then response.write "selected" end if %>>
                    <% = Rs33("nome_combo1") %>
                    </option>
                    <% Rs33.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
					</select></td>

</tr>
			   
			   
			   
			   
			   
			    <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">ID:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="Admin_ID" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Senha:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Password" name="Password" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>">&nbsp;</td>
                  <td align="right" bgcolor="<%=claro%>"> 
                    <input  type="submit" class="inputSubmit" value="Entrar" size="20" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
</table>


</tr>
</table>
</td>
</tr>
</table>
</form>
<center>
</center>


</body>
</html>
<!--#include file="dsn2.asp"-->