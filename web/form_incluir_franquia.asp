<% response.buffer = true %>
<!--#include file="style_imoveis.asp"-->
<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="loggedin03.asp"-->













<%




'Criando conexão com o banco de dados! 
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
	
	dim varSucesso_Franquia
	
	varSucesso_Franquia = request.QueryString("varSucesso_Franquia")
	
%> 


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script>
function isValidDigitNumber (doublecombo)

</script>

<title>Incluir Franquia</title>


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body onload=doublecombo.txt_ip.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="incluir_franquia.asp">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center"> 
          <% if varSucesso_Franquia = "" then%>
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varSucesso_franquia%> 
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Franquia</font></div></td>
                  <td width="235" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><select name="txt_franquia" class="inputBox" id="txt_franquia" style="HEIGHT: 18px; WIDTH: 235px; background:<%=medio%>" >
                         <% if not rs33.eof then %>
				    <% While NOT Rs33.EoF %>
                    <option value="<% = Rs33("nome_combo1") %>" <% if lcase(Rs33("nome_combo1")) = lcase("São Bernardo") then response.write "selected" end if %>>
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
                  <td width="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></div></td>
                  <td width="235" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" id="txt_endereco" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=claro%>"></td>
                </tr>
				 <tr>
                  <td width="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                  <td width="235" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" id="txt_telefone" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>"></td>
                </tr>
				 <tr>
                  <td width="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                  <td width="235" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" id="txt_email" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=claro%>"></td>
                </tr>
               
				
				
				
                <tr>
                  <td width="100">&nbsp;</td>
                  <td width="235"><table width="235" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="117"><input name="image" type="image"  src="bt_enviar003.jpg" width="117" height="18" border="0"></td>
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
<% response.flush%>
  <%response.clear%>
