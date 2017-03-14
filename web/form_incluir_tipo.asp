<% response.buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style_imoveis.asp"-->
<!--#include file="loggedin02.asp"-->
<!--#include file="cores.asp"-->

<%
dim varSucesso_tipo,varExistente
varSucesso_tipo = request.querystring("varSucesso_tipo")
varExistente = request.querystring("varExistente")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Incluir tipo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body onload=doublecombo.txt_tipo.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo"   method="post" action="incluir_tipo.asp">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center">
          <% if varSucesso_tipo = "" then%>
          
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varSucesso_tipo%> 
          foi incluído com sucesso.</font>
          <%end if%>
		   <% if varExistente = "" then%>
          
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varExistente%> </font>
         
          <%end if%>
		  
		  
		  
		  
          </font></div></td>
    </tr>
    <tr>
      <td><table width="345" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td><table width="335" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                  <td width="235" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><input name="txt_tipo" type="text" id="txt_tipo" size="38" maxlength="23" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=claro%>"></td>
                </tr>
                <tr>
                  <td width="100">&nbsp;</td>
                  <td width="235"><table width="235" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="117"><%if session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %><input name="image" type="image"  src="bt_enviar003.jpg" width="117" height="18" border="0"><%else%><img src="bt_enviar003.jpg" border="0"></img><%end if%></td>
                        <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar003.jpg" width="118" height="18" border="0"></a></td>
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
<!--#include file="dsn2.asp"-->
<% response.flush%>
  <%response.clear%>
