<% response.buffer = true %>
<!--#include file="style_imoveis.asp"-->
<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->


<%
dim varSucesso_telefone_acesso,varExistente

varSucesso_telefone_acesso = request.querystring("varSucesso_telefone_acesso")
varExistente = request.querystring("varExistente")


%>














<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script>
function isValidDigitNumber (doublecombo)
{

	
	
	
	if (doublecombo.txt_telefone_acesso.value == "") {
        alert("Você precisa digitar um telefone de acesso!");
        doublecombo.txt_telefone_acesso.focus();
		doublecombo.txt_telefone_acesso.select();
        return false;
    }
}	
</script>

<title>Incluir telefone de acesso</title>


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body onload=doublecombo.txt_telefone_acesso.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="incluir_telefone_acesso.asp">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center"> 
          <% if varSucesso_telefone_acesso = "" then%>
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varSucesso_telefone_acesso%> 
          foi incluído com sucesso.</font> 
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      de acesso</font></div></td>
                  <td width="235" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone_acesso" type="text" id="txt_telefone_acesso" size="38" maxlength="23" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>"></td>
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
