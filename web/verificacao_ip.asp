<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="style2_sugestoes.asp"-->
<%


dim varEnderecoIP,SQLIP,rsIP
	varEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	
	'if session("permissao") <> "6" then
	'varEnderecoIP = "123.345.678"
	'else
	'varEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	'end if
	
	Set rsIP = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQLIP = "Select * from ip where ip ='"&varEnderecoIP&"' ORDER BY id_ip DESC"

rsIP.Open SQLIP, Conexao
 
 if rsIP.eof then
 
 else
' Response.Redirect "archive_imoveis.asp"
 end if
 


if rsIP.eof then
%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>" >
<%

Set rsIP02 = Server.CreateObject("ADODB.RecordSet")
dim SQLIP02
dim rsIP02
SQLIP02 = "Select * from ip where senha_incluiu like '"&session("strPassword")&"' ORDER BY id_ip DESC"

rsIP02.Open SQLIP02, Conexao
 
 if rsIP02.eof then
 response.redirect "senha.asp"
 else
 %>
 <div align="center">
 <form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="atualizar_ip02.asp?varCodIP=<%=rsIP02("id_ip")%>">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48">&nbsp;</td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center"> 
         <%
		 dim varSucessoSenha
		 
		 varSucessoSenha = request.querystring("varSucessoSenha")
		 
		 
		 %>
		 
		 
		 
		  <% if varSucessoSenha = "" then%>
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varSucessoSenha%> 
          foi atualizado com sucesso.</font> 
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">IP</font></div></td>
                  <td width="235" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_ip" type="text" class="inputBox" id="txt_ip" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>" value="<%=rsIP02("ip")%>" size="38" maxlength="23" align="left"></td>
                </tr>
				
                
				
				
				
                <tr>
                  <td width="100">&nbsp;</td>
                  <td width="235"><table width="235" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="117"><input type="image" src="bt_atualizar003.jpg" width="117" height="18"></td>
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
</div>
 <%
 'response.write rsIP02("senha_incluiu")&"||"
 end if


%>



<%
else
response.Redirect "archive_imoveis.asp"
 end if
%>
</body>
</html>
<!--#include file="dsn2.asp"-->