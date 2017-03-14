
<%

option explicit 
response.buffer=true



%>
<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->
<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_email_enviado,varSucesso_email
varCod_email_enviado = request.QueryString("varCod_email_enviado")
varSucesso_email = request.QueryString("varSucesso_email")
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT email_enviado.cod_email_enviado,email_enviado.nome,email_enviado.telefone,email_enviado.email,email_enviado.atendimento,email_enviado.de,email_enviado.para,email_enviado.assunto,email_enviado.mensagem,email_enviado.data  FROM email_enviado where cod_email_enviado="&varCod_email_enviado 
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
	'----------------------verifica comprador--------------------------	
	
	
%>		








<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Email</title>
<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (nform) 



{




{
if (nform.txtNome.value == "") {
		alert("Digite seu nome.");
		nform.txtNome.focus();
		nform.txtNome.select();
		return false;
}
}

{
if (nform.txtEmail.value == "") {
		alert("Digite seu email.");
		nform.txtEmail.focus();
		nform.txtEmail.select();
		return false;
}
}






{
if (nform.txtAssunto.value == "") {
		alert("Digite o assunto.");
		nform.txtAssunto.focus();
		nform.txtAssunto.select();
		return false;
}
}

{
if (nform.txtMensagem.value == "") {
		alert("Digite sua Mensagem.");
		nform.txtMensagem.focus();
		nform.txtMensagem.select();
		return false;
}
}






	
}












</script>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=605,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>


</head>
<!--#include file="style2_sugestoes.asp"-->
<body bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post"  onSubmit="return isValidDigitNumber(this);" name="b2">
<table width="590" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=escuro%>">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  <tr>
      <td width="590" height="18" bgcolor="<%=escuro%>"> 
        <div align="center"></div></td>
  </tr>
  <tr>
    <td width="590" height="54"><table width="590" height="54" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="5" height="54" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="580" height="54"><table width="580" height="54" border="0" cellpadding="0" cellspacing="0">
             
			 
			   
			 
			 
			 
			   
			   
			    <tr> 
                  <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atendimento</font> 
                    </div></td>
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtNome4" type="text" id="txtNome4" value="<%=rs("atendimento")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
			  
			   <tr> 
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">De</font></div></td>
                  <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txtNome3" type="text" id="txtNome3" value="<%=rs("de")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
			    
				<tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para</font></div></td>
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtNome2" type="text" id="txtNome2" value="<%=rs("para")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
				
				
				
				
				<tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome</font></div></td>
                <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txtNome" type="text" id="txtNome" value="<%=rs("nome")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
              <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtEmail" type="text" id="txtTelefone" value="<%if rs("telefone") <> "" then response.write rs("telefone") else response.write "não informado" end if%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
			  
			   <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txtEmail" type="text" id="txtTelefone" value="<%=rs("email")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
			  
			  
              <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto</font></div></td>
                <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtAssunto" type="text" id="txtEmail" value="<%=rs("assunto")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
            </table></td>
            <td width="5" height="54" bgcolor="<%=escuro%>">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="5" height="140" bgcolor="<%=escuro%>">&nbsp;</td>
          <td><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td><table width="288" height="142" border="0" cellpadding="0" cellspacing="0" align="right">
                      <tr> 
                        <td width="290" align="center" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Mensagem</font></td>
                    </tr>
                    <tr> 
                        <td width="290" height="122" bgcolor="<%=escuro%>"> 
                          <div align="center"></div></td>
                    </tr>
                  </table></td>
                <td><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="290" height="122" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txtSugestao" cols="32" rows="8" class="inputBox" id="txtMensagem" style="HEIGHT: 120px; WIDTH: 290px; background: <%=medio%>"  OnKeyPress="return limitfield(this, 500)"><%=rs("mensagem")%></textarea></td>
                    </tr>
                    <tr>
                      <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">

                          <tr>
                              <td width="145" height="18" bgcolor="<%=escuro%>">&nbsp;</td>
                              <td width="145" height="18" bgcolor="<%=escuro%>">&nbsp;</td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
            <td width="5" height="140" bgcolor="<%=escuro%>">&nbsp;</td>
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
		   Set Conexao = Nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
