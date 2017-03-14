
<%

option explicit 
response.buffer=true



%>
<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->
<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_telefone_acesso,varSucesso_email

varCod_telefone_acesso = request.QueryString("varCod_telefone_acesso")

dim varSucessoSenha

varSucessoSenha = request.querystring("varSucessoSenha")


 dim varSucesso_bairro
 dim varExistente
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM telefone_acesso where cod_telefone_acesso="&varCod_telefone_acesso 
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
%>		








<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
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



</head>
<!--#include file="style2_sugestoes.asp"-->
<body bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="atualizar_telefone_acesso.asp?varCod_telefone_acesso=<%=varCod_telefone_acesso%>">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center"> 
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone de Acesso</font></div></td>
                  <td width="235" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone_acesso" type="text" class="inputBox" id="txt_telefone_acesso" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>" value="<%=rs("telefone_acesso")%>" size="38" maxlength="23" align="left"></td>
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
