
<%

option explicit 
response.buffer=true



%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin02.asp"-->
<!--#include file="cores.asp"-->
<%
Dim Conexao,strSQL,rs,intRecordCount,varCodOrigem,varSucesso_email

varCodOrigem = request.QueryString("varCodOrigem")

dim varSucessoOrigem

varSucessoOrigem= request.querystring("varSucessoOrigem")


 dim varSucesso_bairro
 dim varExistente
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM origem where id_origem="&varCodOrigem 
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
%>		








<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Visualizar origem</title>
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
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="atualizar_origem.asp?varCodOrigem=<%=varCodOrigem%>">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center"> 
          <% if varSucessoOrigem = "" then%>
          <%else%>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=varSucessoOrigem%> 
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome</font></div></td>
                  <td width="235" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><input name="txt_origem" type="text" class="inputBox" id="txt_origem" style="HEIGHT: 18px; WIDTH: 235px; background: <%=medio%>" value="<%=rs("origem")%>" size="38" maxlength="40" align="left"></td>
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
