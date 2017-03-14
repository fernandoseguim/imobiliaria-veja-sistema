<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%


response.buffer=true
dim varCod_imovel

varCod_imovel = request.querystring("varCod_imovel")


 dim varSucesso_email
 dim varExistente
   
   dim rs
   Set rs = Server.CreateObject("ADODB.RecordSet")
    
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	
	dim strSQL
	strSQL = "SELECT * FROM imoveis where cod_imovel="&varCod_imovel
	 
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
		
}
}






{
if (nform.txtAssunto.value == "") {
		alert("Digite seu assunto?????.");
		nform.txtAssunto.focus();
		nform.txtAssunto.select();
		return false;
}
}

{
if (nform.txtMensagem.value == "") {
		alert("Digite sua mensagem.");
		nform.txtMensagem.focus();
		nform.txtMensagem.select();
		return false;
}
}


//-------------------------verifica se tem aspas no campo nome------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtNome.value.length; nCount++) 
		{
strTempChar=nform.txtNome.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Nome não pode conter aspas simples!")
nform.txtNome.focus();
nform.txtNome.select();
return false;
}
}


//-----------------------------------------------------------------------------

//-------------------------verifica se tem aspas no campo email------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtEmail.value.length; nCount++) 
		{
strTempChar=nform.txtEmail.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Email não pode conter aspas simples!")
nform.txtEmail.focus();
nform.txtEmail.select();
return false;
}
}


//-----------------------------------------------------------------------------

//-------------------------verifica se tem aspas no campo Sugestao------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtMensagem.value.length; nCount++) 
		{
strTempChar=nform.txtMensagem.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Mensagem não pode conter aspas simples!")
nform.txtMensagem.focus();
nform.txtMensagem.select();
return false;
}
}


//-----------------------------------------------------------------------------

//-------------------------verifica se tem aspas no campo Telefone------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtAssunto.value.length; nCount++) 
		{
strTempChar=nform.txtAssun?????to.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Assunto não pode conter aspas simples!")
nform.txtAssunto.focus();
nform.txtAssunto.select();
return false;
}
}


//-----------------------------------------------------------------------------





	
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
<body onload=b2.txtNome.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" action="atualizar_lastemail.asp?varCodCompradores=<%=varCodCompradores%>"  name="b2">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
    </tr>
	<tr>
      <td width="590" height="30">
<div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Último 
          email enviado</strong></font></div></td>
    </tr>
	
    <tr>
      <td width="590" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="590" height="18"><table width="590" height="18" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="18">&nbsp;</td>
            <td width="580" height="18">
<table width="580" height="18" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data do último email enviado</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" id="txtAssunto" value="<%=rs("dataLastEmail")%>" size="38" maxlength="30" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                    </div></td>
                </tr>
              </table>
            </td>
            <td width="5" height="18">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="140">&nbsp;</td>
            <td width="580" height="140"><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="290" height="218"><table width="289" height="218" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto do último email enviado</font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="200" bgcolor="<%=escuro%>"> 
                          <div align="center"></div></td>
                      </tr>
                    </table></td>
                  <td width="290" height="140"><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="290" height="200"><textarea name="txtMensagem" cols="32" rows="30" class="inputBox" id="txtMensagem" style="HEIGHT: 200px; WIDTH: 292px; background: <%=claro%>; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 800)"><%=rs("textoLastEmail")%></textarea></td>
                      </tr>
                      <tr>
                        <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="145" height="18">&nbsp;</td>
                              <td width="145" height="18">&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
            <td width="5" height="140">&nbsp;</td>
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
		   
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
