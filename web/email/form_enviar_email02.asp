
<%

option explicit 
response.buffer=true



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
		alert("Digite seu assunto.");
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
strTempChar=nform.txtAssunto.value.substring(nCount,nCount+1);
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
<body onload=b2.txtNome.focus(); bgcolor="#406496" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" action="atirar.asp"  name="b2">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
    </tr>
    <tr>
      <td width="590" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="590" height="54"><table width="590" height="54" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="54">&nbsp;</td>
            <td width="580" height="54"><table width="580" height="54" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="#537497" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">D&ecirc;:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtDe" type="text" id="txtDe" value="acempreendimentos@acempreendimentos.com.br" size="38" maxlength="45" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: #537497;">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtPara" type="text" id="txtPara" value="" size="38" maxlength="35" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: 7B9AB9">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="#537497" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" id="txtAssunto" value="" size="38" maxlength="30" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: #537497">
                    </div></td>
                </tr>
              </table></td>
            <td width="5" height="54">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="140">&nbsp;</td>
            <td width="580" height="140"><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="290" height="140"><table width="289" height="140" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="290" height="18" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Mensagem:</font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="122"><div align="center"></div></td>
                      </tr>
                    </table></td>
                  <td width="290" height="140"><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="290" height="122"><textarea name="txtMensagem" cols="32" rows="8" class="inputBox" id="txtMensagem" style="HEIGHT: 120px; WIDTH: 292px; background: 7B9AB9; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 500)"></textarea></td>
                      </tr>
                      <tr>
                        <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="145" height="18"><input name="image" type="image" src="bt_enviar2.jpg" width="145" height="18"></td>
                              <td width="145" height="18"><a href="javascript:document.forms.b2.reset()"><img src="bt_apagar2.jpg" width="145" height="18" border="0"></a></td>
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
 <% response.flush%>
  <%response.clear%>

</body>
</html>
