
<%

option explicit 
response.buffer=true

dim varSucesso

varSucesso = request.querystring("varSucesso")

dim varJaVendeu
dim varProprietario
dim varTelefone
dim varEmail


 varJaVendeu = request.querystring("varJaVendeu")
 varProprietario = request.querystring("varProprietario")
 varTelefone = request.querystring("varTelefone")
 varEmail = request.querystring("varEmail")

'varJaVendeu = "sim"
 'varProprietario = "Nico"
 'varTelefone = "Telefone"
 'varEmail = "wentznico@terra.com.br"

%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>


<title>Enviar email</title>
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
if (nform.txtTelefone.value == "") {
		alert("Digite seu telefone.");
		nform.txtTelefone.focus();
		nform.txtTelefone.select();
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






var strValidNumber1_5="a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,x,z,w,y,1,2,3,4,5,6,7,8,9,0,@,.,_,-";
for (nCount=0; nCount < nform.txtEmail.value.length; nCount++) 
		{
strTempChar1_5=nform.txtEmail.value.substring(nCount,nCount+1);
if (strValidNumber1_5.indexOf(strTempChar1_5,0)==-1) 
{
alert("Ao colocar seu email,use somente minúsculas!");
nform.txtEmail.focus();
nform.txtEmail.select();
return false;
}
}






var strValidNumber1_4="1234567890";
for (nCount=0; nCount < nform.txtTelefone.value.length; nCount++) 
		{
strTempChar1_4=nform.txtTelefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas números!");
nform.txtTelefone.focus();
nform.txtTelefone.select();
return false;
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

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>;}
</STYLE>

</head>

<!--#include file="style_imoveis.asp"-->
<!--#include file="cores02.asp"-->
<body onload=b2.txtNome.focus(); bgcolor="#f7ecbf" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" action="incluir_email.asp" onSubmit="return isValidDigitNumber(this);" name="nform">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0" bgcolor="#f7ecbf">
    <tr>
      <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></td>
    </tr>
    <tr>
      <td width="590" height="126"><table width="590" height="126" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="126">&nbsp;</td>
            <td width="580" height="126"><table width="580" height="126" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="580" height="18"><div align="center"><% if varSucesso = "" then %><% else %><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=varSucesso%></font> <% end if%></div></td>
                </tr>
                <tr>
                  <td width="580" height="90"><table width="580" height="90" border="0" cellpadding="0" cellspacing="0">
                      <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Empresa</font></div></td>
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Imobili&aacute;ria 
                            Veja </font></div></td>
                      </tr>
                      <tr bgcolor="<%=claro%>"> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></div></td>
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Av. 
                            Ant&aacute;rtico, 315 Jardim do Mar, S&atilde;o Bernardo</font></div></td>
                      </tr>
                      <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            3907-07-44 || 4123-72-44</font></div></td>
                      </tr>
                      <tr bgcolor="<%=claro%>"> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">veja@imobiliariaveja.com.br</font></div></td>
                      </tr>
                      <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Fax</font></div></td>
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Não informado</font></div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td width="580" height="18"><div align="center"></div></td>
                </tr>
              </table></td>
            <td width="5" height="126">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td width="590" height="54"><table width="590" height="54" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="54">&nbsp;</td>
            <td width="580" height="54"><table width="580" height="54" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" > 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      Nome </font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtNome" type="text" id="txtNome" value="<% if varProprietario <> "" then response.write varProprietario else response.write session("nome") end if%>" size="38" maxlength="50" align="left" class="inputBox" style="color:#9d9249;border-color:#f7ecbf;HEIGHT: 18px; WIDTH: 290px; background: #f7ecbf">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      Telefone</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtTelefone" type="text" id="txtTelefone" value="<% if varTelefone <> "" then response.write varTelefone else response.write session("telefone") end if%>" size="38" maxlength="50" align="left" class="inputBox" style="color:#9d9249;border-color:<%=claro%>;HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                    </div></td>
                </tr>
				
				
				 <tr> 
                  <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      Email </font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtEmail" type="text" id="txtEmail" value="<% if varEmail <> "" then response.write varEmail else response.write session("email") end if%>" size="38" maxlength="50" align="left" class="inputBox" style="color:#9d9249;border-color:#f7ecbf;HEIGHT: 18px; WIDTH: 290px; background: #f7ecbf">
                    </div></td>
                </tr>
				
				
                <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto 
                      </font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" id="txtAssunto" value="<% if varJaVendeu <> "" then response.write "Vendi meu imóvel!" else response.write "" end if%>" size="38" maxlength="50" align="left" class="inputBox" style="color:#9d9249;border-color:<%=claro%>;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>">
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
                        <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Mensagem</font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="122"><div align="center"></div></td>
                      </tr>
                    </table></td>
                  <td width="290" height="140"><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="290" height="122"><textarea name="txtMensagem" cols="32" rows="8" class="inputBox" id="txtMensagem" style="color:#9d9249;border-color:#f7ecbf;HEIGHT: 120px; WIDTH: 292px;  border: 1px solid #FFFFFF; background: #f7ecbf;"  OnKeyPress="return limitfield(this, 500)"></textarea></td>
                      </tr>
                      <tr>
                        <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="145" height="18"><input name="image" type="image" src="bt_enviar001.jpg" width="145" height="18"></td>
                              <td width="145" height="18"><a href="javascript:document.forms.b2.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
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
