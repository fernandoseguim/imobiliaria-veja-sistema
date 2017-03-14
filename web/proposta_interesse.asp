
<%

option explicit 
response.buffer=true

dim varSucesso
varSucesso = request.querystring("varSucesso")

%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Encomendar um imóvel </title>
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
		
	} else {
		prim = nform.txtEmail.value.indexOf("@")
		if(prim < 2) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("@",prim + 1) != -1) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(".") < 1) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(" ") != -1) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("zipmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("hotmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(".@") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("@.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(".com.br.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("/") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("[") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("]") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("(") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(")") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("..") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		
		
		
		}
		
		
	}
	//--------
	var strValidNumber="1234567890.,";
for (nCount=0; nCount < nform.txtTelefone.value.length; nCount++) 
		{
strTempChar=nform.txtTelefone.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)==-1) 
{
alert("O campo telefone deve ser numérico!")
nform.txtTelefone.focus();
nform.txtTelefone.select();
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
if (nform.txtProposta.value == "") {
		alert("Descreva o imóvel desejado.");
		nform.txtProposta.focus();
		nform.txtProposta.select();
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

//-------------------------verifica se tem aspas no campo proposta------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtProposta.value.length; nCount++) 
		{
strTempChar=nform.txtProposta.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Descrição do imóvel desejado não pode conter aspas simples!")
nform.txtProposta.focus();
nform.txtProposta.select();
return false;
}
}


//-----------------------------------------------------------------------------








	
}
	
	
	
	
		return true;
}












</script>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>



</head>
<!--#include file="style4_proposta.asp"-->
<!--#include file="cores.asp"-->
<body onload="document.b2.txtNome.focus()" bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" action="incluir_proposta_interesse.asp" onSubmit="return isValidDigitNumber(this);" name="b2">
<table width="590" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=escuro%>">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  <tr>
    <td width="590" height="18"><div align="center"><% if varSucesso ="" then response.Write varSucesso else response.write "<font size='1' color='#FFFFFF' face='Verdana, Arial, Helvetica, sans-serif'>"&varSucesso&"</font>" end if%></div></td>
  </tr>
  
  <tr>
    <td width="590" height="90"><table width="590" height="90" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5" height="90">&nbsp;</td>
          <td width="580" height="90"><table width="580" height="90" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome</font></div></td>
                <td width="290" height="18" bgcolor="#537497" style="border:1px solid #FFFFFF;">
<div align="center">
                      <input name="txtNome" type="text" id="txtNome" value="" size="38" maxlength="38" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                  </div></td>
              </tr>
              <tr bgcolor="7B9AB9"> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"> 
                  <div align="center">
                    <input name="txtTelefone" type="text" id="txtTelefone" value="" size="38" maxlength="12" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                  </div></td>
              </tr>
              <tr bgcolor="#537497"> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"> 
                  <div align="center">
                    <input name="txtEmail" type="text" id="txtEmail" value="" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                  </div></td>
              </tr>
              <tr bgcolor="7B9AB9"> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Hor&aacute;rio 
                      para contato</font></div></td>
                <td width="290" height="18" > 
                  <div align="center">
                    <select name="txtHorario" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 292px; background: <%=claro%>">
                      <option value="Qualquer Horário">Qualquer horário</option>
                      <option value="Manhã">Manhã</option>
                      <option value="Tarde">Tarde</option>
                      <option value="Noite">Noite</option>
                    </select>
                  </div></td>
              </tr>
              <tr bgcolor="#537497"> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                    de negocia&ccedil;&atilde;o</font></div></td>
                <td width="290" height="18" > 
                  <div align="center">
                      <select name="txtNegociacao" id="txtNegociacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 292px; background: <%=medio%>">
                        <option value="Qualquer Negociação">Qualquer Negociação</option>
                      <option value="Alugar um imóvel">Aluguel</option>
                      <option value="Comprar um imóvel">Compra</option>
                      
                    </select>
                  </div></td>
              </tr>
            </table></td>
          <td width="5" height="90">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5" height="140">&nbsp;</td>
            <td width="580" height="140"><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="290" height="122"><table width="289" height="122" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            do im&oacute;vel desejado</font></div></td>
                      </tr>
                      <tr>
                        <td width="290" height="104">&nbsp;</td>
                      </tr>
                    </table></td>
                  <td width="291"><div align="left"><textarea name="txtProposta" rows="8" cols="32"  OnKeyPress="return limitfield(this, 500)" style="HEIGHT: 120px; WIDTH: 292px; background: <%=claro%>; border: 1px solid #FFFFFF" class="inputBox"></textarea></div></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td><input name="image" type="image" src="bt_enviar001.jpg" width="145" height="18"></td>
                        <td><a href="javascript:document.forms.b2.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
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
</body>
</html>

