<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->

<%
dim varCodImovel

varCodImovel = request.querystring("varCodImovel")

%>

<html>
<head>
<title></title>




<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (nform) 

{













{
if (nform.nome.value == "") {
		alert("Digite seu nome.");
		nform.nome.focus();
		nform.nome.select();
		return false;
}
}









	//--------
	var strValidNumber="1234567890.,";
for (nCount=0; nCount < nform.telefone.value.length; nCount++) 
		{
strTempChar=nform.telefone.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)==-1) 
{
alert("O campo telefone deve ser numérico!")
nform.telefone.focus();
nform.telefone.select();
return false;
}
}
	
{
if (nform.telefone.value == "") {
		alert("Digite seu telefone.");
		nform.telefone.focus();
		nform.telefone.select();
		return false;
}
}


{
if (nform.email.value == "") {
		alert("Digite seu email.");
		nform.email.focus();
		nform.email.select();
		return false;
}
}










//------------- Verifica se é numérico---------------------
var elem=nform.elements;





for (nCount=0; nCount < elem.length; nCount++)
  
    
  
	
	if(elem[nCount].type.indexOf("text")==0)	{
	var strValidNumber12_1="'";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)!=-1) {



alert("Este campo  não pode conter aspas");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}

//-----------------------------------------------

}










</script>





<!--#include file="style2_primeira.asp"-->

</head>
<body onload=nform.nome.focus(); bgcolor="<%=escuro%>" vlink="#48576C" link="#48576C" alink="#000000">
<div align="center"><br>
  <br>
  <br>
  <br>
  <br>
  <font face="Verdana" size="2" color="#FFFFFF"><strong>Por favor, coloque seu 
  nome, telefone e email, pois assim , você terá um atendimento preferencial e 
  exclusivo. </strong></font> </div>
<form action="procurar_avaliacao.asp" Method="post" name="nform" onSubmit="return isValidDigitNumber(this);">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">
          <tr>
            <td width="244" align="center" bgcolor="<%=claro%>"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="<%=claro%>"> 
                  <td width="120" align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Nome:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input name="nome" type="Text" class="inputBox" id="nome" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" size="15" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Telefone:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input name="telefone" type="text" class="inputBox" id="telefone" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" size="15" ></td>
</tr>

<tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Email:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input name="email" type="text" class="inputBox" id="email" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" size="15" ></td>
</tr>

                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>">&nbsp;</td>
                  <td align="right" bgcolor="<%=claro%>"> 
                    <input  type="submit" class="inputSubmit" value="Entrar" size="20" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
</table>


</tr>
</table>
</td>
</tr>
</table>
</form>
<center>
</center>


</body>
</html>
<!--#include file="dsn2.asp"-->