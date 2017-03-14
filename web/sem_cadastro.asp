<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->

<%
dim varCodImovel

varCodImovel = request.querystring("varCodImovel")

%>

<html>
<head>
<title>Sem cadastro</title>




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





var strValidNumber1_4="1234567890";
for (nCount=0; nCount < nform.telefone.value.length; nCount++) 
		{
strTempChar1_4=nform.telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas números!");
nform.telefone.focus();
nform.telefone.select();
return false;
}
}





var strValidNumber1_5="a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,x,z,w,y,1,2,3,4,5,6,7,8,9,0,@,.,_,-";
for (nCount=0; nCount < nform.email.value.length; nCount++) 
		{
strTempChar1_5=nform.email.value.substring(nCount,nCount+1);
if (strValidNumber1_5.indexOf(strTempChar1_5,0)==-1) 
{
alert("Ao colocar seu email,use somente minúsculas!");
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





<!--#include file="style_imoveis02.asp"-->

</head>
<body onload=nform.nome.focus(); bgcolor="#FFFFFF" vlink="#48576C" link="#48576C" alink="#000000">
<div align="center"><br>
  
<table width="300" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
      <td><div align="center"><font color="#9d9249"><font face="Verdana" size="2"><strong>Por 
          favor, coloque seu nome, telefone e email, pois assim , você terá um 
          atendimento preferencial e exclusivo. </strong></font></font></div></td>
    </tr>
  </table>
<form action="mostrar_imovel2.asp?varCodImovel=<%=varCodImovel%>" Method="post" name="nform" onSubmit="return isValidDigitNumber(this);">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">
          <tr>
            <td width="244" align="center" bgcolor="#f7ecbf"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="#f7ecbf"> 
                  <td width="120" align="right" bgcolor="#f7ecbf"><font color="#9d9249" size="1" face="Verdana">Nome:</font></td>
                  <td bgcolor="#f7ecbf"> 
                    <input name="nome" type="Text" class="inputBox" id="nome" style="color:#9d9249;HEIGHT: 18px; WIDTH: 120px; background:<%=claro%>" size="15" ></td>
</tr>
                <tr bgcolor="#f7ecbf"> 
                  <td align="right" bgcolor="#f7ecbf"><font color="#9d9249" size="1" face="Verdana">Telefone:</font></td>
                  <td bgcolor="#f7ecbf"> 
                    <input name="telefone" type="text" class="inputBox" id="telefone" style="color:#9d9249;HEIGHT: 18px; WIDTH: 120px; background:<%=claro%>" size="15" ></td>
</tr>

<tr bgcolor="#f7ecbf"> 
                  <td align="right" bgcolor="#f7ecbf"><font color="#9d9249" size="1" face="Verdana">Email:</font></td>
                  <td bgcolor="#f7ecbf"> 
                    <input name="email" type="text" class="inputBox" id="email" style="color:#9d9249;HEIGHT: 18px; WIDTH: 120px; background:<%=claro%>" size="15" ></td>
</tr>

                <tr bgcolor="#f7ecbf"> 
                  <td bgcolor="#f7ecbf">&nbsp;</td>
                  <td align="right" bgcolor="#f7ecbf"> 
                    <input  type="submit" class="inputSubmit" value="Entrar" size="20" style="HEIGHT: 18px; WIDTH: 120px; background:#9d9249" ></td>
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