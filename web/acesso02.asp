<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_imoveis02.asp"-->


<html>
<head>
<title></title>



<script>
function isValidDigitNumber (doublecombo)
{
{


var strValidNumber1_7="1234567890";
for (nCount=0; nCount < doublecombo.telefone.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário telefone só pode conter números!");
doublecombo.telefone.focus();
doublecombo.telefone.select();
return false;
}
}




if (doublecombo.nome.value == "") {
        alert("O formulário Seu Nome  está vazio!");
        doublecombo.nome.focus();
		doublecombo.nome.select();
        return false;
    }


if (doublecombo.telefone.value == "") {
        alert("O formulário Seu Telefone está vazio!");
        doublecombo.telefone.focus();
		doublecombo.telefone.select();
        return false;
    }

if (doublecombo.CodComprador.value == "") {
        alert("O formulário Código de referência está vazio!");
        doublecombo.CodComprador.focus();
		doublecombo.CodComprador.select();
        return false;
    }




}
}









</script>









</head>
<body onload=doublecombo.nome.focus(); bgcolor="#FFFFFF" vlink="#48576C" link="#48576C" alink="#000000">
<br>
<center>
<table width="710" height="60" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td style="border:1px solid <%=claro%>;"> 
        <table width="700" height="50" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="<%=escuro%>"> 
              <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acesso01.asp" style="color:#FFFFFF;text-decoration:none;">Sua 
                conta de permutante</a></strong></font></div></td>
            <td bgcolor="<%=claro%>"> 
              <div align="center"><font color="<%=escuro%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acesso03.asp" style="color:<%=escuro%>;text-decoration:none;">Sua 
                conta de vendedor de im&oacute;veis</a></strong></font></div></td>
            <td bgcolor="<%=escuro%>"> 
              <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acesso02.asp" style="color:#FFFFFF;text-decoration:none;">Sua 
                conta de comprador</a></strong></font></div></td>
  </tr>
</table>
      </td>
  </tr>
</table>
</center>
<div align="center"><br>
  <br>
  <table width="400" height="80" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="400">
<div align="center"><font color="<%=escuro%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Para 
          verificar se existem im&oacute;veis indicados para voc&ecirc;, digite 
          seu telefone e clique em entrar.</strong></font></div></td>
    </tr>
  </table>
  <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></strong> 
  <br>
  <br>
  <font face="Verdana" size="1" color="<%=escuro%>"> 
  <%
	If Request("SecondTry") = "True" Then
		If Request("WrongPW") = "True" Then
			Response.Write "<center>Dados inválidos. Tente novamente.</center>"
		Else
			Response.Write "<center>Dados inválidos.  Tente novamente.</center>"
		End If
	End If
%>
  </font> </div>
<form action="acessoLink02.asp" onSubmit="return isValidDigitNumber(this);" Method="Post" name="doublecombo">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>" style="border:1px solid <%=claro%>;"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="<%=claro%>" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
               
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="<%=escuro%>" size="1" face="Verdana"><strong>Seu 
                    Telefone:</strong></font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="text" name="telefone" value="<%=session("varTelefone")%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=escuro%>;" ></td>


                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>">&nbsp;</td>
                  <td align="right" bgcolor="<%=claro%>"> 
                    <input  type="submit" class="inputSubmit" value="Entrar" size="20" style="HEIGHT: 18px; WIDTH: 120px; background:<%=escuro%>;" ></td>
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