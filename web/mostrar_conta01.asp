<!--#include file="cores.asp"-->
<!--#include file="style2_primeira.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>">
<%
dim varNome
dim varTelefone
dim varCodComprador
dim varCodImovel
dim varCodPermuta

varNome = request.querystring("varNome")
varTelefone = request.querystring("varTelefone")
varCodComprador = request.querystring("varCodComprador")
varCodImovel = request.querystring("varCodImovel")
varCodPermuta = request.querystring("varCodPermuta")

%>

<center>
<br>
<br>

<font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>A sua encomenda foi enviada com sucesso!</strong></font>

<% if varCodPermuta = "" then %>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">A 
  partir de agora você tem uma conta de comprador de imóveis, para verificar a 
  sua conta entre &eacute; s&oacute; clicar em entrar do formul&aacute;rio abaixo:</font><br>

<form action="acesso_compradores02.asp" onSubmit="return isValidDigitNumber(this);" Method="Post" name="doublecombo">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="<%=claro%>" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Nome:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="nome" value="<%=varNome%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Telefone:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="text" name="telefone" value="<%=varTelefone%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>

<tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Código do comprador:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="CodComprador" value="<%=varCodComprador%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
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


<%else%>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">A 
  partir de agora você tem três contas no nosso site , uma de comprador de imóveis, 
  outra de vendedor de imóveis e uma de permutante de imóveis ,para verificar 
  a sua conta basta clicar no bot&atilde;o entrar do formul&aacute;rio que te 
  interessa:</font> <br>
  <br>
  
<% end if %>
<br>
  <br>

<%

if varCodPermuta <> "" then
%>
<br>

<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Esse 
  primeiro formulário é o da conta de comprador</font> <br>

<form action="acesso_compradores02.asp" onSubmit="return isValidDigitNumber(this);" Method="Post" name="doublecombo">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="<%=claro%>" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Nome:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="nome" value="<%=varNome%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Telefone:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="text" name="telefone" value="<%=varTelefone%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>

<tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Código do comprador:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="CodComprador" value="<%=varCodComprador%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
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
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Esse 
  segundo formulário é o da conta de vendedor de im&oacute;vel</font><br>
<form action="acesso_imoveis03.asp" onSubmit="return isValidDigitNumber(this);" Method="Post" name="doublecombo">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="<%=claro%>" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Nome:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="nome" value="<%=varNome%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Telefone:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="text" name="telefone" value="<%=varTelefone%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>

<tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Código do imóvel:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="CodImovel" value="<%=varCodImovel%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
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
<br>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Esse 
  terceiro formulário é o da conta de permutante</font> <br>
<form action="acesso_permutantes01.asp" onSubmit="return isValidDigitNumber(this);" Method="Post" name="doublecombo">
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="<%=claro%>" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Nome:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="Text" name="nome" value="<%=varNome%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>
                <tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Seu Telefone:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input type="text" name="telefone" value="<%=varTelefone%>" size="15" class="inputBox" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" ></td>
</tr>

<tr bgcolor="<%=claro%>"> 
                  <td align="right" bgcolor="<%=claro%>"><font color="#FFFFFF" size="1" face="Verdana">Código da permuta:</font></td>
                  <td bgcolor="<%=claro%>"> 
                    <input name="CodPermutante" type="Text" class="inputBox" id="CodPermutante" style="HEIGHT: 18px; WIDTH: 120px; background:<%=medio%>" value="<%=varCodPermuta%>" size="15" ></td>
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
<br>
<br>
<%end if%>
</font>

</center>
</body>
</html>
