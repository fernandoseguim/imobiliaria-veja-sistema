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
<br>
<br>
<br>
<font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>O seu imóvel foi enviado com sucesso!</strong></font>
<br>

<% if varCodPermuta = "" then %>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">A 
  partir de agora você tem uma conta de vendedor de imóvel, nessa conta voc&ecirc; 
  poder&aacute;<br>
  verificar poss&iacute;veis compradores para o seu im&oacute;vel , nosso sistema 
  cruza informa&ccedil;&otilde;es e <br>
  e coloca a sua disposi&ccedil;&atilde;o uma lista de poss&iacute;veis interessados 
  em fazer neg&oacute;cio com voc&ecirc;.</font><br>
<br>
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


<%else%>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">A 
  partir de agora você tem três contas no nosso site , uma de comprador de imóveis, 
  outra de vendedor de imóveis e uma de permutante de imóveis ,nessas contas voc&ecirc; 
  poder&aacute; encontrar uma lista de interessados em fazer neg&oacute;cio com 
  voc&ecirc;:</font> <br>
<% end if %>
<br>

<br>

<%

if varCodPermuta <> "" then
%>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Esse 
  primeiro formulário é a entrada para a sua conta de vendedor de um im&oacute;vel 
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
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Esse 
  segundo formulário é a entrada da sua conta de comprador de um im&oacute;vel</font> 
  <br>

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
<br>
<br>
  <font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Esse 
  terceiro formulário é a entrada para a sua conta de permutante de im&oacute;vel</font> 
  <br>
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
