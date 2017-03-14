<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="loggedin.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div align="center"><br>
  <strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Verifica&ccedil;&atilde;o 
  do telefone no banco de dados</font></strong> <br>
  <br>
</div>
<form action="verifica_tudo01.asp?SearchFor=<%=SearchFor%>" onSubmit="return isValidDigitNumber2(this);" Method="GET" name="b2" >

<table border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="<%=claro%>">
          <tr>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>Procurar</b></font></td>
            <td bgcolor="<%=claro%>">
<input type="text" name="SearchFor" class="inputBox" value="<%=SearchFor%>" style=" background:<%=medio%>">
            </td>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>em</b></font></td>
            <td bgcolor="<%=claro%>">
	
	


<select name="SearchWhere" class="inputBox" style=" background:<%=medio%>">

<option value="telefone"  >Telefone</option>

</select>



















            </td>
            <td bgcolor="<%=claro%>">
<input type="submit" value="Buscar" class="inputSubmit" style="background:<%=medio%>;"></td>
</tr>
</table>
</form>

<div align="center"><br>
  <%


dim SearchWhere
dim SearchFor

SearchWhere = request.querystring("SearchWhere")
SearchFor = request.querystring("SearchFor")

session("SearchWhere") = SearchWhere
session("SearchFor") = SearchFor


dim Conexao


Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	 
   Conexao.Open dsn


dim rsCompradores22,strSQLCompradores22
   
	
	
	if  session("SearchFor") <> "" then
	strSQLCompradores22 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores FROM compradores where telefone like '%" & session("SearchFor") & "%' or telefone02 like '%" & session("SearchFor") & "%' or telefone03 like '%" & session("SearchFor") & "%'   ORDER BY cod_compradores" 
	else
	strSQLCompradores22 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores FROM compradores where cod_compradores like '0'  ORDER BY cod_compradores" 		
	end if
	
	
	
			 
Set rsCompradores22 = Server.CreateObject("ADODB.RecordSet")

	rsCompradores22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCompradores22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCompradores22.ActiveConnection = Conexao
	
	
	rsCompradores22.Open strSQLCompradores22, Conexao


if not rsCompradores22.eof then

'response.write "existe um registro de comprador"

%>
  <strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Existem 
  compradores com esse telefone</font></strong> </div>
  <br>
<table width="750" height="100" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50"> 
      <iframe src="verifica_compradores01.asp?SearchFor=<%=session("SearchFor")%>" name="meio" width="750px" height="100px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>





<%



end if

%>
<div align="center">
 
  <%
dim rsImovel22,strSQLImovel22
   
	
	
	if  session("SearchFor") <> "" then
	strSQLImovel22 = "SELECT imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.cod_imovel FROM imoveis where telefone like '%" & session("SearchFor") & "%' or telefone02 like '%" & session("SearchFor") & "%' or telefone03 like '%" & session("SearchFor") & "%'   ORDER BY cod_imovel" 
	else
	strSQLImovel22 = "SELECT imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.cod_imovel FROM imoveis where cod_imovel like '0'  ORDER BY cod_imovel" 		
	end if
	
	
	
			 
Set rsImovel22 = Server.CreateObject("ADODB.RecordSet")

	rsImovel22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsImovel22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsImovel22.ActiveConnection = Conexao
	
	
	rsImovel22.Open strSQLImovel22, Conexao


if not rsImovel22.eof then

%>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Existem 
  imóveis com esse telefone </strong></font> </div>
  <br>
<table width="750" height="100" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50"> 
      <iframe src="verifica_imoveis01.asp?SearchFor=<%=session("SearchFor")%>" name="meio" width="750px" height="100px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>



<%



end if

%>
<div align="center">
  <%
dim rsPermuta22,strSQLPermuta22
   
	
	
	if  session("SearchFor") <> "" then
	strSQLPermuta22 = "SELECT permuta.telefone,permuta.cod_permuta FROM permuta where telefone like '%" & session("SearchFor") & "%'   ORDER BY cod_permuta" 
	else
	strSQLPermuta22 = "SELECT permuta.telefone,permuta.cod_permuta FROM permuta where  cod_permuta like '0'  ORDER BY cod_permuta" 		
	end if
	
	
	
			 
Set rsPermuta22 = Server.CreateObject("ADODB.RecordSet")

	rsPermuta22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsPermuta22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsPermuta22.ActiveConnection = Conexao
	
	
	rsPermuta22.Open strSQLPermuta22, Conexao


if not rsPermuta22.eof then

%>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Existem permutas 
  com esse telefone </strong></font> </div>
  <br>
<table width="750" height="100" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="50"> 
      <iframe src="verifica_permuta01.asp?SearchFor=<%=session("SearchFor")%>" name="meio" width="750px" height="100px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>




<%


end if

%>

<%

if (rsCompradores22.eof and rsImovel22.eof and rsPermuta22.eof and session("SearchFor") <> "" ) then

response.redirect "form_incluir_compradores33.asp"

end if


%>
<br>
<br>



</body>
</html>
