<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="loggedin02.asp"-->
<!--#include file="cores.asp"-->



</head>
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<table width="800" height="18" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td bgcolor="<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_procurados.asp">Im&oacute;veis 
          procurados</a></strong></font></div></td>
      <td bgcolor="<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_referencia_procurados.asp">Refer&ecirc;ncias 
          procuradas</a></strong></font></div></td>
      <td bgcolor="<%=medio%>"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_procurados.asp">Compradores 
          procurados</a></strong></font></div></td>
      <td bgcolor="<%=claro%>"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_procurados.asp">Permutantes 
          procurados</a></strong></font></div></td>
		  
      <td bgcolor="<%=medio%>"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_senha.asp">Senhas 
          do sistema</a></strong></font></div></td>
		   
    <td bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ip.asp">IPs 
        do sistema</a></strong></font></div></td>
    </tr>
	

	
  </table>
<br>

<table width="460" height="18" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
   <td width="153" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_origem.asp">Origem</a></strong></font></div></td>
      
    <td width="153" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imovel_clicado.asp">Im&oacute;veis 
        clicados </a></strong></font></div></td>
      
    <td width="153" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_contas_procuradas.asp">Contas 
        acessadas </a></strong></font></div></td>
 
  <td width="153" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_tipo.asp">Tipos de im�veis</a></strong></font></div></td>
 
 
 
  </tr>
</table>
<center>
<br>
  <a href="excluir_imovel_clicado.asp"><img src="bt_excluir002.jpg" width="95" height="18" border="0"></a><br>
  <br>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lista de 
  im&oacute;veis clicados por internautas</strong></font> 
</center>


<center>
</center>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere
dim varCod_imovel

SearchWhere = request.querystring("SearchWhere")
SearchFor = request.querystring("SearchFor")

session("SearchWhere") = SearchWhere
session("SearchFor") = SearchFor


Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio
color2 = claro




Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

  

SQL ="SELECT imovel_clicado.cod_clicado,imovel_clicado.nome,imovel_clicado.telefone,imovel_clicado.codigo_clicado,imovel_clicado.endereco_ip,imovel_clicado.data,imovel_clicado.tipo,imovel_clicado.quartos,imovel_clicado.vagas,imovel_clicado.cidade,imovel_clicado.bairro,imovel_clicado.valor,imovel_clicado.negociacao FROM imovel_clicado ORDER BY cod_clicado DESC" 
	
 





%>
<%

Dim LinkTemp
'essa vari�vel vai ser usada como contador


'as vari�veis acima s�o usadas para trocar a cor das tabelas que conter�o os valores
'dos recordsets.






dim intPage
'essa vari�vel vai receber um valor inicial "1" que mostra que estamos na primeira p�gina.

dim intPageCount
'Essa vari�vel vai receber o valor da quantidade de p�ginas do recordset.

dim intRecordCount
'Essa vari�vel vai receber o n�mero de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a vari�vel intPage recebe o valor "1" na primeira p�gina.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conex�o o recordset utilizar�.
	
RS.Open SQL, Conn, 1, 3
'o recordset � aberto
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por p�gina.

RS.CacheSize = RS.PageSize
'o Cache tamb�m conter� 20 registros por p�gina.

intPageCount = RS.PageCount
'A vari�vel intPageCount recebe o valor do n�mero de p�gina do recordset retornado.

intRecordCount = RS.RecordCount
'A vari�vel intRecordCount recebe o valor do n�mero de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%><br>
<center><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%rs.movefirst%><strong><%=rs("data")%></strong> at�  <%rs.movelast%><strong><%=rs("data")%></strong><br><br> foram <strong><%=rs.RecordCount%></strong> acessos.</font></center>
<form  Method="Post" name="Formulario" action="multi_excluir_imovel.asp?varCod_imovel=<%=varCod_imovel%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&page=<%=cInt(intPage)%>" >
  <table width="1200" border="0" cellspacing="0" cellpadding="0">
    <tr bgcolor="#000000"> 
      <td width="115" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome 
          do internauta</strong></font></div></td>
	  <td width="115" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone 
          do internauta</strong></font></div></td>
      
	  <td width="115" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>C�digo 
          de refer�ncia do im&oacute;vel</strong></font></div></td>
      
     
	  
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>IP 
          do internauta</strong></font></div></td>
      
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
          de acesso</strong></font></div></td>
    
	 <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade 
          </strong></font></div></td>
      
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div></td>
     
	 <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo 
          </strong></font></div></td>
      
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negociacao</strong></font></div></td>
    
	 <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></div></td>
      
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></div></td>
    
	  <td width="115" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
    
	
	
	
	</tr>
    <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage � maior que o n�mero de p�ginas ent�o intPage � igual ao n�mero de p�ginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage � menor ou igual a zero ent�o intPage igual a "1"
	'a vari�vel intPage sempre vai ser for�ada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados ent�o.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a p�gina exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a vari�vel intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posi��o exata do primeiro registro da p�gina correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage � igual ao n�mero de p�ginas no recordset , estamos na �ltima 
			'p�gina ent�o.
				intFinish = intRecordCount
				'a vari�vel intFinish recebe o valor do n�mero do �ltimo recordset.
				'intFinish corresponde ao valor do �ltimo registro da p�gina correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a vari�vel intFinish recebe o valor de intStart + o valor
				'do n�mero de registros na p�gina menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros ent�o
		For intRecord = 1 to RS.PageSize
		'um contador inRecord � colocado at� o n�mero de registros na p�gina.
%>
    <%










'------------------------------------------------

%>
    <%
If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If

dim vValor


 session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)


%>
    <% session("page")=intPage%>
   
    <tr> 
      <td width="115" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("nome")%></strong></font></div></td>
	  <td width="115" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("telefone")%></strong></font></div></td>
	  
       <td width="115" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('visualizar_imovel_clicado.asp?varCod_imovel=<%=rs("codigo_clicado")%>')"><%=rs("codigo_clicado")%></a></strong></font></div></td>
	 
	  <td width="115" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("endereco_ip")%></strong></font></div></td>
	 
      
		  
		  
      
      
	  
	  <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("data")%></font></div></td>
     
	  <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cidade")%></font></div></td>
    
	 <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("bairro")%></font></div></td>
     <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if rs("tipo") <> "tqualquer" then response.write rs("tipo") else response.write "qualquer um" end if%></font></div></td>
     <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("negociacao")%></font></div></td>
    
 <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("quartos")%></font></div></td>
     <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("vagas")%></font></div></td>
     <td width="115" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=formatNumber(rs("valor"),2)%></font></div></td>
    
		
	
	
	</tr>
    <%
'-----------------------------------------------









rs.movenext
If RS.EOF Then Exit for
Next

%>
  </table>
</form>





<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#003366"> 
        <%If cInt(intPage) > 1 Then%>
        <!-- se a p�gina atual for maior que "1" ent�o o link anteriro � colocado na 
			  na tela .-->
      <a href="?page=<%=intPage - 1%>" style="color:#000000">
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se p�gina atual � menor que o total de p�ginas e intPage maior que um
			  ou seja, se n�o estiver na primeira p�gina e nem na �ltima ent�o. -->
        <font color="#000000">P�gina</font> <%=cInt(intPage)%> <font color="#000000">de</font> 
        <%=cInt(intPageCount)%> </font> 
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage � menor que o n�mero de p�ginas ent�o colocar o bot�o pr�ximo -->
        <a href="?page=<%=intPage + 1%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Pr�ximo</b></font> 
        </a> 
        <%End If%>
        </font></div></td>
        </tr>
      </table>










 <%else%>
 
  <table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
 </table>
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Im&oacute;veis n�o encontrados</font></div>
</font> 
<%
End If%>
<%else%>
<table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
	
 </table>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Im&oacute;veis </font><font color="<%=escuro%>"> 
  n�o encontrados</font></div>
</font> 
<%
End if
%>
   
<%
  rs.Close
           'fecha a conex�o
           Conexao.Close
           Set rs = Nothing
		   
		   set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>
  <br>
  <br>
  <center>

</center>
</body>
</html>

