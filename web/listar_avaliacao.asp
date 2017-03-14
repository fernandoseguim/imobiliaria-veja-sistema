<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%response.Buffer = true %>


<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")


 
 '------------------Vamos verificar se o telefone existe, caso não--------------------
 '------------------não exista, esse cliente e incluído no banco de dados-----------
 



if session("nome") = "" then

session("nome") = request.querystring("nome")

end if


if session("email") = "" then

session("email") = request.querystring("email")

end if



if session("telefone") = "" then

session("telefone") = request.querystring("telefone")

end if








  
   
   Set rsAvaliacao = Server.CreateObject("ADODB.RecordSet")
    
	strSQLAvaliacao = "SELECT * FROM imoveis Where telefone = '"&session("telefone")&"'"
	 
   
   
rsAvaliacao.CursorLocation = 3
rsAvaliacao.CursorType = 3

        rsAvaliacao.Open strSQLAvaliacao, Conexao3 
		
		
		
		
	
	







 
 
 
 
 '---------------------------------------------------------------------------------
 
 
 
 
 
 
 
 






   vValor=request.form("stage22")
   session("vValor")=vValor
   
   if session("vValor") = "" then
session("vValor") = request.querystring("vValor")
end if
   
   session("vValor1")=left(session("vValor"),10)
   session("vValor2")=right(session("vValor"),10)
   
   






dim vnegociacao

 vnegociacao=request.form("example2")
 session("vnegociacao") = vnegociacao
 
 if session("vnegociacao") = "" then
 session("vnegociacao") = request.querystring("vnegociacao")
 end if
 

dim vOcupacao
 
 vOcupacao = request.Form("txt_ocupacao")
 session("vOcupacao") = vOcupacao
 
 if session("vOcupacao") = "" then
session("vOcupacao") = request.querystring("vOcupacao")
end if



dim vVagas
 
 vVagas = request.Form("txt_garagem")
 session("vVagas") = vVagas
 
 if session("vVagas") = "" then
session("vVagas") = request.querystring("vVagas")
end if


dim vQuartos
 
 vQuartos = request.Form("txt_quartos")
 session("vQuartos") = vQuartos
 
 if session("vQuartos") = "" then
session("vQuartos") = request.querystring("vQuartos")
end if

 vTipo=request.form("txt_tipo")
  session("vTipo") = vTipo
  
  if session("vTipo") = "" then
  session("vTipo") = request.querystring("vTipo")
  end if
  
  

dim vCidade2

 vCidade2=request.form("combo1")
	
	
	
	session("vCidade2") = vCidade2
	 if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")

end if


dim vBairro2
	 vBairro2=request.form("combo2")
	 session("vBairro2") = vBairro2
	 if session("vBairro2") = "" then
session("vBairro2") = request.querystring("vBairro2")

end if


dim vvila2,vvila
vvila2=request.form("combo5")
	
	
	
	session("vvila2") = vvila2
	 if session("vvila2") = "" then
session("vvila2") = request.querystring("vvila2")

end if


dim Nome
 dim Telefone
 
 Nome = request.form("txt_nome")
 
 
 session("nome") = Nome
 
 
 Telefone = request.form("txt_telefone")
 

 session("telefone") = Telefone

if session("nome") = "" then
session("nome") = request.querystring("nome")

end if


if session("telefone") = "" then
session("telefone") = request.querystring("telefone")

end if




dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	if session("vCidade2") <> "cqualquer" then
	
	strSQL4 = "SELECT * FROM combo2 where id_combo1 ="&session("vCidade2")&"  ORDER BY nome_combo2" 
	else
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	end if
	
	Conexao.Open dsn
	
	rs4.Open strSQL4, Conexao


dim rs44,strSQL44
   
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	
	if session("vBairro2") <> "bqualquer" then
	strSQL44 = "SELECT * FROM combo3 where id_combo2 ="&session("vBairro2")&"  ORDER BY nome_combo3" 
	else
	strSQL44 = "SELECT * FROM combo3   ORDER BY nome_combo3"
	end if 
	
	
	
	
	rs44.Open strSQL44, Conexao




%>






<%

'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open dsn

'

Sql333 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs333 = Conexao333.Execute ( Sql333 ) 



 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
%> 





<html>

<!--#include file="style4_imoveis.asp"-->

<title>Avaliação</title>
<head>
<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (doublecombo) 



{




{


if (doublecombo.txt_nome.value == "Seu nome:") {
		alert("Por favor,deixe seu nome na busca , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}

if (doublecombo.txt_nome.value == "") {
		alert("Por favor,deixe seu nome na busca , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}









if (doublecombo.txt_telefone.value == "Seu telefone:") {
		alert("Por favor, coloque seu telefone , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_telefone.value == "") {
		alert("Por favor, coloque seu telefone , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}





var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}



if (doublecombo.combo1.value == "cqualquer") {
		alert("Você precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}



if (doublecombo.example2.value == "nqualquer") {
		alert("Por favor, escolha um tipo de negociação , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.example2.focus();
		
		return false;
}












}
}


</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=520,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>

<script language="javascript">
function funScroll()
{
window.scrollTo(0,0)

}		
</script>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>

</head>












<body  bgcolor="EAA813" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">


<br>
<br>







<br>
<%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2




 
  
   
  
   
   if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")
end if
   
   
   














	  Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	if session("vCidade2") <> "cqualquer" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select * from combo1 where id_combo1 ="&session("vCidade2")
 
 rs2.open SQL2,Conexao,2,1
 
 vCidade = rs2("nome_combo1")
 else
 vCidade = vCidade2
 end if

	session("vCidade")= vCidade
	
	
	
	
	 
	 if session("vBairro2") <> "bqualquer" then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& session("vBairro2")
 
 rs3.open SQL3,Conexao,2,1

 vBairro = rs3("nome_combo2")
 else
 vBairro = vBairro2
	end if                                      
									
	 
	 
	 
	 session("vBairro")= vBairro
	 
	 
	 
	 
	 
	 
	 
	 if session("vvila2") <> "vlqualquer" and session("vvila2") <> "" then
	
	dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select * from combo3 where id_combo3 ="&session("vvila2")
 
 rs22.open SQL22,Conexao,2,1
 
 vvila = rs22("nome_combo3")
 else
 vvila = vvila2
 end if

	session("vvila")= vvila
	 
	 
	 
	 
	 
	 
	 
	  
	  
	 if session("vCidade")="sao bernardo" then
	  session("vCidade")="São Bernardo"
	  end if
	 
	   if session("vCidade")="santo andre" then
	  session("vCidade")="Santo André"
	  end if
	  
	   if session("vCidade")="sao caetano" then
	  session("vCidade")="São Caetano"
	  end if
	  
	  
	  if session("vBairro")="bairro assuncao" then
	  session("vBairro")="Bairro Assunção"
	  end if
	  
	   if session("vBairro")="ceramica" then
	  session("vBairro")="Cerâmica"
	  end if
	  
	  
	  if session("vBairro")="jd sao caetano" then
	  session("vBairro")="JD São Caetano"
	  end if 
	  'Acima as variáveis recebem os valores dos formulários para fazer busca.
	
'Codificar para receber "qualquer um" vcidade +2 pois vBairro está conectado a vcidade
'----------------------------------------------------------
if session("vCidade") = "" then
session("vCidade") = request.querystring("vCidade")
end if

if session("vBairro") = "" then
session("vBairro") = request.querystring("vBairro")
end if



if session("vTipo") = "" then
session("vTipo") = request.querystring("vTipo")
end if

if session("vNegociacao") = "" then
session("vNegociacao") = request.querystring("vNegociacao")
end if

 '---------------------------------------------------  
 
 
 
 '------------------pegar número de quartos------------------
 
 
 '--------------------------------------------------------
 
 
 '--------------------------Pegar número de Vagas na Garagem-------------
 
 
 
 
 
 
'----------------------------------------------------------------









'----------------------Ocupação do imóvel--------------------------

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 '------------------------------------------------------------------
 
 
 
 
 '----------------------Cidade--------------------------------
 dim stringIndex
 
stringIndex = " where cod_imovel<>"&"0"&"" 

if  session("vCidade") <> "cqualquer" then
stringCidade = "and cidade='"& session("vCidade")&"'"
else
stringCidade = ""
end if
 '--------------------------Bairro----------------------------

if session("vBairro") <> "bqualquer"  then
stringBairro = " and Bairro='"&session("vBairro")&"'"
else
stringBairro = "" 
end if


'--------------------------Vila----------------------------










'------------------------------------------------------------------------------------------

'------------------Tipo-----------------------------

dim StringTipo

if session("vTipo") <> "tqualquer" then
stringTipo = " and Tipo='"&session("vTipo")&"'"
else
stringtipo = ""
end if







'------------------------------------------------------




'----------------------------Vagas na garagem----------------


dim StringVagas

if session("vVagas") <> "gqualquer" then
stringVagas = " and vagas ="&session("vVagas")&""
else
stringVagas = ""
end if

'---------------------------------------------------------------









'---------------------------Ocupação------------------------










'--------------------------------------------------------------





'-------------------Negociação---------------------------

dim stringNegociacao


stringNegociacao = " and negociacao ='"&"venda"&"'"
'------------------------------------------------------------------------------


'---------------------------Quartos------------------------------


if  session("vQuartos") <> "qqualquer" then
stringQuartos = " and quartos ="&int(session("vQuartos"))&""
else
stringQuartos = ""

end if



'----------------------------pegar área total-------------------


dim vArea_Total



vArea_Total = request.form("txt_area_total")


session("vArea_Total") = vArea_Total

dim vArea_Total_Minima
dim vArea_Total_Maxima
dim vPorcentual

vPorcentual = int(vArea_Total)*10/100

vAreaTotal_Minima = int(vArea_Total) - int(vPorcentual)

vAreaTotal_Maxima = int(vArea_Total) + int(vPorcentual)



dim string_Area_Total


if  int(session("vArea_Total")) <> 0  then
string_Area_Total = " and area_total >="& int(vAreaTotal_Minima) &" and area_total <="& int(vAreaTotal_Maxima) &""
else
string_Area_Total = ""
end if




'-------------------------------------------------------------


'---------------------------------Valor-----------------------------------



dim stringValor





dim total


strSQL ="SELECT SUM(valor) as total  FROM Imoveis"&stringIndex&stringCidade&stringBairro&stringTipo&stringNegociacao&stringVagas&stringQuartos&string_Area_Total&"  and valor <> '0' "




'----------------------incluir em imóveis caso o telefone não exista---------------

	
	if rsAvaliacao.eof then
	
	dim vimagem 
	
	vimagem = "imovel00000.jpg"
	
	
 Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,qualidade,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,data_futuro_contato,assunto_futuro_contato,telefone02,telefone03,suites,chaves_do_imovel,melhor_horario_visita,imovel_em_negociacao,metros_de_frente,metros_de_fundo,metros_lateral_esquerda,metros_lateral_direita,origem_captacao,responsavel_cadastramento,data_ultimo_acesso) values( '"& session("nome") &"','"& "não informado" &"','"& session("telefone") &"','"& session("email") &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "icon_foto2.gif" &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& int(session("vArea_Total")) &"','"& "0" &"','"& session("vQuartos") &"','"& "0" &"','"& session("vVagas") &"','"& "venda" &"','"& int("0") &"','"& now() &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"&"não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "Avaliação" &"','"& now() &"','"& "não informado"&"','"& "não informado" &"','"& "0"&"','"&"não informado"&"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "0/0/2007 00:00:00" &"','"& "não informado" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "internet" &"','"& "internet" &"','"& now() &"')"

 
	
	
	
	
	end if
	







'----------------------------------------------------------------------------------







'----------------------------------------------------Fim da instrução SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 
  
  
  '------------------Verifica se o internauta já tem conta---------------------------
  

  
  
  if session("vVagas") = "gqualquer" then
 vVagasConta = "0"
 else
 vVagasConta = session("vVagas")
 end if
 
  if session("vQuartos") = "qqualquer" then
 vQuartosConta = "0"
 else
 vQuartosConta = session("vQuartos")
 end if
 

 session("vValorMedio") = vValorMedio
 
 
 session("vQuartosConta") = vQuartosConta
 session("vVagasConta") = vVagasConta
 session("vValorConta") = vValorConta
 
 
 dim vNegociacaoConta
 
 
 
 
 
 
 if session("vNegociacao") = "Venda" then
 vNegociacaoConta = "compra"
 
else
 
 
 vNegociacaoConta = session("vNegociacao")
 end if
 
 
 if session("vNegociacao") = "Aluguel" then
 vNegociacaoConta = "aluguel"
 end if
 
 
 
 session("vNegociacaoConta") = vNegociacaoConta
  

	
	
	
	
	
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 



Set RS = Server.CreateObject("ADODB.Recordset")
'um objeto recordset é instânciado.

Dim LinkTemp
'essa variável vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
'as variáveis acima são usadas para trocar a cor das tabelas que conterão os valores
'dos recordsets.






dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeira página.

dim intPageCount
'Essa variável vai receber o valor da quantidade de páginas do recordset.

dim intRecordCount
'Essa variável vai receber o número de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a variável intPage recebe o valor "1" na primeira página.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	
RS.Open strSQL, Conn, 1, 3
'o recordset é aberto
	
RS.PageSize = 8
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.


'verifica se existem registros retornados.

if not rs.eof then
%>
<%






	
%>
<%



Set rs2 = Server.CreateObject("ADODB.Recordset")

SQL2 = "SELECT *FROM Imoveis"&stringIndex&stringCidade&stringBairro&stringTipo&stringNegociacao&stringVagas&stringQuartos&string_Area_Total&" AND StandBy='"&"excluido"&"' and valor <> '0' "


	
RS2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS2.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	



	
RS2.Open SQL2, Conn, 1, 3
'o recordset é aberto
	
RS2.PageSize = 8
'Aqui configura-se o recordset para 20 registros por página.

RS2.CacheSize = RS2.PageSize
'o Cache também conterá 20 registros por página.



dim intPageCount2
intPageCount2 = RS2.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.


dim intRecordCount2
intRecordCount2 = RS2.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.



total = rs("total")

dim subTotal



subTotal = int(total)/int(intRecordCount2)

if not rs2.eof then

%>
<table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="120">
<center>
<strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">O valor aproximado 
do seu imóvel é <br>
<br><br><br><br>
        <font size="6" face="Verdana, Arial, Helvetica, sans-serif"><%=formatnumber(subTotal,2)%></font></font></strong> 
        <br>
		<br>
		<br>
		
		
		
		<br>
        <strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Esse 
        valor não é exato, ele é obtido em comparação com os outros imóveis cadastrados 
        no sistema. <br>
        <br>
        OBS: Podem ocorrem variações no valor encontrado, assim para uma avalia&ccedil;&atilde;o 
        mais criteriosa procure um dos nossos representantes.<br>
        </font></strong> 
      </center></td>
  </tr>
</table>


<%


else
%>
<center>
<table width="250" height="120" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&atilde;o 
          foi poss&iacute;vel avaliar o seu im&oacute;vel , pois o sistema n&atilde;o 
          encontrou outros im&oacute;veis para fazer uma compara&ccedil;&atilde;o 
          com o seu. Procure um de nossos representantes para fazer uma avalia&ccedil;&atilde;o 
          mais criteriosa. </strong></font></div></td>
  </tr>
</table>
</center>
<%

end if

Else

%>
não há retorno 
<% end if %>
<%


RS.close
Set RS = Nothing




Conexao3.close
set conexao3 = nothing

Conexao333.close
set conexao333 = nothing

%>
<% response.flush%>
<%response.clear%>
<!--#include file="dsn2.asp"-->

<br>
  
 </table>
 
</body>
</html>
