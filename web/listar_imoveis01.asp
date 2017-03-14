<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_imoveis.asp"-->
<%response.Buffer = true %>



<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if


dim SqlImoveis001
 dim rsImoveis001


'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql303 = "SELECT * FROM combo1 ORDER BY nome_combo1 ASC" 

Set rs303 = Server.CreateObject("ADODB.RecordSet")

	rs303.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs303.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs303.ActiveConnection = Conexao3
	
	
	rs303.Open sql303, Conexao3




	
	
	
	
	
	
	
	'--------------------------------------------------------------------
	
dim vCidade2

 vCidade2=request.form("combo1")

session("vCidade2") = vCidade2
	 if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")

end if
%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")


dim rs404,strSQL404,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs404 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL505
	
	if session("vCidade2") <> "cqualquer" then
	
	strSQL404 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 ="&session("vCidade2")&"  ORDER BY nome_combo2" 
	else
	strSQL404 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	end if
	
	
	
	'strSQL404 = "SELECT * FROM combo2 where id_combo1 = 4  ORDER BY id_combo2" 
	
	
	Conexao.Open dsn
	
	
	
	rs404.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs404.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs404.ActiveConnection = Conexao
	
	
	rs404.Open strSQL404, Conexao

dim rs55
dim strSQL55

Set rs55 = Server.CreateObject("ADODB.RecordSet")
	strSQL55 = "SELECT * FROM imoveis ORDER BY cod_imovel DESC" 



rs55.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs55.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs55.ActiveConnection = Conexao

rs55.open strSQL55, Conexao


%>





<%


'

Sql333 = "SELECT * FROM combo2 ORDER BY id_combo2" 
Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs33.ActiveConnection = Conexao3
	
	
	rs33.Open sql333, Conexao3



dim rsFrontPage,SQLFrontPage,objFSO 

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Set rsFrontPage = Server.CreateObject("ADODB.RecordSet")

SQLFrontPage = "SELECT * FROM imoveis where presenca_primeira like '"&"incluido"&"' ORDER BY cod_imovel DESC"

rsFrontPage.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsFrontPage.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsFrontPage.ActiveConnection = Conexao


rsFrontPage.open SQLFrontPage,Conexao

dim intRecordCount01 


intRecordCount01 = rsFrontPage.RecordCount




'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT * FROM tipo  ORDER BY tipo ASC" 
	
	
	rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao
	
	
	
	 rs444Tipo22.Open strSQL444Tipo22, Conexao










'-------------------------------------------------------------------------------------------------








%> 


<%


varNotFind = request.QueryString("varNotFind")


 






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
  
  


	
	
	
	


dim vBairro2
	 vBairro2=request.form("combo2")
	 session("vBairro2") = vBairro2
	 if session("vBairro2") = "" then
session("vBairro2") = request.querystring("vBairro2")

end if





dim Nome
 dim Telefone
 dim email
 
 Nome = request.form("txt_nome")
 
 
 session("nome") = Nome
 
 
 Telefone = request.form("txt_telefone")
 

 session("telefone") = Telefone


email = request.form("txt_email")
 

 session("email") = email



if session("nome") = "" then
session("nome") = request.querystring("nome")

end if


if session("telefone") = "" then
session("telefone") = request.querystring("telefone")

end if

if session("email") = "" then
session("email") = request.querystring("email")

end if


'-------------------Incluir nos clientes online-------------------
dim Sql001
dim rs001

Sql001 = "SELECT cod_compradores,telefone,atendimento FROM compradores where telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"' ORDER BY cod_compradores ASC" 

Set rs001 = Server.CreateObject("ADODB.RecordSet")

	rs001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs001.ActiveConnection = Conexao
	
	
	rs001.Open sql001, Conexao

dim atendimento001
if not rs001.eof then
atendimento001 = rs001("atendimento")
else
atendimento001= "não informado"
end if


dim vCod_cliente001
if not rs001.eof then
vCod_cliente001 = rs001("cod_compradores")
else
vCod_cliente001= "0"
end if


dim data001,data_full001
dim hora001, dia001, mes001, ano001 

hora001 = hour(now())
dia001 = day(now())
mes001 = month(now())
ano001 = year(now())

data001 = now()

data_full001 = hora001&"/"&dia001&"/"&mes001&"/"&ano001


if vCod_cliente001 <> "0" then

dim rsVerifica001
dim SqlVerifica001

'verificar se já existe esse cliente online
SqlVerifica001 = "SELECT * FROM cliente_online where  telefone like '"&session("telefone")&"' ORDER BY cod_online ASC" 

Set rsVerifica001 = Server.CreateObject("ADODB.RecordSet")

	rsVerifica001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsVerifica001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsVerifica001.ActiveConnection = Conexao
	
	
	rsVerifica001.Open SQlVerifica001, Conexao





if rsVerifica001.eof then
Conexao.execute"Insert into cliente_online(nome, email, telefone, data,data_full,atendimento,cod_cliente) values( '"& session("nome") &"','"& session("email") &"','"& session("telefone") &"','"& data001 &"','"& data_full001 &"','"& atendimento001 &"','"& vCod_cliente001 &"')"

end if



 else
 
 end if

'------------------------------------------------------------------------



'----------------------transformar combo1 em cidade-----------------------

if session("vCidade2") <> "cqualquer" and session("vCidade2") <> "" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade2")
 
 
 rs2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs2.ActiveConnection = Conexao
 
 
 
 
 rs2.open SQL2,Conexao,2,1
 
 vCidade = rs2("nome_combo1")
 
 '-----------------------------
           rs2.Close           
		   
           Set rs2 = Nothing
		   
'---------------------------------
 
 else
 vCidade = vCidade2
 end if

session("vCidade") = vCidade




'--------------------------------------------------------------------------------

'-------------------------Transformar combo2 em bairro----------------------------

 if session("vBairro2") <> "bqualquer" and session("vBairro2") <> "" then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 from combo2 where id_combo2 ="& session("vBairro2")
 
 
 rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao
 
 
 
 
 
 rs3.open SQL3,Conexao,2,1

 vBairro = rs3("nome_combo2")
 
 '-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------
 
 
 else
 vBairro = vBairro2
	end if    


session("vBairro") = vBairro


'----------------------------------------------------------------------------------

















'------------------------Montar instrução SQL-----------------------------------------

 '----------------------Cidade--------------------------------
 dim stringIndex
 
stringIndex = " where cod_imovel<>"&"0"&"" 

dim stringCidade

if  session("vCidade") <> "cqualquer" and session("vCidade") <> "" then
stringCidade = "and cidade='"& session("vCidade")&"'"
else
stringCidade = ""
end if
 '--------------------------Bairro----------------------------


dim stringBairro

if session("vBairro") <> "bqualquer" and session("vBairro") <> ""  then
stringBairro = " and Bairro='"&session("vBairro")&"'"
else
stringBairro = "" 
end if



'------------------Tipo-----------------------------

dim StringTipo

if session("vTipo") <> "tqualquer" and session("vTipo") <> "" then
stringTipo = " and Tipo='"&session("vTipo")&"'"
else
stringtipo = ""
end if







'------------------------------------------------------




'----------------------------Vagas na garagem----------------


dim StringVagas

if session("vVagas") <> "gqualquer" and session("vVagas") <> "" then
stringVagas = " and vagas >="&session("vVagas")&""
else
stringVagas = ""
end if

'---------------------------------------------------------------















'-------------------Negociação---------------------------

if  session("vNegociacao") <> "nqualquer" and session("vNegociacao") <> "" then
stringNegociacao = " and negociacao='"&session("vNegociacao")&"'"
else
stringNegociacao = ""

end if

'------------------------------------------------------------------------------


'---------------------------Quartos------------------------------


if  session("vQuartos") <> "qqualquer" and session("vQuartos") <> "" then
stringQuartos = " and quartos >="&int(session("vQuartos"))&""
else
stringQuartos = ""

end if


'-------------------------------------------------------------


'---------------------------------Valor-----------------------------------



dim stringValor



if  session("vValor") <> "vqualquer" and session("vValor") <> ""  then
stringValor = " and Valor >="& session("vValor1") &" and Valor <="& session("vValor2") &""
else
stringValor = ""
end if





  
		  dim SqlTelefone_acesso
		  dim rsTelefone_acesso
		  
		  SqlTelefone_acesso = "SELECT * FROM telefone_acesso where telefone_acesso like '"&session("telefone")&"' ORDER BY cod_telefone_acesso ASC" 

Set rsTelefone_acesso = Server.CreateObject("ADODB.RecordSet")

	rsTelefone_acesso.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsTelefone_acesso.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsTelefone_acesso.ActiveConnection = Conexao
	
	
	rsTelefone_acesso.Open sqlTelefone_acesso, Conexao
		  
		dim permissao_sem_foto  
		  
		  
		  if not rsTelefone_acesso.eof  then
		  
		 permissao_sem_foto = "permitido"
		 
		 else
		 
		  permissao_sem_foto = "não permitido" 
		   
		   end if





if session("vnegociacao") <> "Aluguel"   then

if permissao_sem_foto = "não permitido"   then

strSQL ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM Imoveis"&stringIndex&stringCidade&stringBairro&stringTipo&stringVagas&stringNegociacao&stringQuartos&stringVagas&stringValor&" and (imovel_em_negociacao like  '"&"imóvel OK"&"') and foto_grande1 <> 'imovel00000.jpg' ORDER BY cod_imovel DESC" 
else


strSQL ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM Imoveis"&stringIndex&stringCidade&stringBairro&stringTipo&stringVagas&stringNegociacao&stringQuartos&stringVagas&stringValor&" and (imovel_em_negociacao like  '"&"imóvel OK"&"')  ORDER BY cod_imovel DESC" 


end if	
	
	
	
else

strSQL ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM Imoveis"&stringIndex&stringCidade&stringBairro&stringTipo&stringVagas&stringNegociacao&stringQuartos&stringVagas&stringValor&" and (imovel_em_negociacao like  '"&"imóvel OK"&"')  ORDER BY cod_imovel DESC" 


end if




'-------------------------------------------------------------------------------------


'--------------------------primeira parte da paginação---------------------------------





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
	
RS.MaxRecords = 50	
	
	
RS.Open strSQL, Conn, 1, 3
'o recordset é aberto
	
RS.PageSize = 5
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount





'A variável intRecordCount recebe o valor do número de registros retornados no recordset.

'If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.


'end if

'--------------------------------------------------------------------------------------

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Listar imóveis</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=520,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>



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

if (doublecombo.txt_email.value == "Seu email:") {
		alert("Por favor, coloque seu email , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "") {
		alert("Por favor, coloque seu email , pois assim , você terá um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
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
















}
}


</script>


</head>

<body>
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"  method="post" action="listar_imoveis01.asp">
<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="106"><img src="top01.jpg" width="794" height="106"></td>
  </tr>
  <tr>
    <td height="237"><table width="794" height="257" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="552" height="257"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="552" height="257">
                <param name="movie" value="frontpage001.swf">
                <param name="quality" value="high">
                <embed src="frontpage001.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="552" height="257"></embed></object>
              </img></td>
          <td width="242" bgcolor="#e0a94e"><div align="center">
              <table width="232" height="247" border="0" cellpadding="0" cellspacing="0" bgcolor="#e6dca9">
                <tr>
                  <td bgcolor="#e6dca9">
<div align="center">
                      <table width="222" height="237" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e0a94e"><div align="center">
                              <table width="212" height="227" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td bgcolor="#e6dca9"><div align="center"> 
                                      <table width="202" height="217" border="0" cellpadding="0" cellspacing="0">
                                        <tr> 
                                          <td><div align="center"> 
                                              <table width="202" border="0" cellspacing="0" cellpadding="0">
                                                <tr> 
                                                  <td height="20"><input name="txt_nome" onFocus="doublecombo.txt_nome.value=''"  type="text" class="inputBox" id="txt_nome"  style="HEIGHT: 18px; WIDTH: 202px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" value="<%=session("nome")%>"></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><input name="txt_telefone" onFocus="doublecombo.txt_telefone.value=''"  type="text" class="inputBox" id="txt_nome2"  style="HEIGHT: 18px; WIDTH: 202px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" value="<%=session("telefone")%>"></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><input name="txt_email" onFocus="doublecombo.txt_email.value=''"  type="text" class="inputBox" id="txt_nome3"  style="HEIGHT: 18px; WIDTH: 202px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" value="<%=session("email")%>"></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="combo1" id="combo1" onChange="javascript:atualizacarros(this.form);" size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                     
													  <option value="cqualquer" selected>Cidade</option>
				  <% if not rs303.eof then %>
                  <% While NOT Rs303.EoF %>
                  
				  <option value="<% = Rs303("id_combo1") %>"<%if session("vCidade2")<> "cqualquer" then%><%if int(rs303("id_combo1")) = int(session("vCidade2")) then response.write "selected" else response.write "" end if %><%end if%>> 
				  
				 <% = Rs303("nome_combo1") %> </option>
                  
                  
                  <% Rs303.MoveNext %>
                  <% Wend %>
				  <option value="cqualquer">qualquer uma</option>
                  <%else%>
                  <option value=""></option>
                  <%end if%>
												   
												    </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="combo2" id="combo2"   size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                      <option value="bqualquer" selected>Bairro/Região</option>
				  <% if not rs404.eof then%>
                  <% While NOT Rs404.EoF %>
                  
				    <option value="<% = Rs404("id_combo2") %>" <% if session("vBairro2") <> "bqualquer" then if Rs404("id_combo2") = int(session("vBairro2"))  then response.write "selected" else response.write "" end if end if %>> 
                  <% = Rs404("nome_combo2") %>
                  </option>
				  
				  
				  
                  <% Rs404.MoveNext %>
				  
                  <% Wend %>
				   <option value="bqualquer">qualquer um</option>
				  
                  <% else %>
                  <option value=""></option>
                  <% end if %>
                                                    </select></td>
                                                </tr>
												
												 <tr> 
                                                  <td height="20"><select name="txt_tipo" id="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                         <option value="<%=session("vTipo")%>" selected><%if session("vTipo") <> "tqualquer" and session("vTipo") <> "" then  response.write session("vTipo") else response.write "Tipo" end if%></option>
                                                        <option value="tqualquer">Qualquer 
                                                        um</option>
                                                        <% if not rs444Tipo22.eof then%>
                                                        <% While NOT (rs444Tipo22.EoF) %>
                                                        <option value="<% = rs444Tipo22("tipo") %>"> 
                                                        <% =rs444Tipo22("tipo") %>
                                                        </option>
                                                        <% rs444Tipo22.MoveNext %>
                                                        <% Wend %>
                                                        <% else %>
                                                        <option value=""></option>
                                                        <% end if %>
                                                      </select></td>
                                                </tr>
												
												
												
												
                                                <tr> 
                                                  <td height="20"><select name="txt_quartos" id="txt_quartos" size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                       <option value="<%=session("vQuartos")%>"><% if session("vQuartos") <> "qqualquer" and session("vQuartos") <> "" then response.write session("vQuartos") else response.write "Quartos" end if%></option>
                                                        <option value="qqualquer">Qualquer 
                                                        um</option>
                                                        <option value="01">01</option>
                                                        <option value="02">02</option>
                                                        <option value="03">03</option>
                                                        <option value="04">04</option>
                                                        <option value="05">05</option>
                                                        <option value="06">06</option>
                                                        <option value="07">07</option>
                                                        <option value="08">08</option>
                                                        <option value="09">09</option>
                                                      </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="txt_garagem" id="txt_garagem" size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                         <option value="<%=session("vVagas")%>"><% if session("vVagas") <> "gqualquer" and session("vVagas") <> "" then response.write session("vVagas") else response.write "Vagas" end if%></option>
                                                       
													    <option value="gqualquer">Qualquer um</option>
                                                        <option value="01">01</option>
                                                        <option value="02">02</option>
                                                        <option value="03">03</option>
                                                        <option value="04">04</option>
                                                        <option value="05">05</option>
                                                        <option value="06">06</option>
                                                        <option value="07">07</option>
                                                        <option value="08">08</option>
                                                        <option value="09">09</option>
                                                      </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="example2" id="example2" onChange="redirect2(this.options.selectedIndex)" size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                        <option value="nqualquer" selected>Negociação 
                                                        </option>
                                                        <option value="nqualquer" >Qualquer 
                                                        um </option>
                                                        <option  value="Aluguel">Aluguel 
                                                        </option>
                                                        <option value="Venda">Venda 
                                                        </option>
														 <option value="<%=session("vnegociacao")%>" selected><% if session("vnegociacao") <> "nqualquer" and session("vnegociacao") <> "" then response.write session("vnegociacao") else response.write "Negociação" end if%></option>
                                                      </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="stage22" id="stage22" size="1"  class="inputBox" style="HEIGHT: 18px; font-size : 12px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;">
                                                        <option value="vqualquer" >Valor</option>
                                                        <option value="vqualquer">Qualquer um</option>
														 <% if session("vnegociacao") = "Aluguel" then %>
			 <option value="<%=session("vValor")%>" selected><% if session("vValor") <> "vqualquer" and session("vValor") <> "" then response.write FormatNumber(session("vValor1"),2)&" até "&FormatNumber(session("vValor2"),2) else response.write "Valor" end if%></option>
			<option value="0000000000 0000000200">Até 200,00</option>
                <option value="0000000000 0000000200">Até 200,00</option>
                  <option value="0000000201 0000000500">201,00 até 500,00</option>
                  <option value="0000000501 0000000750">501,00 até 750,00</option>
                  <option value="0000000751 0000001000">751,00 até 1000,00</option>
                  <option value="0000001001 0000001500">1001,00 até 1500,00</option>
                  <option value="0000001501 0000002000">1501,00 até 2000,00</option>
                  <option value="0000002001 0000002500">2001,00 até 2500,00</option>
                  <option value="0000002501 0000003000">2501,00 até 3000,00</option>
                  <option value="0000003001 0000003500">3001,00 até 3500,00</option>
                  <option value="0000003501 0000004000">3501,00 até 4000,00</option>
                  <option value="0000004001 1000000000">Acima de 4000,00</option>
               <%else%>
			   <option value="<%=session("vValor")%>" selected><% if session("vValor") <> "vqualquer" and session("vValor") <> "" then response.write FormatNumber(session("vValor1"),2)&" até "&FormatNumber(session("vValor2"),2) else response.write "Valor" end if%></option>
			   <option value="vqualquer" >Valor</option>
                                                        <option value="vqualquer">Qualquer 
                                                        um</option>
                                                        <option value="0000000000 0000020000">Até 
                                                        20.000,00</option>
                                                        <option value="0000020001 0000050000">20.001,00 
                                                        até 50.000,00</option>
                                                        <option value="0000050001 0000080000">50.001,00 
                                                        até 80.000,00</option>
                                                        <option value="0000080001 0000110000">80.001,00 
                                                        até 110.000,00</option>
                                                        <option value="0000110001 0000150000">110.001,00 
                                                        até 150.000,00</option>
                                                        <option value="0000150001 0000200000">150.001,00 
                                                        até 200.000,00</option>
                                                        <option value="0000200001 0000250000">200.001,00 
                                                        até 250.000,00</option>
                                                        <option value="0000250001 0000300000">250.001,00 
                                                        até 300.000,00</option>
                                                        <option value="0000300001 0000350000">300.001,00 
                                                        até 350.000,00</option>
                                                        <option value="0000350001 0000400000">350.001,00 
                                                        até 400.000,00</option>
														
														<option value="0000400001 0000600000">400.001,00 
                                                        até 600.000,00</option>
														<option value="0000600001 0000800000">600.001,00 
                                                        até 800.000,00</option>
														
														<option value="0000800001 0001000000">800.001,00 
                                                        até 1000.000,00</option>
														
                                                        <option value="0001000001 1000000000">Acima 
                                                        de 1000.000,00</option>
			   <%end if%>
                                                      </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><input name="image" type="image" src="bt_procurar303.jpg" width="201" height="18"></td>
                                                </tr>
                                              </table>
                                            </div></td>
                                        </tr>
                                      </table>
                                    </div></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="16"><img src="subtop01.jpg" width="794" height="16"></td>
  </tr>
  <tr>
      <td height="40">
	  <% if intRecordCount <> 0 then %>
	    <div align="center"><font color="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foram 
          encontrados <%=intRecordCount%> im&oacute;veis com fotos, mas existe 
          mais op&ccedil;&otilde;es sem fotos, ligue e fale com um de nossos atendentes</strong></font></div></td>
   <%end if%>
    </tr>
 
 <%
 
 
 

dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where (telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"') and atendimento <>'"&"internet"&"' and atendimento <>'"&"não informado"&"' " 
	
	
	
	rs444VerificaConta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta.ActiveConnection = Conexao
	
	
	
	
	
	 rs444VerificaConta.Open strSQL444VerificaConta, Conexao
	

if  not rs444VerificaConta.eof then






vCadastrado = "sim"

%>
 
 
  
  <tr>
  <td>
  
  <table width="794" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="200" bgcolor="#e0a94e"> 
        <div align="center">
          <table width="785" height="190" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td bgcolor="#e6dca9"><table width="600" height="150" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
    <td width="600"><div align="center"><font color="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="5"><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:red;text-decoration:none;"  target="_blank">Ol&aacute; 
                              sr(a) <%=session("nome")%></a></font><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:red;text-decoration:none;" target="_blank"><br>
                              obrigado por retornar ao nosso site, o senhor est&aacute; 
                              conosco desde <%=rs444VerificaConta("data")%>, o 
                              seu atendente &eacute; o sr(a)<%=rs444VerificaConta("atendimento")%>, 
                              </a></strong></font><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acessoLink03.asp?varTelefone=<%=session("telefone")%>" style="color:red;text-decoration:none;" target="_blank">para 
                              visitar um im&oacute;vel ligue para seu atendente 
                              ou agende a visita na ficha do im&oacute;vel escolhido.</a></strong></font></div></td>
  </tr>
</table>
 </td>
            </tr>
          </table>
		  </div>
</table>
<tr>
      <td height="20">
<div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:#e0a94e;text-decoration:none;" target="_blank"> 
          </a></strong></font></div></td>
    </tr>
</td>
</tr>
  
  <%else%>
  
  
  <%end if%>
  
  
  <tr>
  <td>
  
  <%


If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.



	
If CInt(intPage) > CInt(intPageCount) Then 
intPage = intPageCount
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.
end if



	If CInt(intPage) <= 0 Then
	 intPage = 1
	end if
	'se intPage é menor ou igual a zero então intPage igual a "1"
	'a variável intPage sempre vai ser forçada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados então.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a página exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posição exata do primeiro registro da página correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual ao número de páginas no recordset , estamos na última 
			'página então.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor do número do último recordset.
				'intFinish corresponde ao valor do último registro da página correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'do número de registros na página menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros então
		For intRecord = 1 to RS.PageSize
		'um contador inRecord é colocado até o número de registros na página.
%>

<%

varCodImovel = rs("cod_imovel")

%>

<table width="795" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="795" height="200" style="border:1px solid #ddddc5;"><table width="784" height="190" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td bgcolor="#e9dca8"><table width="774" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td><table width="774" height="180" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="299" height="180" style="border:1px solid #FFFFFF;" bgcolor="#f7ecbf"> 
                                <% If objFSO.FileExists(Server.MapPath(rs("foto_pequena"))) = True Then%>
                                <a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')"><img src="<%=rs("foto_pequena")%>" width="299" height="178" border="0"></img></a> 
                                <%else%>
                                <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><strong>Foto não disponível</strong></a></font></div><%end if%></td>
                              <td width="5" height="180">&nbsp;</td>
                              <td width="469" height="180" bgcolor="#f7ecbf"><table width="469" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="469" height="18"><table width="469" border="0" cellspacing="0" cellpadding="0">
                                        <tr> 
                                          <td width="117" height="18" style="border:1px solid #e9dca8;" >
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("cidade")%></font></div> 
                                            </td>
                                          <td width="117" height="18" style="border:1px solid #e9dca8;">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("bairro")%></font></div> 
                                            </td>
                                          <td width="117" height="18" style="border:1px solid #e9dca8;">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%if rs("tipo") <> "tqualquer" then response.write rs("tipo") else response.write "não informado" end if%></font></div> 
                                            </td>
                                          <td width="117" height="18" style="border:1px solid #e9dca8;">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=formatnumber(rs("valor"),2)%></font></div> 
                                            </td>
                                        </tr>
                                      </table></td>
                                  </tr>
                                  <tr>
                                    <td width="469" height="164">
									
									<% if rs("qualidade") = "bom negócio" and rs("imovel_em_negociacao") <> "Vendido pela Veja" then%>
									
									<table width="469" height="164" border="0" cellpadding="0" cellspacing="0">
        <tr>
                                          <td width="379">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><strong><%=rs("obs_imovel")%></strong></a><br>
                                              <br>
                                              <strong>Atualizado em:</strong><strong><%=rs("data_atualizacao")%></strong><br>
                                              <br>
                                              <strong>Referência:<%=rs("cod_imovel")%></strong></font></div>
											  <br>
											  <%
					 
					
					 
					  SqlCompradores001 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores,compradores.valor FROM compradores where telefone like '"&rs("telefone")&"' or telefone02 like '"&rs("telefone")&"' or telefone03 like '"&rs("telefone")&"' ORDER BY cod_compradores Desc" 

Set rsCompradores001 = Server.CreateObject("ADODB.RecordSet")

	rsCompradores001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCompradores001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCompradores001.ActiveConnection = Conexao
	
	
	rsCompradores001.Open sqlCompradores001, Conexao
					  
		if not rsCompradores001.eof then 
		
		
		
		
		
		
					  %>
					  <br>
					  
					  
					  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Este 
                                              propriet&aacute;rio est&aacute; 
                                              vendendo este im&oacute;vel para 
                                              comprar outro no valor de <%=formatnumber(rsCompradores001("valor"),2)%>, 
                                              existe a possibilidade de <font size="4">permuta</font> 
                                              ,clique no texto acima para saber 
                                              mais</font> </strong></font></div>
					  <%
					  end if
					  
					  rsCompradores001.close

                     set rsCompradores001 = nothing
					  
					  
					  %>
											  <br>
											  
											  
											  
											  </td>
          <td width="90"><table width="90" height="164" border="0" cellpadding="0" cellspacing="0">
              <tr>
                                                <td height="90" ><table width="80" height="80" border="0" align="center" cellpadding="0" cellspacing="0">
                                                    <tr>
                                                      <td bgcolor="#e9dca8" style="border:1px solid #d2c48c;"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="text-decoration:none;color:#9d9249"><img src="bt_top01.jpg" border="0"></img></a></font></div></td>
                                                    </tr>
                                                  </table></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table>
	  
	  
	  
	  <% elseif rs("imovel_em_negociacao") = "Vendido pela Veja" then %>
									
		
		<table width="469" height="164" border="0" cellpadding="0" cellspacing="0">
        <tr>
                                          <td width="379">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><strong><%=rs("obs_imovel")%></strong></a><br>
                                              <br>
                                              <strong>Atualizado em:</strong><strong><%=rs("data_atualizacao")%></strong><br>
                                              <br>
                                              <strong>Referência:<%=rs("cod_imovel")%></strong></font></div>
											  
											  <br>
											  
											   <%
					 
					
					 
					  SqlCompradores001 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores,compradores.valor  FROM compradores where telefone like '"&rs("telefone")&"' or telefone02 like '"&rs("telefone")&"' or telefone03 like '"&rs("telefone")&"' ORDER BY cod_compradores Desc" 

Set rsCompradores001 = Server.CreateObject("ADODB.RecordSet")

	rsCompradores001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCompradores001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCompradores001.ActiveConnection = Conexao
	
	
	rsCompradores001.Open sqlCompradores001, Conexao
					  
		if not rsCompradores001.eof then 
		
		
		
		
		
		
					  %>
					  <br>
					  
					  
					  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Este 
                                              propriet&aacute;rio est&aacute; 
                                              vendendo este im&oacute;vel para 
                                              comprar outro no valor de <%=formatnumber(rsCompradores001("valor"),2)%>, 
                                              existe a possibilidade de <font size="4">permuta</font> 
                                              ,clique no texto acima para saber 
                                              mais</font> <br>
                                              </strong></font></div>
					  <%
					  end if
					  
					  rsCompradores001.close

                     set rsCompradores001 = nothing
					  
					  
					  %>
											  <br>
											  
											  
											  
											
											  
											  
											  
											  </td>
          <td width="90"><table width="90" height="164" border="0" cellpadding="0" cellspacing="0">
              <tr>
                                                <td height="90" ><table width="80" height="80" border="0" align="center" cellpadding="0" cellspacing="0">
                                                    <tr>
                                                      <td bgcolor="#ecdf94" style="border:1px solid #d7c971;"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><img src="bt_sorriso02.jpg" border="0"></img></a></font></div></td>
                                                    </tr>
                                                  </table></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table>
	  
	  <%else%>
	  
	                                  <div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><strong><%=rs("obs_imovel")%></strong></a><br>
                                        <br>
                                        <strong>Atualizado em:</strong><strong><%=rs("data_atualizacao")%></strong><br>
                                        <br>
                                        <strong>Referência:<%=rs("cod_imovel")%></strong></font></div>
		
		
		<%
					 
					
					 
					  SqlCompradores001 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores,compradores.valor FROM compradores where telefone like '"&rs("telefone")&"' or telefone02 like '"&rs("telefone")&"' or telefone03 like '"&rs("telefone")&"' ORDER BY cod_compradores Desc" 

Set rsCompradores001 = Server.CreateObject("ADODB.RecordSet")

	rsCompradores001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCompradores001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCompradores001.ActiveConnection = Conexao
	
	
	rsCompradores001.Open sqlCompradores001, Conexao
					  
		if not rsCompradores001.eof then 
		
		
		
		
		
		
					  %>
		
		 <br>
					  
					  
					  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Este 
                                        propriet&aacute;rio est&aacute; vendendo 
                                        este im&oacute;vel para comprar outro 
                                        no valor de <%=formatnumber(rsCompradores001("valor"),2)%>, 
                                        existe a possibilidade de <font size="4">permuta</font> 
                                        ,clique no texto acima para saber mais</font> 
                                        </strong></font></div>
					  <%
					  end if
					  
					  rsCompradores001.close

                     set rsCompradores001 = nothing
					  
					  
					  %>
											  <br>
									
			<% end if %>						
									
									
									</td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
  </tr>
</table>



<br>




<%
RS.MoveNext


	  





 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
  
  
  </td></tr>
  
  
  
</table>
</form>






<table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vValor=<%=session("vValor")%>&vValor1=<%=session("vValor1")%>&vValor2=<%=session("vValor2")%>&vTipo=<%=session("vTipo")%>&vNegociacao=<%=session("vNegociacao")%>&vOcupacao=<%=session("vOcupacao")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="text-decoration:none;"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
            <strong> 
            <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
            <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
            Página <%=cInt(intPage)%> de <%=cInt(intPageCount)%> </strong></font> 
            <strong><font color="#000000"> 
            <%End If%></font>
            </font></strong></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%>
            <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
            <strong><a href="?page=<%=intPage + 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vValor=<%=session("vValor")%>&vValor1=<%=session("vValor1")%>&vValor2=<%=session("vValor2")%>&vTipo=<%=session("vTipo")%>&vNegociacao=<%=session("vNegociacao")%>&vOcupacao=<%=session("vOcupacao")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="color:#000000;text-decoration:none;">Próximo</a></strong> 
            </a> 
            <%End If%>
            </font></div></td>
        </tr>
      </table>



<%end if%>
<%else%>

<table width="794" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="300" bgcolor="#e0a94e"> 
        <div align="center">
          <table width="785" height="290" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td bgcolor="#e6dca9"><div align="center"> 
                <table width="400" height="200" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><p align="center"><font color="#e0a94e" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#9d9249">Nenhum 
                        im&oacute;vel foi encontrado,</font></strong></font><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:#9d9249;text-decoration:none;" target="_blank">existem 
                        mais op&ccedil;&otilde;es sem foto</a> ,tente novamente.</strong></font></p>
                      <p align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></p> 
                      <div align="center"></div></td>
                  </tr>
                </table>
                <p>&nbsp;</p>
                </div></td>
            </tr>
          </table>
		  </div>
</table>

<%end if%>
<br>
<br>


<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui é criada uma variável "groups" que receberá os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a variável group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero até o número de elementos do array "groups" */

group2[i2]=new Array()
/* aqui é criado o array "group" que receberá valores conforme o número de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receberá valores de opções. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receberá valores de opções. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("201,00 até 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 até 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 até 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 até 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 até 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 até 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 até 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 até 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 até 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Até  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 até 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 até 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 até 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 até 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 até 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 até 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 até 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 até 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 até 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("400.001,00 até 600.000,00","0000400001 0000600000")
group2[3][13]=new Option("600.001,00 até 800.000,00","0000600001 0000800000")
group2[3][14]=new Option("800.001,00 até 1000.000,00","0000800001 0001000000")
group2[3][15]=new Option("Acima de 1000.000,00","0001000001 1000000000")









/* aqui temos um array bidimensional "group" que receberá valores de opções. */


var temp2=document.doublecombo.stage22
/* aqui a variável "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui é criada a função "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que dá um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que é escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" será o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a variável "location" recebe os valores de "stage2" que corresponde ao endereço de
link para o carregamento de página. */


//-->
</script>

<%

'-----------------------Primeira verificação----------------------------------









'-----------------------------------------------------------------------------






'------------------segunda verificação---------------------------
dim rsVerificacao008

dim strSQLVerificacao008

 Set rsVerificacao008 = Server.CreateObject("ADODB.RecordSet")
    
	strSQLVerificacao008 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores Where (telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%'or telefone03  like '%"&session("telefone")&"%')"
	 
   
   
rsVerificacao008.CursorLocation = 3
rsVerificacao008.CursorType = 3

        rsVerificacao008.Open strSQLVerificacao008, Conexao 




	
	
	
	
	

if    rsVerificacao008.eof  and vTipo <> "" then
 
 dim vValorConta
 dim vQuartosConta
 dim vVagasConta
 dim vValorMedio
 
 if session("vValor") = "vqualquer" then
 vValorMedio = "0"
 else
 vValorMedio = (int(session("vValor1")) + int(session("vValor2")))/2
 end if
 
  
  
  
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
 
 
 
 
 dim vAtendimento23
 



'Cadastrar em compradores

dim rs444VerificaConta23,strSQL444VerificaConta23
   
    Set rs444VerificaConta23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta23 = "SELECT imoveis.captacao,imoveis.cod_imovel,imoveis.telefone,imoveis.telefone02,imoveis.telefone03  FROM imoveis  where telefone like'%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	rs444VerificaConta23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta23.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta23.ActiveConnection = Conexao3
	
	
	
	
	
	
	 rs444VerificaConta23.Open strSQL444VerificaConta23, Conexao3
	

if  not rs444VerificaConta23.eof  then

vAtendimento23 = rs444VerificaConta23("captacao")

else

vAtendimento23 = "internet"

'Conexao3.execute"Insert into permuta(Foto_imovel,Nome,Email,Telefone,endereco_vend,cidade_vend,bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby)values( '"& "imovel00000.jpg" &"','"& session("nome") &"','"& session("email") &"','"& session("telefone") &"','"& "não informado" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "não informado" &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& "não informado" &"','"& "0" &"','"& "não informado" &"','"& now() &"','"& vQuartosConta_vend &"','"& vQuartosConta_comp &"','"& vValorMedio_vend &"','"& vValorMedio_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"&vVila_comp&"','"& vVagasConta_vend &"','"& vVagasConta_comp &"','"& "excluido" &"')" 
		

 end if

 
 
 '------------------Verifica se o internauta já tem conta---------------------------
  
  dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"'" 
	
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao




   if vTipo <> "" and (not rs444VerificaConta2.eof) then              
		if rs444VerificaConta2("acessos") <> "" then
		
		 
	 Conexao.execute"update compradores set acessos='"&int(rs444VerificaConta2("acessos"))+1&"' where cod_compradores="&rs444VerificaConta2("cod_compradores")
	 
			else
			
			 	 
	 Conexao.execute"update compradores set acessos='"&"1"&"' where cod_compradores="&rs444VerificaConta2("cod_compradores")
	 
		end if
		
		
		
  end if

 
 
 
 
 
 
 if rs444VerificaConta2.eof then
  
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem,area_total,area_construida,condominio,condicoes_pagamento,data_contato,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vQuartosConta") &"','"& session("vNegociacaoConta") &"','"& int(session("vValorMedio")) &"','"& now() &"','"& "não informado" &"','"& vAtendimento23 &"','"& now() &"','"& session("vVila") &"','"& session("vVagasConta") &"','"& "não informado" &"','"& "comprador a contatar" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "Busca de imóveis" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "não informado" &"','"& now() &"','"& session("vOrigem_Franquia") &"')"
	
	end if
	
	
	
	end if
	
	
  
  
  
  
  '--------adicionar acessos----------------------------------------------
  
  
  
  
  
  
  
  '----------------Adicionar em buscas efetuadas----------------
  
  
  if vTipo <> "" then 
  
  dim EnderecoIP02

EnderecoIP02 = request.ServerVariables("REMOTE_ADDR")
  
	'Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem,area_total,area_construida,condominio,condicoes_pagamento) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vQuartosConta") &"','"& session("vNegociacaoConta") &"','"& int(session("vValorMedio")) &"','"& now() &"','"& "não informado" &"','"& vAtendimento23 &"','"& now() &"','"& session("vVila") &"','"& session("vVagasConta") &"','"& "não informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "não informado" &"')"
	
	
  
	Conexao.execute"Insert into imoveis_procurados(cidade,bairro,tipo,negociacao,valor,data,enderecoIP,quartos,vagas,nome,telefone,email,atendimento,origem_franquia) values( '"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vNegociacao") &"','"& session("vValor") &"','"& now() &"','"& EnderecoIP02 &"','"& session("vQuartos") &"','"& session("vVagas") &"','"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"&Atendimento001&"','"& session("vOrigem_Franquia") &"')"
	
  
  
  end if
  
  
  
  
  
  
  
  
  '----------------------------adicionar acesso--------------------------------



if session("telefone") <> "" and session("acessos") = "" then



dim rs444VerificaConta022,strSQL444VerificaConta022
   
    Set rs444VerificaConta022 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta022 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'"&session("telefone")&"' or telefone02 like'"&session("telefone")&"' or telefone03 like'"&session("telefone")&"'" 
	 
	 
	 
	

	rs444VerificaConta022.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta022.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta022.ActiveConnection = Conexao
	
	
	
	 
	 
	 
	 
	 
	 rs444VerificaConta022.Open strSQL444VerificaConta022, Conexao
	
if not rs444VerificaConta022.eof then

  dim vNumero_acessos
  
  vNumero_acessos = rs444VerificaConta022("acessos")
  end if
  
  if vNumero_acessos = "" then
  vNumero_acessos = "1"
  end if
  
  
  vNumero_acessos = int(vNumero_acessos) + 1


if  not rs444VerificaConta022.eof then




	 Conexao.execute"update compradores set data_ultimo_acesso='"&now()&"',acessos='"&vNumero_acessos&"'  where cod_compradores="&rs444VerificaConta022("cod_compradores")
	end if 
      
	  
	  session("acessos") = "acessado"
	  
	  

end if
  
  
  
  
  
  '---------------------------------------------------------------------
  
  %>

<%

	
	
dim rs444VerificaConta02,strSQL444VerificaConta02
   
    Set rs444VerificaConta02 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta02 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"'" 
	 
	 
	 
	

	rs444VerificaConta02.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta02.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta02.ActiveConnection = Conexao
	
	
	
	 
	 
	 
	 
	 
	 rs444VerificaConta02.Open strSQL444VerificaConta02, Conexao
	

if  not rs444VerificaConta02.eof then




	 Conexao.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta02("cod_compradores")
	
	
	end if 
	









%>





 <% response.flush%>
  <%response.clear%>

<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizão

rsMarcas3.ActiveConnection = Conexao3

rsMarcas3.Open SqlMarcas3, Conexao3


While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizão

rsCarros3.ActiveConnection = Conexao3

rsCarros3.Open SqlCarros3, Conexao3


'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1
While NOT rsCarros3.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



  rsMarcas3.Close           
		   
           Set rsMarcas3 = Nothing
		   
		   
		   rsCarros3.Close           
		   
           Set rsCarros3 = Nothing






End Function


        




%> 

<%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%'permissao_sem_foto%>
<%'strSQL%>
</body>
</html>
