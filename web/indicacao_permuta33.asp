<!--#include file="dsn.asp"-->
<!--#include file="style_conta.asp"-->
<!--#include file="cores02.asp"-->






<%

'Criando conex�o com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn



%>



<%

dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	
	
	
	
	dim strSQL
	dim rs
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where cod_permuta="&varCodPermuta
	 rs.CursorLocation = 3
      rs.CursorType = 3
	 rs.Open strSQL, Conexao





%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela2) {
   openWindow2 = window.open(abrejanela2,'openWin2','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow2.focus( )
   }

</SCRIPT>

</head>

<body bgcolor="#FFFFFF">
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_conta_permuta01.asp?varCodPermuta=<%=varCodPermuta%>">
</form>


<%







dim negrito,negrito2
dim vValor_vend,vValor_vend1,vValor_vend2
dim vValor_comp,vValor_comp1,vValor_comp2
dim vCidade_vend,vCidade_comp


dim varIndicacaoCidadeVend
dim varIndicacaoBairroVend
dim varIndicacaoVilaVend
dim varIndicacaoQuartosVend
dim varIndicacaoVagasVend
dim varIndicacaoValorVend
dim varIndicacaoTipoVend


dim varIndicacaoCidadeComp
dim varIndicacaoBairroComp
dim varIndicacaoVilaComp
dim varIndicacaoQuartosComp
dim varIndicacaoVagasComp
dim varIndicacaoValorComp
dim varIndicacaoTipoComp

dim varIndicacaoCodigo


 varIndicacaoCidadeVend = rs("cidade_vend")
 varIndicacaoBairroVend = rs("bairro_vend")
 varIndicacaoVilaVend = rs("vila_vend")
 varIndicacaoQuartosVend = rs("quartos_vend")
 varIndicacaoVagasVend = rs("vagas_vend")
 varIndicacaoValorVend = rs("valor_vend")
 varIndicacaoTipoVend = rs("tipo_vend")
 
 
 
 session("varIndicacaoCidadeVend") = varIndicacaoCidadeVend
 session("varIndicacaoBairroVend") = varIndicacaoBairroVend
 session("varIndicacaoVilaVend") = varIndicacaoVilaVend
 session("varIndicacaoQuartosVend") = varIndicacaoQuartosVend
 session("varIndicacaoVagasVend") = varIndicacaoVagasVend
 session("varIndicacaoValorVend") = varIndicacaoValorVend
 session("varIndicacaoTipoVend") = varIndicacaoTipoVend
 
 
 
 
 
 
 varIndicacaoCidadeComp = rs("cidade_comp")
 varIndicacaoBairroComp = rs("bairro_comp")
 varIndicacaoVilaComp = rs("vila_comp")
 varIndicacaoQuartosComp = rs("quartos_comp")
 varIndicacaoVagasComp = rs("vagas_comp")
 varIndicacaoValorComp = rs("valor_comp")
 varIndicacaoTipoComp = rs("tipo_comp")
 
 
 session("varIndicacaoCidadeComp") = varIndicacaoCidadeComp
 session("varIndicacaoBairroComp") = varIndicacaoBairroComp
 session("varIndicacaoVilaComp") = varIndicacaoVilaComp
 session("varIndicacaoQuartosComp") = varIndicacaoQuartosComp
 session("varIndicacaoVagasComp") = varIndicacaoVagasComp
 session("varIndicacaoValorComp") = varIndicacaoValorComp
 session("varIndicacaoTipoComp") = varIndicacaoTipoComp
 
 
 
 varIndicacaoCodigo=request.querystring("varIndicacaoCodigo")
 
session("varIndicacaoCodigo") = varIndicacaoCodigo


 
 
 
 
 
 
 
 
 
 
  '---------Selecionar permutante pelo telefone------------------------------------------------
		   
		     dim rs202,SQL444Permuta202
 Set rs202 = Server.CreateObject("ADODB.RecordSet")
 SQL444Permuta202 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone='"& rs("telefone")&"' order by cod_permuta DESC" 
	
	
	rs202.CursorLocation = 3
         rs202.CursorType = 3
           rs202.ActiveConnection = Conexao
	
	
	rs202.open SQL444Permuta202,Conexao,2,1  
	
			
	if  not rs202.eof then
		   
		   
		   
		   
		   
		   
'------------------------Sua Cidade--------------------------

stringIndex202 = " where cod_permuta<>"&"0"&""
 
 
 
  if   rs202("cidade_vend") = "n�o informado" or rs202("cidade_vend") = "" or rs202("cidade_vend") = "cqualquer" or  rs202("cidade_vend") = "qualquer um" then
	stringCidadeVend202 = ""
 else

stringCidadeVend202 = " and (Cidade_comp='"&rs202("cidade_vend")&"' or Cidade_comp='"&"n�o informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend202

 if   rs202("bairro_vend") = "n�o informado" or rs202("bairro_vend") = "" or rs202("bairro_vend") = "bqualquer" or  rs202("bairro_vend") = "qualquer um" then
	stringBairroVend202 = ""
 else
'stringBairroVend = ""
stringBairroVend202 = " and (Bairro_comp like'%"&rs202("bairro_vend")&"%' or Bairro_comp like'%"&"n�o informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend202

'" and (Vila_comp='"&rs("vila_vend")&"' or Vila_comp='"&"n�o informado"&"' or Vila_comp='"&"vlqualquer"&"' or Vila_comp='"&"qualquer um"&"' )"

 if   rs202("vila_vend") = "n�o informado" or rs202("vila_vend") = "" or rs202("vila_vend") = "vlqualquer" or rs202("vila_vend") = "qualquer um" then
	stringVilaVend202 =  ""
 else

stringVilaVend202 = ""

end if






 '--------------------------Tipo do seu im�vel------------------------
 
 
 dim stringTipoVend202
 
 
 if rs202("tipo_vend") = "n�o informado" or rs202("tipo_vend") = "" or rs202("tipo_vend") = "tqualquer" or rs202("tipo_vend") = "qualquer um"  then

stringTipoVend202 = ""

else
stringTipoVend202 = " and Tipo_comp like '%"&rs202("tipo_vend")&"%'"
 
 end if


 
 '-----------------------N�mero de quartos do seu im�vel-----------------
 
 
 
 
 dim stringQuartosVend202
 
 
 

stringQuartosVend202 = " and Quartos_comp <="&int(rs202("quartos_vend"))&""

 


 '-----------------------N�mero de Vagas do seu im�vel-----------------
 
 
 
 
 dim stringVagasVend202
 
 
 



stringVagasVend202 = " and vagas_comp <="&int(rs202("vagas_vend"))&""

 




 
 
 
 
 '-----------------------------Valor de venda do seu im�vel----------------
 
 
 
dim PorcentualVend202

dim vValorMenorVend202
dim vValorMaiorVend202

PorcentualVend202 = int(rs202("valor_vend"))*20/100

   


   vValorMenorVend202 = int(rs202("valor_vend")) - int(PorcentualVend202)
   vValorMaiorVend202 = int(rs202("valor_vend")) + int(PorcentualVend202)

 
 
 
 
 
	 dim stringValorVend202
  
	
	
	
	stringValorVend202 = " and Valor_comp >="&  vValorMenorVend202 &""
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp202
  if rs202("cidade_comp")="n�o informado" or rs202("cidade_comp")="" or rs202("cidade_comp")="cqualquer" or rs202("cidade_comp") = "qualquer um" then
	stringCidadeComp202 = ""
	else
	
	stringCidadeComp202 = " and Cidade_vend ='"& rs202("cidade_comp") &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp202

	if rs202("bairro_comp") = "n�o informado" or  rs202("bairro_comp") = "" or  rs202("bairro_comp") = "bqualquer" or rs202("bairro_comp") = "qualquer um" then
	
	
	
	
	
	stringBairroComp202 = ""
	
	
	
	
	else
	
	
	
	'stringBairroComp = " and Bairro_vend ='"& rs("bairro_comp") &"'"
	
	
	
	
 
dim Numero_Indicacoes202
dim Numero_Indicacoes02202




Numero_Indicacoes202 = 0
Numero_Indicacoes02202 = 0


dim soma02202
dim soma202

soma202 = 0
soma02202 = 0

dim Variavel202
dim Retorno202
dim contar202
Variavel202 = rs202("bairro_comp")
Retorno202 = Split(rs202("bairro_comp"),", ")

contar202=0

dim stringBairro3202
dim stringBairro4202
dim stringBairro5202

for contar202=0 to UBound(Retorno202)

stringBairro3202 = "and ( "
stringBairro4202 = " Bairro_vend='"&Retorno202(contar202)&"'or  " &stringBairro4202

stringBairro5202 = " cod_permuta=0)"


stringBairroComp202 = stringBairro3202&stringBairro4202&stringBairro5202



next


stringBairro3202 = ""
stringBairro4202 = ""
stringBairro5202 = ""

	
	
	

	
	
	end if
	
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 'and Vila_vend ='"& rs("vila_comp") &"'
	 dim stringVilaComp202

	if rs202("vila_comp") <> "n�o informado" and rs202("vila_comp") <> "" and rs202("vila_comp") <> "vlqualquer" and rs202("vila_comp") <> "qualquer um" then
	stringVilaComp202 = ""
	else
	
	stringVilaComp202 = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp
  'if rs("tipo_comp")="n�o informado" or rs("tipo_comp")="" or rs("tipo_comp")="tqualquer" or rs("tipo_comp") = "qualquer um" then
	'stringTipoComp = ""
	'else
	
	
	'stringTipoComp = " and Tipo_vend ='"& rs("tipo_comp")&"'"
	'end if
	
	
	
	'--------------------------Tipo----------------------------

if rs202("tipo_comp") <> "qualquer um" and rs202("tipo_comp") <> "n�o informado" and rs202("tipo_comp") <> "" then




 
dim Numero_IndicacoesTipoComp202
dim Numero_Indicacoes02TipoComp202




Numero_IndicacoesTipoComp202 = 0
Numero_Indicacoes02TipoComp202 = 0


dim soma02TipoComp202
dim somaTipoComp202

somaTipoComp202 = 0
soma02TipoComp202 = 0

dim VariavelTipoComp202
dim RetornoTipoComp202
dim contarTipoComp202
VariavelTipoComp202 =  rs202("tipo_comp")
RetornoTipoComp202 = Split(rs202("tipo_comp"),", ")

contarTipoComp202=0

dim stringTipo3Comp202
dim stringTipo4Comp202
dim stringTipo5Comp202

for contarTipoComp202=0 to UBound(RetornoTipoComp202)

stringTipo3Comp202 = "and ( "
stringTipo4Comp202 = " tipo_vend='"&RetornoTipoComp202(contarTipoComp202)&"'or  " &stringTipo4Comp202

stringTipo5Comp202 = " cod_permuta=0)"


stringTipo2Comp202 = stringTipo3Comp202&stringTipo4Comp202&stringTipo5Comp202







next

stringTipo3Comp202 = ""
stringTipo4Comp202 = ""
stringTipo5Comp202 = ""


else
stringTipo2Comp202 = ""
end if

	
	
	
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp202
  
	
	stringQuartosComp202 = " and Quartos_vend >="& int(rs202("quartos_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 

	 dim stringVagasComp202
 
	
	stringVagasComp202 = " and vagas_vend >="& int(rs202("vagas_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------

dim PorcentualComp202

dim vValorMenorComp202
dim vValorMaiorComp202

PorcentualComp202 = int(rs202("valor_comp"))*20/100

   


   vValorMenorComp202 = int(rs202("valor_comp")) - int(PorcentualComp202)
   vValorMaiorComp202 = int(rs202("valor_comp")) + int(PorcentualComp202)


	 dim stringValorComp202
  
	
	
	'stringValorComp202 = " and Valor_vend >="& vValorMenorComp202 &" and Valor_vend <="& vValorMaiorComp202 &""
	
	stringValorComp202 = " and Valor_vend <="& int(vValorMaiorComp202) &""
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	'strSQL444 = "SELECT * FROM permuta"
	'&stringVilaVend
	'&stringVilaComp
	varIndicacaoCodigo202=rs202("cod_permuta")
	
	dim strSQL2
	
	strSQL2 = "SELECT permuta.cod_permuta,permuta.nome,permuta.atendimento   FROM permuta"&stringIndex202&stringCidadeVend202&stringBairroVend202&stringVilaVend202&stringTipoVend202&stringQuartosVend202&stringVagasVend202&stringValorVend202&stringCidadeComp202&stringBairroComp202&stringVilaComp202&stringTipo2Comp202&stringQuartosComp202&stringVagasComp202&stringValorComp202&" and standby <> 'incluido' and cod_permuta not like "&varIndicacaoCodigo202
	
 
	
'---------------------------------------------------------------	
	
	
	'strSQL2 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringVilaVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringVilaComp&stringTipo2Comp&stringQuartosComp&stringVagasComp&stringValorComp&" and cod_permuta not like "&varCodPermuta
	
	if vNome = "" then
	vNome = "n�o informado"
	end if
	
	if vTelefone = "" then
	vTelefone = "n�o informado"
	end if
	
	
	 dim vEnderecoIP , vData
  vData = now()
  
 
 vEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	
	
  
  '--------------incluir conta acessada-----------------
 
  dim JaComprador
	 
	 JaComprador = request.querystring("JaComprador")
	 
	 if JaComprador <> "" then
	'Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP2 &"','"& now() &"')"
	JaComprador = "JaExiste"
     else
	 
	 'JaComprador = "JaExiste"
	 Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data,atendimento) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_permuta") &"','"& "Permuta" &"','"& EnderecoIP2 &"','"& now() &"','"& rs("atendimento") &"')"
	
	JaComprador = "JaExiste"
	 end if
  
 
 
  
  
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 

dim rs2

Set RS2 = Server.CreateObject("ADODB.Recordset")
'um objeto recordset � inst�nciado.

Dim LinkTemp
'essa vari�vel vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
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
	
RS2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

RS2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

RS2.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conex�o o recordset utilizar�.
	
RS2.Open strSQL2, Conn, 1, 3
'o recordset � aberto
	
RS2.PageSize = 10
'Aqui configura-se o recordset para 20 registros por p�gina.

RS2.CacheSize = RS2.PageSize
'o Cache tamb�m conter� 20 registros por p�gina.

intPageCount = RS2.PageCount
'A vari�vel intPageCount recebe o valor do n�mero de p�gina do recordset retornado.

intRecordCount = RS2.RecordCount
'A vari�vel intRecordCount recebe o valor do n�mero de registros retornados no recordset.



If NOT (RS2.BOF AND RS2.EOF) Then
'verifica se existem registros retornados.
%>


<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Veja 
      abaixo , as indica��es de permuta para voc�.</font></strong> <br>
	<br></td>
  </tr>
</table>




<%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage � maior que o n�mero de p�ginas ent�o intPage � igual ao n�mero de p�ginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage � menor ou igual a zero ent�o intPage igual a "1"
	'a vari�vel intPage sempre vai ser for�ada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados ent�o.
			 
			 RS2.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a p�gina exata que o registro atual
			'reside
			
			intStart = RS2.AbsolutePosition
			'a vari�vel intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posi��o exata do primeiro registro da p�gina correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage � igual ao n�mero de p�ginas no recordset , estamos na �ltima 
			'p�gina ent�o.
				intFinish = intRecordCount
				'a vari�vel intFinish recebe o valor do n�mero do �ltimo recordset.
				'intFinish corresponde ao valor do �ltimo registro da p�gina correspondente.
			Else
				intFinish = intStart + (RS2.PageSize - 1)
				'a vari�vel intFinish recebe o valor de intStart + o valor
				'do n�mero de registros na p�gina menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros ent�o
		For intRecord = 1 to RS2.PageSize
		'um contador inRecord � colocado at� o n�mero de registros na p�gina.

%>
<br>
<% varCodPermuta =RS2("cod_permuta") %>
<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="170"><table width="784" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td style="border:1px solid #ddddc5;"><table width="774" height="170" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e9dca8"><div align="center"><font face="Verdana, arial" size="2" color="FFFFFF"><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')" style="color:#9d9249;text-decoration:none;"><strong>Ol�, 
                      meu nome &eacute; <%=rs2("nome")%> ,o sitema VEJA analizou 
                      os dados do seu e do meu im�vel e dectetou a possibilidade 
                      de efetuarmos uma permuta entre nossos im�veis. Lique j� 
                      para 4123-72-44 e fale com meu atendente sr(a) <%=rs2("atendimento")%>. 
                      para que cada um de n�s visitemos os im�veis de um e de 
                      outro, para ver mais detalhes do meu im�vel clique aqui, 
                      muito obrigado.</strong></a> </font></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>


 <%
RS2.MoveNext


	  




If RS2.EOF Then Exit for
Next	
%>


<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><font face="Verdana, arial" size="1"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se a p�gina atual for maior que "1" ent�o o link anteriro � colocado na 
			  na tela .-->
               <a href="?page=<%=intPage - 1%>&varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" style="color:#000000">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" >
              
			  <!-- se p�gina atual � menor que o total de p�ginas e intPage maior que um
			  ou seja, se n�o estiver na primeira p�gina e nem na �ltima ent�o. -->
			  
             
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage � menor que o n�mero de p�ginas ent�o colocar o bot�o pr�ximo -->
             <a href="?page=<%=intPage + 1%>&varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" style="color:#000000"><b>Pr�ximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table>

<%End If


Else

%>
<strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">N�o foram 
encontradas permuta para o seu im�vel.</font></strong> 
<%end if%>
<% end if%>
<%
Function EscreveFuncaoJavaScript ( Conexao )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsMarcas3.ActiveConnection = Conexao
	
	
	rsMarcas3.Open SqlMarcas3, Conexao
	

While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"




Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsCarros3.ActiveConnection = Conexao
	
	
	rsCarros3.Open SqlCarros3, Conexao
	
	





'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros3.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



rsMarcas3.close

set rsMarcas3 = nothing


rsCarros3.close

set rsCarros3 = nothing





End Function



%>
<%
Function EscreveFuncaoJavaScript2 ( Conexao )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo3.options[doublecombo.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas4 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas4 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsMarcas4.ActiveConnection = Conexao
	
	
	rsMarcas4.Open SqlMarcas4, Conexao




While NOT rsMarcas4.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas4("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas4("id_combo1")&" order by nome_combo2"





Set rsCarros4 = Server.CreateObject("ADODB.RecordSet")

	rsCarros4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsCarros4.ActiveConnection = Conexao
	
	
	rsCarros4.Open SqlCarros4, Conexao
	
	




'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros4.EoF

Response.Write "doublecombo.combo4.options[" & i & "] = new Option('" & rsCarros4("nome_combo2") & "','" & rsCarros4("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros4.MoveNext
Wend
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas4.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



rsMarcas4.close

set rsMarcas4 = nothing

rsCarros4.close

set rsCarros4 = nothing



End Function
%>
<%  EscreveFuncaoJavaScript ( Conexao ) %>
<%  EscreveFuncaoJavaScript2 ( Conexao ) %>
</body>
</html>
