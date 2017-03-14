<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%response.Buffer = true %>









<html>

<!--#include file="style4_imoveis.asp"-->

<title>Indicação de permuta</title>
<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=claro%>;}
</STYLE>

<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (doublecombo) 



{




{
if (doublecombo.combo1.value == "cqualquer") {
		alert("Você precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}
}
}


</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=600,height=510,resizable=yes,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow222(abrejanela222) {
   openWindow222 = window.open(abrejanela222,'openWin222','width=600,height=510,resizable=yes,scrollbars=yes')
   openWindow222.focus( )
   }

</SCRIPT>


</head>












<body onLoad="funScroll()" bgcolor="EAA813" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<center>
</center><br>
<center>
  <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

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


 varIndicacaoCidadeVend=request.querystring("varIndicacaoCidadeVend")
 varIndicacaoBairroVend=request.querystring("varIndicacaoBairroVend")
 varIndicacaoVilaVend=request.querystring("varIndicacaoVilaVend")
 varIndicacaoQuartosVend=request.querystring("varIndicacaoQuartosVend")
 varIndicacaoVagasVend=request.querystring("varIndicacaoVagasVend")
 varIndicacaoValorVend=request.querystring("varIndicacaoValorVend")
 varIndicacaoTipoVend=request.querystring("varIndicacaoTipoVend")
 
 
 
 session("varIndicacaoCidadeVend") = varIndicacaoCidadeVend
 session("varIndicacaoBairroVend") = varIndicacaoBairroVend
 session("varIndicacaoVilaVend") = varIndicacaoVilaVend
 session("varIndicacaoQuartosVend") = varIndicacaoQuartosVend
 session("varIndicacaoVagasVend") = varIndicacaoVagasVend
 session("varIndicacaoValorVend") = varIndicacaoValorVend
 session("varIndicacaoTipoVend") = varIndicacaoTipoVend
 
 
 
 
 
 
 varIndicacaoCidadeComp=request.querystring("varIndicacaoCidadeComp")
 varIndicacaoBairroComp=request.querystring("varIndicacaoBairroComp")
 varIndicacaoVilaComp=request.querystring("varIndicacaoVilaComp")
 varIndicacaoQuartosComp=request.querystring("varIndicacaoQuartosComp")
 varIndicacaoVagasComp=request.querystring("varIndicacaoVagasComp")
 varIndicacaoValorComp=request.querystring("varIndicacaoValorComp")
 varIndicacaoTipoComp=request.querystring("varIndicacaoTipoComp")
 
 
 session("varIndicacaoCidadeComp") = varIndicacaoCidadeComp
 session("varIndicacaoBairroComp") = varIndicacaoBairroComp
 session("varIndicacaoVilaComp") = varIndicacaoVilaComp
 session("varIndicacaoQuartosComp") = varIndicacaoQuartosComp
 session("varIndicacaoVagasComp") = varIndicacaoVagasComp
 session("varIndicacaoValorComp") = varIndicacaoValorComp
 session("varIndicacaoTipoComp") = varIndicacaoTipoComp
 
 
 
 varIndicacaoCodigo=request.querystring("varIndicacaoCodigo")
 
session("varIndicacaoCodigo") = varIndicacaoCodigo


 
 
 
 
 
 
 
 
 
 
 '------------------------Sua Cidade--------------------------


 dim stringIndex,stringCidadeVend
 
 stringIndex = " where cod_permuta<>"&"0"&""
 
 
 
  if   varIndicacaoCidadeVend = "não informado" or varIndicacaoCidadeVend = "" or varIndicacaoCidadeVend = "cqualquer" or  varIndicacaoCidadeVend = "qualquer um" then
	stringCidadeVend = ""
 else

stringCidadeVend = " and (Cidade_comp='"&varIndicacaoCidadeVend&"' or Cidade_comp='"&"não informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------






dim stringBairroVend

 if   varIndicacaoBairroVend = "não informado" or varIndicacaoBairroVend = "" or varIndicacaoBairroVend = "bqualquer" or  varIndicacaoBairroVend = "qualquer um" then
	stringBairroVend = ""
 else
'stringBairroVend = ""
stringBairroVend = " and (Bairro_comp like'%"&varIndicacaoBairroVend&"%' or Bairro_comp like'%"&"não informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if








 'if   varIndicacaoBairroVend = "não informado" or varIndicacaoBairroVend = "" or varIndicacaoBairroVend = "bqualquer" or  varIndicacaoBairroVend = "qualquer um" then
	'stringBairroVend = ""
 'else

'stringBairroVend = " and (Bairro_comp='"&varIndicacaoBairroVend&"' or Bairro_comp='"&"não informado"&"' or Bairro_comp='"&"bqualquer"&"' or Bairro_comp='"&"qualquer um"&"')"

'end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend
'" and (Vila_comp='"&varIndicacaoVilaVend&"' or Vila_comp='"&"não informado"&"')"

 if    varIndicacaoVilaVend = "não informado" or varIndicacaoVilaVend = "" or varIndicacaoVilaVend = "vlqualquer" or  varIndicacaoVilaVend = "qualquer um" then
	stringVilaVend =""
 else

stringVilaVend = " "
end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend
 
 
 if  varIndicacaoTipoVend = "não informado" or varIndicacaoTipoVend = "" or varIndicacaoTipoVend = "tqualquer" or  varIndicacaoTipoVend = "qualquer um" then

stringTipoVend = ""

else
stringTipoVend = " and Tipo_comp like '%"&varIndicacaoTipoVend&"%'"
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend
 
 
 

stringQuartosVend = " and Quartos_comp <="&int(varIndicacaoQuartosVend)&""

 

'-----------------------Número de vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend
 
 
 

stringVagasVend = " and vagas_comp <="&int(varIndicacaoVagasVend)&""
 
 
 
 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 dim PorcentualVend

dim vValorMenorVend
dim vValorMaiorVend

PorcentualVend = int(varIndicacaoValorVend)*20/100

   


   vValorMenorVend = int(varIndicacaoValorVend) - int(PorcentualVend)
   vValorMaiorVend = int(varIndicacaoValorVend) + int(PorcentualVend)

 
 
	 dim stringValorVend
 
	
	'stringValorVend = " and Valor_comp >="& vValorMenorVend &" and Valor_comp <="& vValorMaiorVend &""
  
    stringValorVend = " and Valor_comp >="& int(vValorMenorVend) &" "
  
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp
  if varIndicacaoCidadeComp = "não informado" or varIndicacaoCidadeComp = "" or varIndicacaoCidadeComp = "cqualquer" or  varIndicacaoCidadeComp = "qualquer um" then
	stringCidadeComp = ""
	else
	
	stringCidadeComp = " and Cidade_vend ='"& varIndicacaoCidadeComp &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp

	if varIndicacaoBairroComp = "não informado" or varIndicacaoBairroComp = "" or varIndicacaoBairroComp = "bqualquer" or  varIndicacaoBairroComp = "qualquer um"  then
	stringBairroComp = ""
	else
	
	'stringBairroComp = " and Bairro_vend ='"& varIndicacaoBairroComp &"'"
	
	
		
 
dim Numero_Indicacoes
dim Numero_Indicacoes02




Numero_Indicacoes = 0
Numero_Indicacoes02 = 0


dim soma02
dim soma

soma = 0
soma02 = 0

dim Variavel
dim Retorno
dim contar
Variavel = varIndicacaoBairroComp
Retorno = Split(Variavel,", ")

contar=0

dim stringBairro3
dim stringBairro4
dim stringBairro5

for contar=0 to UBound(Retorno)

stringBairro3 = "and ( "
stringBairro4 = " Bairro_vend='"&Retorno(contar)&"'or  " &stringBairro4

stringBairro5 = " cod_permuta=0)"


stringBairroComp = stringBairro3&stringBairro4&stringBairro5

next




	
stringBairro3 = ""
stringBairro4 = ""
stringBairro5 = ""
	
	
	
	
	
	
	end if
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 '" and Vila_vend ='"& varIndicacaoVilaComp &"'"
	 dim stringVilaComp

	if varIndicacaoVilaComp = "não informado" or varIndicacaoVilaComp = "" or varIndicacaoVilaComp = "vlqualquer" or  varIndicacaoVilaComp = "qualquer um" then
	stringVilaComp = ""
	else
	
	stringVilaComp = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp

	'if  varIndicacaoTipoComp = "não informado" or varIndicacaoTipoComp = "" or varIndicacaoTipoComp = "tqualquer" or  varIndicacaoTipoComp = "qualquer um" then

'stringTipoComp = ""

'else
	
	'stringTipoComp = " and Tipo_vend ='"&varIndicacaoTipoComp &"'"
	'end if
	
	'--------------------------Tipo----------------------------

if varIndicacaoTipoComp <> "qualquer um" and varIndicacaoTipoComp <> "não informado" and varIndicacaoTipoComp <> "não informado" then




 
dim Numero_IndicacoesTipoComp
dim Numero_Indicacoes02TipoComp




Numero_IndicacoesTipoComp = 0
Numero_Indicacoes02TipoComp = 0


dim soma02TipoComp
dim somaTipoComp

somaTipoComp = 0
soma02TipoComp = 0

dim VariavelTipoComp
dim RetornoTipoComp
dim contarTipoComp
VariavelTipoComp =  varIndicacaoTipoComp
RetornoTipoComp = Split(varIndicacaoTipoComp,", ")

contarTipoComp=0

dim stringTipo3Comp
dim stringTipo4Comp
dim stringTipo5Comp

for contarTipoComp=0 to UBound(RetornoTipoComp)

stringTipo3Comp = "and ( "
stringTipo4Comp = " tipo_vend='"&RetornoTipoComp(contarTipoComp)&"'or  " &stringTipo4Comp

stringTipo5Comp = " cod_permuta=0)"


stringTipo2Comp = stringTipo3Comp&stringTipo4Comp&stringTipo5Comp







next

stringTipo3Comp = ""
stringTipo4Comp = ""
stringTipo5Comp = ""


else
stringTipo2Comp = ""
end if

	
	
	
	
	
 
 
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp
  
	
	stringQuartosComp = " and Quartos_vend >="& int(varIndicacaoQuartosComp) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
  '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp
 
	
	stringVagasComp = " and vagas_vend >="& int(varIndicacaoVagasComp) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------


dim PorcentualComp

dim vValorMenorComp
dim vValorMaiorComp

PorcentualComp = int(varIndicacaoValorComp)*20/100

   


   vValorMenorComp = int(varIndicacaoValorComp) - int(PorcentualComp)
   vValorMaiorComp = int(varIndicacaoValorComp) + int(PorcentualComp)


	 dim stringValorComp
  
	
	'varIndicacaoValorComp
	'stringValorComp = " and Valor_vend >="& vValorMenorComp &" and Valor_vend <="& vValorMaiorComp &""
	stringValorComp = "  and Valor_vend <="& vValorMaiorComp &""
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais   FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringVilaVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringVilaComp&stringTipo2Comp&stringQuartosComp&stringVagasComp&stringValorComp&" and standby <> 'incluido' and cod_permuta not like "&varIndicacaoCodigo
	
	
	
	if vNome = "" then
	vNome = "não informado"
	end if
	
	if vTelefone = "" then
	vTelefone = "não informado"
	end if
	
	
	 dim vEnderecoIP , vData
  vData = now()
  
 
 vEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	
	
  
  
  
  
  
  
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
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.



If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%>
  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Foram encontrados 
  <%=intRecordCount%> permutantes</font></strong> <br>
   
  <br>
  <table width="537" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="537" height="36"><table width="537" border="0" cellspacing="0" cellpadding="0">
        
		<%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.

	If CInt(intPage) <= 0 Then intPage = 1
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

dim varCodPermuta



dim Conexao2,rs7

	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL7
	
	
	


 dim Conexao22,rs77

	Set rs77 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL77
	
	
	dim vimagem
	if rs("cod_imovel") <> "não informado" and rs("cod_imovel") <> "" then
	 strSQL77 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where cod_imovel="&rs("cod_imovel")
	
	 rs77.CursorLocation = 3
      rs77.CursorType = 3
	  rs77.ActiveConnection = conn
	  
	 rs77.Open strSQL77, conn
	 
	 
	 
   if not rs77.eof then
   vimagem = rs77("foto_grande")
   else
   vimagem = "imovel00000.jpg"
  end if
	
	else
	vimagem = "imovel00000.jpg"
	end if
	
	
	 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	


%>

<% varCodPermuta =RS("cod_permuta") %>
 <tr>
            <td><table width="568" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="568" height="11"><img src="top_display2.jpg" width="568" height="11"></td>
  </tr>
  <tr> 
    <td width="568" height="153"><table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
          <td><table width="552" height="153" border="0" cellpadding="0" cellspacing="0" bgcolor="FE9225">
              <tr> 
                              <td width="552" height="16" bgcolor="FE9225"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Estou 
                                  interessado em im&oacute;vel na cidade de<strong> 
                                  <a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><font face="Verdana, arial" size="1" color="white"><%=RS("Cidade_comp")%></font></a> 
                                  </strong><strong> </strong></font><font face="Verdana, arial" size="1" color="white"><strong><%if rs("bairro_comp") = "bqualquer" then response.Write "" else response.write "no bairro de "&RS("Bairro_comp") end if %></strong></font></div></td>
              </tr>
              <tr> 
                              <td width="552" height="16" bgcolor="E17508"><div align="center"><a href="javascript:newWindow222('visualizar_permuta33.asp?varCodPermuta=<%=varCodPermuta%>')"> 
                                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
                                    mais detalhes</strong></font></div>
                                  </a></div></td>
              </tr>
              <tr> 
                <td><table width="552" height="115" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="173" bgcolor="FE9225"> 
                        <center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
                              <td bgcolor="<%=escuro%>"><% If objFSO.FileExists(Server.MapPath(vimagem)) = True Then %><a href="javascript:newWindow222('visualizar_permuta33.asp?varCodPermuta=<%=varCodPermuta%>')"><img src="<%=vimagem%>" width="158" height="90" border=0></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow222('visualizar_permuta22.asp?varCodPermuta=<%=varCodPermuta%>')" style="color:#FFFFFF"><strong>Foto não disponível</strong></a></font></div><% end if %></td>
                            </tr>
                          </table>
                        </center>
					  
					  </td>
                      <td bgcolor="FE9225"><div align="center"><font face="Verdana, arial" size="1" color="FFFFFF"><%=RS("descricao_vend")%><br><br><strong>Código da permuta <%=RS("cod_permuta")%></strong></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="568" height="11"><img src="bottom_display2.jpg" width="568" height="11"></td>
  </tr>
</table></td>
 
 </tr>
 <tr>
          <td height="18"> </td>
 </tr>
 
       
		
		 <%
RS.MoveNext


	  




If colorchanger = 1 Then
	colorchanger = 0
	color1 = "#537497"
	color2 = "#94ADC8"
Else
	colorchanger = 1
	color1 = "#94ADC8"
	color2 = "#537497"
End If

if corfonte = "black" then
 corfonte = "white"
 
 else
 
 corfonte = "white"
 end if
 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
		
		
		
		
      </table></td>
  </tr>
  
  <tr>
    <td width="537" height="18"><table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><font face="Verdana, arial" size="1"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
               <a href="?page=<%=intPage - 1%>&varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" style="color:#000000">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" >
              
			  <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
			  
             
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
             <a href="?page=<%=intPage + 1%>&varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" style="color:#000000"><b>Próximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>
</center>
<center>

</center>

<%End If


Else

%>
<center>
<Font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><Strong>Permuta não encontrada!!</strong></Font><br>
  
  
</center> 
  <% end if %>
  <%


RS.close
Set RS = Nothing

'---------------------------


set rs7 = nothing
'----------------------------



'---------------------------


set rs77 = nothing
'----------------------------



'---------------------------


set objfso= nothing
'----------------------------








%>


  
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
group2[3][12]=new Option("Acima de 400.000,00","0000400001 1000000000")









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





%>
  <% response.flush%>
  <%response.clear%>
  <!--#include file="dsn2.asp"-->

<br>



</body>
</html>
