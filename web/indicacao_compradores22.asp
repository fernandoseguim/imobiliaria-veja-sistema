<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->


<%response.Buffer = true %>

<%

dim stringCidade2,stringBairro2,stringNegociacao2,stringQuartos2,stringVagas2,stringValor2,stringTipo2,stringIndex2
dim vNegocio
dim vValorMenor,vValorMaior
dim varCodIndicacao

dim varIndicacaoCidade
dim varIndicacaoBairro
dim varIndicacaoNegociacao
dim varIndicacaoQuartos
dim varIndicacaoVagas
dim varIndicacaoValor
dim varIndicacaoTipo

dim varIndicacaoCondominio
dim varIndicacaoAreaConstruida
dim varIndicacaoSuites
dim varIndicacaoPiscina
dim varIndicacaoPortaria
dim varIndicacaoQuintal
dim varIndicacaoQuadras
dim varIndicacaoEdicula
dim varIndicacaoOcupacao


varIndicacaoCidade = request.querystring("varIndicacaoCidade")
varIndicacaoBairro = request.querystring("varIndicacaoBairro")
varIndicacaoNegociacao = request.querystring("varIndicacaoNegociacao")
varIndicacaoQuartos = request.querystring("varIndicacaoQuartos")
varIndicacaoVagas = request.querystring("varIndicacaoVagas")
varIndicacaoTipo = request.querystring("varIndicacaoTipo")
varIndicacaoValor = request.querystring("varIndicacaoValor")

varIndicacaoCondominio = request.querystring("varIndicacaoCondominio")
varIndicacaoAreaConstruida = request.querystring("varIndicacaoAreaConstruida")
varIndicacaoSuites = request.querystring("varIndicacaoSuites")
varIndicacaoPiscina = request.querystring("varIndicacaoPiscina")
varIndicacaoPortaria = request.querystring("varIndicacaoPortaria")
varIndicacaoQuintal = request.querystring("varIndicacaoQuintal")
varIndicacaoQuadras = request.querystring("varIndicacaoQuadras")
varIndicacaoEdicula = request.querystring("varIndicacaoEdicula")
varIndicacaoOcupacao = request.querystring("varIndicacaoOcupacao")


session("varIndicacaoCidade") = varIndicacaoCidade 
session("varIndicacaoBairro") = varIndicacaoBairro 
session("varIndicacaoNegociacao") = varIndicacaoNegociacao 
session("varIndicacaoQuartos") = varIndicacaoQuartos 
session("varIndicacaoVagas") = varIndicacaoVagas 
session("varIndicacaoTipo") = varIndicacaoTipo 
session("varIndicacaoValor") = varIndicacaoValor 




varIndicacaoValor = int(varIndicacaoValor)



vValorMenor = int("0")
vValorMaior = int("0")

dim porcentual


%>

  <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2,objFSO

dim varCodCompradores

varCodCompradores = request.QueryString("varCodCompradores")

dim Conexao

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	  Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
 
 
 dim vDataAtual
 
 dim vDataAtual2
 
if len(now()) = 19 then
vDataAtual = left(now(),11)


end if


if len(now()) = 18 then
vDataAtual = left(now(),10)


end if


if len(now()) = 17 then
vDataAtual = left(now(),9)


end if




 '------------------------Cidade---------------------------

stringIndex2 = " where cod_imovel<>"&"0"&""


if  varIndicacaoCidade <> "qualquer um" and  varIndicacaoCidade <> "não informado" and   varIndicacaoCidade <> ""  then
stringCidade2 = " and cidade='"&varIndicacaoCidade&"'"
else
stringCidade2 = ""
end if

 
 


 '--------------------------Bairro----------------------------

if varIndicacaoBairro <> "qualquer um" and varIndicacaoBairro <> "não informado" and varIndicacaoBairro <> "" then





 
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
Variavel = varIndicacaoBairro
Retorno = Split(Variavel,", ")

contar=0

dim stringBairro3
dim stringBairro4
dim stringBairro5

for contar=0 to UBound(Retorno)

stringBairro3 = "and ( "
stringBairro4 = " Bairro='"&Retorno(contar)&"'or  " &stringBairro4

stringBairro5 = " cod_imovel=0)"


stringBairro2 = stringBairro3&stringBairro4&stringBairro5

next









else
stringBairro2 = ""
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if varIndicacaoTipo <> "qualquer um" and varIndicacaoTipo <> "tqualquer" and varIndicacaoTipo <> "" then
dim Numero_IndicacoesTipo
dim Numero_Indicacoes02Tipo




Numero_IndicacoesTipo = 0
Numero_Indicacoes02Tipo = 0


dim soma02Tipo
dim somaTipo

somaTipo = 0
soma02Tipo = 0

dim VariavelTipo
dim RetornoTipo
dim contarTipo
VariavelTipo =  varIndicacaoTipo 
RetornoTipo = Split(varIndicacaoTipo ,", ")

contarTipo=0

dim stringTipo3
dim stringTipo4
dim stringTipo5

for contarTipo=0 to UBound(RetornoTipo)

stringTipo3 = "and ( "
stringTipo4 = " tipo='"&RetornoTipo(contarTipo)&"'or  " &stringTipo4

stringTipo5 = " cod_imovel=0)"


stringTipo2 = stringTipo3&stringTipo4&stringTipo5







next

stringTipo3 = ""
stringTipo4 = ""
stringTipo5 = ""


else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negociação---------------------------
vNegocio = "venda"
if varIndicacaoNegociacao = "compra" then
vNegocio = "venda"
end if

if varIndicacaoNegociacao = "compra" then
vNegocio = "venda"
end if


if varIndicacaoNegociacao = "Aluguel" then
vNegocio = "aluguel"
end if

if varIndicacaoNegociacao = "aluguel" then
vNegocio = "aluguel"
end if


if  varIndicacaoNegociacao <> "qualquer um" and  varIndicacaoNegociacao <> "não informado" and varIndicacaoNegociacao <> "nqualquer" and varIndicacaoNegociacao <> "" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  varIndicacaoQuartos <> "0" and varIndicacaoQuartos <> "" then
stringQuartos2 = " and quartos >="&varIndicacaoQuartos&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------

'---------------------------Vagas------------------------------


if  varIndicacaoVagas <> "0" and varIndicacaoVagas <> "" then
stringVagas2 = " and vagas >="&varIndicacaoVagas&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------







'---------------------------------Valor-----------------------------------



 porcentual = int(varIndicacaoValor)*10/100

   vValorMenor = int(varIndicacaoValor)-int(porcentual)
   vValorMaior = int(varIndicacaoValor)+int(porcentual)
   



'stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

stringValor2 = " and Valor <="& vValorMaior &""




'---------------------------------Condominio-----------------------------------



dim stringCondominio3


Porcentual02 = int(varIndicacaoCondominio)*10/100
   


   vCondominioMenor = int(varIndicacaoCondominio) - int(Porcentual02)
   vCondominioMaior = int(varIndicacaoCondominio) + int(Porcentual02)




if  int(varIndicacaoCondominio) <> 0  then
'stringCondominio = " and Condominio >="& vCondominioMenor &" and Condominio <="& vCondominioMaior &""
'stringCondominio3 = " and Condominio <="&  vCondominioMaior &""
stringCondominio3 = ""
else
stringCondominio3 = ""
end if


'---------------------------------------------------------------------------


'---------------------------------Área Construida-----------------------------------



dim stringAreaConstruida


Porcentual03 = int(varIndicacaoAreaConstruida)*10/100
   


   vAreaConstruidaMenor = int(varIndicacaoAreaConstruida) - int(Porcentual03)
   vAreaConstruidaMaior = int(varIndicacaoAreaConstruida) + int(Porcentual03)



if  int(varIndicacaoAreaConstruida) <> 0 and varIndicacaoAreaConstruida <> "" then
'stringAreaTotal = " and area_total >="& vAreaTotalMenor &" and area_total <="& vAreaTotalMaior &""
stringAreaConstruida = " and area_construida >="& vAreaConstruidaMenor &""

else
stringAreaConstruida = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Suítes-----------------------------------------


dim stringSuites
 
if  varIndicacaoSuites <> "suiqualquer" and varIndicacaoSuites <> "não" and varIndicacaoSuites <> "0" and varIndicacaoSuites <>"00" and varIndicacaoSuites <>"" then
stringSuites = "  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina
 
if  varIndicacaoPiscina <> "pisqualquer" and varIndicacaoPiscina <> "não" and varIndicacaoPiscina <> "0" and varIndicacaoPiscina <>"00" and varIndicacaoPiscina <>"" then
stringPiscina = "  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria
 
if  varIndicacaoPortaria <> "porqualquer" and varIndicacaoPortaria <> "não" and varIndicacaoPortaria <> "0" and varIndicacaoPortaria <>"00" and varIndicacaoPortaria <>"" then
stringPortaria = "  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal
 
if  varIndicacaoQuintal <> "quiqualquer" and varIndicacaoQuintal <> "não" and varIndicacaoQuintal <> "0" and varIndicacaoQuintal <>"00" and varIndicacaoQuintal <>"" then
stringQuintal = "  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras
 
if  varIndicacaoQuadras <> "quaqualquer" and varIndicacaoQuadras <> "não" and varIndicacaoQuadras <> "0" and varIndicacaoQuadras <>"00" and varIndicacaoQuadras <>"" then
stringQuadras = "  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula
 
if  varIndicacaoEdicula <> "ediqualquer" and varIndicacaoEdicula <> "não" and varIndicacaoEdicula <> "0" and varIndicacaoEdicula <>"00" and varIndicacaoEdicula <>"" then
stringEdicula = "  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula = ""
end if



'--------------------------------------------------------------------------------




'--------------------------Ocupação--------------------------------------
dim stringOcupacao
 
if  varIndicacaoOcupacao <> "oqualquer" and varIndicacaoOcupacao <> "não informado" and  varIndicacaoOcupacao <> "" then
stringOcupacao = "  and ocupacao ='"&varIndicacaoOcupacao&"'  and ocupacao IS NOT NULL"




else

stringOcupacao = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

'stringStandby = "  and imovel_em_negociacao <>  '"&"Vendido pela Veja"&"' and imovel_em_negociacao <>  '"&"Vendido por outros"&"' and imovel_em_negociacao <>  '"&"Suspenso"&"' and imovel_em_negociacao <>  '"&"Com proposta"&"' and (imovel_em_negociacao <>  '"&"incluido"&"' or imovel_em_negociacao IS NULL)"

'stringStandby = " and (imovel_em_negociacao like  '"&"Suspenso"&"' or imovel_em_negociacao like  '"&"imóvel OK"&"' or imovel_em_negociacao like  '"&"Imóvel a recaptar"&"') "

 stringStandby = " and ( imovel_em_negociacao like  '"&"imóvel OK"&"') "

 
 
 
 
 
 
 '------------------------------------------------------------------
 
 
 
 




strSQL ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento,imoveis.origem_franquia FROM imoveis"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio3&stringAreaConstruida&stringSuites&stringPiscina&stringPortaria&stringQuintal&stringQuadras&stringEdicula&stringStandby&stringOcupacao&" ORDER  BY cod_imovel DESC"
	

'----------------------------------------------------Fim da instrução SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 
  
  
  
  
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




'----------------soma os recordsets---------------------------------



 soma = rs.recordcount
 
 if soma02 = "" then
 soma02 = soma
 else
 
 soma02 = int(soma02) + int(soma)
 
 end if






'-----------------------------------------------------------------------







%>

    <html>
	<!--#include file="style4_imoveis.asp"-->
	<title>Indicações</title>
	<head>
	<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: #DF700F;}
</STYLE>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow33(abrejanela33) {
   openWindow33 = window.open(abrejanela33,'openWin33','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow33.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow356(abrejanela356) {
   openWindow356 = window.open(abrejanela356,'openWin356','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow356.focus( )
   }

</SCRIPT>

	
	</head>
	<body bgcolor="EAA813">
	
	
	
	
<div align="center"><br>
  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Foram encontradas 
  <%=rs.recordcount%> indica&ccedil;&otilde;es</font></strong><br>
</div>
<br>

<br>


<br>

<table width="708" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  
  

  
  
    <td width="518" height="155">
	
	  
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
%>
<% varCodimovel = rs("COD_Imovel") %>
  
  
  
  
	
	
	
	
	
	
	  <table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="708" height="11"><img src="top_display2.jpg" width="708" height="11"></td>
  </tr>
  <tr> 
          <td width="708" height="153">
<table width="708" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="8" height="153"><img src="left_display2.jpg" width="8" height="155"></td>
          <td><table width="692" height="153" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                <td width="552" height="16" bgcolor="FE9225"><table width="692" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="130"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
                            <td width="140"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro/Região</strong></font></div></td>
                           <td width="140"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div></td>
						    <td width="99"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
                            <td width="80"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o</strong></font></div></td>
                            <td width="103"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
                          </tr>
                        </table></td>
              </tr>
              <tr> 
                <td width="552" height="16" bgcolor="E17508">
				<table width="692" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="130"><div align="center"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')"><font face="Verdana, arial" size="1" color="white"><%=RS("Cidade")%></font></a></div></td>
                            <td width="140"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Bairro")%></font></div></td>
                            <td width="140"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Vila")%></font></div></td>
						    <td width="99"><div align="center"><font face="Verdana, arial" size="1" color="white"><%if rs("tipo") = "tqualquer" then response.Write"não informado" else response.Write RS("Tipo") end if%></font></div></td>
                            <td width="80"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Negociacao")%></font></div></td>
                            <td width="103"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=FormatNumber(RS("Valor"),2)%></font></div></td>
                          </tr>
                        </table>
				
				
				
				</td>
              </tr>
              <tr> 
                <td bgcolor="FE9225"><table width="692" height="115" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                      <td width="173" bgcolor="FE9225"> 
                        <center>
                          <table width="160" height="93" border="1" cellpadding="0" cellspacing="0" bordercolor="6C3404">
                            <tr>
							<%
							
							
							vDataAtual2 = ""
							
							if len(rs("data_atualizacao")) = 19 then


                            vDataAtual2 = left(rs("data_atualizacao"),11)


                               end if


                             if len(rs("data_atualizacao")) = 18 then
                                    vDataAtual2 = left(rs("data_atualizacao"),10)


                                   end if


                             if len(rs("data_atualizacao")) = 17 then
                              vDataAtual2 = left(rs("data_atualizacao"),9)


                                end if
								
								
								if len(rs("data_atualizacao")) =0 then
								vDataAtual2 = "0/0/2007"
								end if
								
								
								
%>
							
                              <td bgcolor="<%=escuro%>"><% If objFSO.FileExists(Server.MapPath(rs("foto_pequena"))) = True Then %><a href="javascript:newWindow33('visualizar_imovel33.asp?varCod_imovel=<%=varCodimovel%>')"><img src="<%=rs("foto_pequena")%>" width="158" height="91" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow33('visualizar_imovel33.asp?varCod_imovel=<%=varCodimovel%>')" style="color:#FFFFFF"><strong>Foto não disponível</strong></a></font></div>
                                      <%end if%>
                                     
                                    </td>			  
							  
							  

							  
							  
							  
							  
							  
							  
							  
                                     
                                        
                                     
							  
							  
                            </tr>
                          </table>
                        </center>
					  
					  </td>
                      <td bgcolor="FE9225"><div align="center"><font face="Verdana, arial" size="1" color="#FFFFFF"><%=RS("obs_imovel")%><br>
                                <br>
                                <strong><a href="javascript:newWindow356('form_enviar_email22.asp?varCod_imovel=<%=rs("cod_imovel")%>&varCodCompradores=<%=varCodCompradores%>')"><img src="bt_email22.jpg" width="25" height="18" border="0"></a>Data 
                                de última atualização : <%=RS("data_atualizacao")%> 
                                || Data de inclusão: <%=RS("data")%> <br>
                                <br>
                                <a href="javascript:newWindow33('visualizar_imovel33.asp?varCod_imovel=<%=varCodimovel%>')" style="color:#FFFFFF">Veja 
                                a ficha completa na referência <%=RS("cod_imovel")%> 
                                </a></strong></font></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
                <td width="8" height="153"><img src="right_display2.jpg" width="8" height="155"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
          <td width="708" height="11"><img src="bottom_display2.jpg" width="708" height="11"></td>
  </tr>
</table>
<br>
	
	
	 <%
RS.MoveNext


	  





 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next



RS.close
Set RS = Nothing

	
%>
		
	
	
	
  </tr>
</table>
	
	</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>&varIndicacaoCondominio=<%=varIndicacaoCondominio%>&varIndicacaoAreaTotal=<%=varIndicacaoAreaTotal%>&varIndicacaoSuites=<%=varIndicacaoSuites%>&varIndicacaoPiscina=<%=varIndicacaoPiscina%>&varIndicacaoPortaria=<%=varIndicacaoPortaria%>&varIndicacaoQuintal=<%=varIndicacaoQuintal%>&varIndicacaoQuadras=<%=varIndicacaoQuadras%>&varIndicacaoEdicula=<%=varIndicacaoEdicula%>&varIndicacaoOcupacao=<%=varIndicacaoOcupacao%>&varCodCompradores=<%=varCodCompradores%>&varCod_imovel=<%=varCodimovel%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" size="1" > 
            </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%> 
            <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
            <a href="?page=<%=intPage + 1%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>&varIndicacaoCondominio=<%=varIndicacaoCondominio%>&varIndicacaoAreaTotal=<%=varIndicacaoAreaTotal%>&varIndicacaoSuites=<%=varIndicacaoSuites%>&varIndicacaoPiscina=<%=varIndicacaoPiscina%>&varIndicacaoPortaria=<%=varIndicacaoPortaria%>&varIndicacaoQuintal=<%=varIndicacaoQuintal%>&varIndicacaoQuadras=<%=varIndicacaoQuadras%>&varIndicacaoEdicula=<%=varIndicacaoEdicula%>&varIndicacaoOcupacao=<%=varIndicacaoOcupacao%>&varCodCompradores=<%=varCodCompradores%>&varCod_imovel=<%=varCodimovel%>"><b><font color="#000000" face="Verdana, arial" size="1">Próximo</font></b></a><a href=""> 
            </a> 
            <%End If%>
            </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>


<%End If


Else

%>
  <% 
response.write "<html><body bgcolor='EAA813'><br><br><br><center><font size='1' face='Verdana, Arial, Helvetica, sans-serif'><strong>Indicação não encontrada!</strong></font></center></body></html>"

%>


<br>

  <% end if %>
  





  <%

set objFSO = nothing



  Conexao.close
 Set Conexao = nothing
 


response.write vCondominioMaior

%>
  <% response.flush%>
  <%response.clear%>
  <!--#include file="dsn2.asp"-->
 
<br>


<%'varIndicacaoAreaConstruida%>
</body>
</html>
