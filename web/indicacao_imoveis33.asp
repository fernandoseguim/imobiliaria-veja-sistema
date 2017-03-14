<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_conta.asp"-->
<%

'Criando conexão com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn

dim objFSO

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")



dim varCod_imovel


varCod_imovel = request.QueryString("varCod_imovel")

dim strSQL

strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento  FROM imoveis where  cod_imovel="&varCod_imovel
	
dim rs

Set rs = Server.CreateObject("ADODB.RecordSet")	
	

rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao



rs.Open strSQL, Conexao	
	
	





%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela2) {
   openWindow2 = window.open(abrejanela2,'openWin2','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow2.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2121(abrejanela2121) {
   openWindow2121 = window.open(abrejanela2121,'openWin2121','width=650,height=600,resizable=yes,scrollbars=yes,Left=0,Top=0')
   openWindow2121.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2323(abrejanela2323) {
   openWindow2323 = window.open(abrejanela2323,'openWin2323','width=400,height=600,resizable=yes,scrollbars=yes,Left=0,Top=0')
   openWindow2323.focus( )
   }

</SCRIPT>



<title>Conta de imóvel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF">

<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_conta_imovel01.asp?varCod_imovel=<%=varCod_imovel%>&vPerguntaPermuta=<%=vPerguntaPermuta%>&vPerguntaCompradores=<%=vPerguntaCompradores%>">
</form>

  <%

dim stringCidade2,stringBairro2,stringNegociacao2,stringQuartos2,stringValor2,stringTipo2
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


varIndicacaoCidade = rs("cidade")
varIndicacaoBairro = rs("bairro")
varIndicacaoNegociacao = rs("negociacao")
varIndicacaoQuartos = rs("quartos")
varIndicacaoVagas = rs("vagas")
varIndicacaoTipo = rs("tipo")
varIndicacaoValor = rs("valor")
varIndicacaoValor = int(varIndicacaoValor)
vValorMenor = int("0")
vValorMaior = int("0")

dim porcentual



%>
  <%





'dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

'dim negrito,negrito2,varCodComprador



	
 
 
'------------------------Cidade---------------------------

stringIndex2 = " where cod_compradores<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "não informado" and rs("cidade") <> "" then
stringCidade2 = " and (cidade='"&rs("cidade")&"' or cidade='"&"não informado"&"')"
else
stringCidade2 = ""
end if



 '--------------------------Bairro----------------------------

if rs("bairro") <> "qualquer um" and rs("bairro") <> "não informado" and rs("bairro") <> "" then
stringBairro2 = " and (Bairro like '%"&rs("bairro")&"%' or Bairro like'%"&"não informado"&"%')"
else
stringBairro2 = ""
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" and  rs("tipo") <> "não informado" and rs("tipo") <> "" then
stringTipo2 = " and Tipo like '%"&rs("Tipo")&"%'"
else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negociação---------------------------
'vNegocio = "Compra"
if rs("negociacao") = "venda" then
vNegocio = "compra"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if

if  rs("negociacao") <> "qualquer um" and rs("negociacao") <> "" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> 0 and rs("quartos") <> "" then
stringQuartos2 = " and quartos<="&rs("quartos")&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> 0 and rs("vagas") <> "" then
stringVagas2 = " and vagas <="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------




if rs("valor") <> "" and rs("valor") <> "0,00" and rs("valor") <> "0" then
'---------------------------------Valor-----------------------------------



 
   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)
  








stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

else

stringValor2 = ""

end if

'---------------------------------Condominio-----------------------------------



dim stringCondominio101


Porcentual02101 = int(rs("condominio"))*10/100
   


   vCondominioMenor101 = int(rs("condominio")) - int(Porcentual02101)
   vCondominioMaior101 = int(rs("condominio")) + int(Porcentual02101)




if  int(rs("condominio")) <> 0 and rs("condominio") <> ""  then
stringCondominio101 = " and Condominio >="& int(rs("condominio")) &" "
else
stringCondominio101 = ""
end if


'---------------------------------------------------------------------------


'---------------------------------Área Total-----------------------------------



dim stringAreaTotal101


Porcentual03101 = int(rs("area_total"))*10/100
   


   vAreaTotalMenor101 = int(rs("area_total")) - int(Porcentual03101)
   vAreaTotalMaior101 = int(rs("area_total")) + int(Porcentual03101)



if  int(rs("area_total")) <> 0 and rs("area_total") <> "" then
stringAreaTotal101 = " and area_total >="& vAreaTotalMenor101 &" and area_total <="& vAreaTotalMaior101 &""
else
stringAreaTotal101 = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Suítes-----------------------------------------


dim stringSuites101
 
if  rs("suites") <> "suiqualquer" and rs("suites") <> "não" and rs("suites") <> "0" and rs("suites") <>"00" and rs("suites") <>"" then
stringSuites101 = "  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites101 = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina101
 
if  rs("piscina") <> "pisqualquer" and rs("piscina") <> "não" and rs("piscina") <> "0" and rs("piscina") <>"00" and rs("piscina") <>"" then
stringPiscina101 = "  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina101 = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria101
 
if  rs("portaria") <> "porqualquer" and rs("portaria") <> "não" and rs("portaria") <> "0" and rs("portaria") <>"00" and rs("portaria") <>"" then
stringPortaria101 = "  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria101 = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal101
 
if  rs("quintal") <> "quiqualquer" and rs("quintal") <> "não" and rs("quintal") <> "0" and rs("quintal") <>"00" and rs("quintal") <>"" then
stringQuintal101 = "  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal101 = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras101
 
if  rs("quadras") <> "quaqualquer" and rs("quadras") <> "não" and rs("quadras") <> "0" and rs("quadras") <>"00" and rs("quadras") <>"" then
stringQuadras101 = "  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras101 = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula101
 
if  rs("edicula") <> "ediqualquer" and rs("edicula") <> "não" and rs("edicula") <> "0" and rs("edicula") <>"00" and rs("edicula") <>"" then
stringEdicula101 = "  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula101 = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Ocupação--------------------------------------
dim stringOcupacao101
 
if  rs("ocupacao") <> "oqualquer" and rs("ocupacao") <> "não informado"  then
stringOcupacao101 = "  and ocupacao ='"&rs("ocupacao")&"'  and ocupacao IS NOT NULL"




else

stringOcupacao101 = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

'stringStandby = " and standby like '"&"suspenso"&"' and standby like '"&"comprador OK"&"'"

stringStandby = " and ( standby like 'comprador OK') "








'---------------------------------------------------------------------------



	'strSQL444 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.suites,compradores.piscina,compradores.portaria,compradores.area_total,compradores.quintal,compradores.quadras,compradores.condominio,compradores.edicula,compradores.clique   FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio101&stringAreaTotal101&stringSuites101&stringPiscina101&stringPortaria101&stringQuintal101&stringQuadras101&stringEdicula101&stringOcupacao101&stringStandby
	
	dim strSQL2
	strSQL2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.suites,compradores.piscina,compradores.portaria,compradores.area_total,compradores.quintal,compradores.quadras,compradores.condominio,compradores.edicula,compradores.clique   FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	




'strSQL2 ="SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	


'----------------------------------------------------Fim da instrução SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 
  
 
 
 '--------------incluir conta acessada-----------------
 
  dim JaComprador
	 
	 JaComprador = request.querystring("JaComprador")
	 
	 if JaComprador <> "" then
	'Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP2 &"','"& now() &"')"
	JaComprador = "JaExiste"
     else
	 
	 'JaComprador = "JaExiste"
	 Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data,atendimento) values( '"& rs("proprietario") &"','"& rs("telefone") &"','"& rs("cod_imovel") &"','"& "Imóvel" &"','"& EnderecoIP2 &"','"& now() &"','"& rs("captacao") &"')"
	
	JaComprador = "JaExiste"
	 end if
  
 
 
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 
dim rs2


Set RS2 = Server.CreateObject("ADODB.Recordset")
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
	
RS2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS2.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	
RS2.Open strSQL2, Conn, 1, 3
'o recordset é aberto
	
RS2.PageSize = 5
'Aqui configura-se o recordset para 20 registros por página.

RS2.CacheSize = RS2.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS2.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS2.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.

If  (RS2.BOF AND RS2.EOF) Then
response.write "Não foram encontrados compradores para o seu imóvel."
end if

If NOT (RS2.BOF AND RS2.EOF) Then
'verifica se existem registros retornados.
%>
 <table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"> <strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Veja as 
  indicações de compradores para o seu imóvel.</font></strong> <br>
  <br></td>
  </tr>
</table>

  <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount 
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.

	If CInt(intPage) <= 0 Then intPage = 1 
	'se intPage é menor ou igual a zero então intPage igual a "1"
	'a variável intPage sempre vai ser forçada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados então.
			 
			 RS2.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a página exata que o registro atual
			'reside
			
			intStart = RS2.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posição exata do primeiro registro da página correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual ao número de páginas no recordset , estamos na última 
			'página então.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor do número do último recordset.
				'intFinish corresponde ao valor do último registro da página correspondente.
			Else
				intFinish = intStart + (RS2.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'do número de registros na página menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros então
		For intRecord = 1 to RS2.PageSize
		'um contador inRecord é colocado até o número de registros na página.
%>
  <% varCodCompradores = rs2("cod_compradores") %>

<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="170"><table width="784" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td style="border:1px solid #ddddc5;"><table width="774" height="170" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e9dca8"><table width="774" height="170" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="20"><div align="center"><font color="#996600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Procuro 
                            por : <%=rs2("cidade")%> ,<%=rs2("bairro")%> , <%=rs2("tipo")%> 
                            com<%=rs2("vagas")%> vagas na garagem e <%=rs2("quartos")%> 
                            dormitórios no valor de <%=FormatNumber(rs2("valor"),2)%></strong></font></div></td>
                      </tr>
                      <tr>
                        <td height="150"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('visualizar_comprador01.asp?varCodCompradores=<%=varCodCompradores%>')" style="color:#996600;text-decoration:none;"><%=rs2("descricao")%> 
                            <br>
                      <strong>Se quiser saber mais clique aqui.</strong></a></font></div>
					  <br>
					  <%
					 
					
					 
					  SqlImoveis001 = "SELECT imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.cod_imovel,imoveis.tipo,imoveis.valor FROM imoveis where telefone like '"&rs2("telefone")&"' or telefone02 like '"&rs2("telefone")&"' or telefone03 like '"&rs2("telefone")&"' ORDER BY cod_imovel ASC" 

Set rsImoveis001 = Server.CreateObject("ADODB.RecordSet")

	rsImoveis001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsImoveis001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsImoveis001.ActiveConnection = Conexao
	
	
	rsImoveis001.Open sqlImoveis001, Conexao
					  
		if not rsImoveis001.eof then 
		
		
		
		
		
		
					  %>
					  <br>
					  
					  
					  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Aten&ccedil;&atilde;o 
                          este comprador tem <%=rsImoveis001("tipo")%> no valor 
                          de <%=formatnumber(rsImoveis001("valor"),2)%> reais 
                          para entrar no neg&oacute;cio,clique no texto acima 
                          para saber mais.</font><br>
                          </strong></font></div>
					  <%
					  end if
					  
					  rsImoveis001.close

                     set rsImoveis001 = nothing
					  
					  
					  %>
					  </td>
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
RS2.MoveNext


	  





 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS2.EOF Then Exit for

Next

	
%>

<%end if%>
<%end if%>

<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"><table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&varCod_imovel=<%=varCod_imovel%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoValor=<%=varIndicacaoValor%>&JaComprador=<%=JaComprador%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" size="1" > 
            </font></div></td>
          
        <td> 
          <%If cInt(intPage) < cInt(intPageCount)  Then%>
          <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
          <a href="?page=<%=intPage + 1%>&varCod_imovel=<%=varCod_imovel%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoValor=<%=varIndicacaoValor%>&JaComprador=<%=JaComprador%>"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Próximo</strong></font></a></td><% end if%>
        </tr>
      </table></td>
  </tr>
</table>



  

<%
Function EscreveFuncaoJavaScript ( Conexao )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 

Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas3.ActiveConnection = Conexao
	
	
	rsMarcas3.Open SqlMarcas3, Conexao



While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros3.ActiveConnection = Conexao
	
	
	rsCarros3.Open SqlCarros3, Conexao



'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros3.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
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


rsMarcas3.close

set rsMarcas3 = nothing


rsCarros3.close


set rsCarros3 = nothing


End Function
%> 

<%
Function EscreveFuncaoJavaScript2 ( Conexao )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo3.options[doublecombo.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas33 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas33.ActiveConnection = Conexao
	
	
	rsMarcas33.Open SqlMarcas33, Conexao





While NOT (rsMarcas33.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas33("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros33 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 WHERE id_combo1 =" & rsMarcas33("id_combo1")&" order by nome_combo2"






Set rsCarros33 = Server.CreateObject("ADODB.RecordSet")

	rsCarros33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros33.ActiveConnection = Conexao
	
	
	rsCarros33.Open SqlCarros33, Conexao





'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
 
While NOT (rsCarros33.EoF)

Response.Write "doublecombo.combo4.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend

Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 





rsMarcas33.Close           
		   
           Set rsMarcas33 = Nothing
             
			rsCarros33.Close           
		   
           Set rsCarros33 = Nothing 





End Function
%> 



<%  EscreveFuncaoJavaScript2 ( Conexao ) %>
</body>
</html>
