<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_conta.asp"-->

<%
'dim Variavel
'dim VariavelTipo
%>
<%
Function EscreveFuncaoJavaScript ( Conexao )
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
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas3.ActiveConnection = Conexao
	
	
	rsMarcas3.Open SqlMarcas3, Conexao







While NOT (rsMarcas3.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"




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
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT (rsCarros3.EoF)

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




<%

'Criando conexão com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn





dim varCodCompradores


varCodCompradores = request.QueryString("varCodCompradores")

dim strSQL

strSQL = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento  FROM compradores where  cod_compradores="&varCodCompradores
	
dim rs

Set rs = Server.CreateObject("ADODB.RecordSet")	
	

rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao



rs.Open strSQL, Conexao


'-----------------------pegar código de permuta--------------------
 



'--------------------------------------------------------------------------------

%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela2) {
   openWindow2 = window.open(abrejanela2,'openWin2','width=650,height=600,resizable=yes,scrollbars=yes,Left=0,Top=0')
   openWindow2.focus( )
   }

</SCRIPT>


<title>Conta de comprador</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow555(abrejanela555) {
   openWindow555 = window.open(abrejanela555,'openWin555','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow555.focus( )
   }

</SCRIPT>


</head>
<% response.Buffer = true %>
<body bgcolor="#FFFFFF">


<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_conta_comprador01.asp?varCodCompradores=<%=varCodCompradores%>">
</form>


  
    
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





dim varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas

dim negrito,negrito2,objFSO

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	 
	
	
 
 
 


stringIndex2 = " where cod_imovel<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "não informado" and rs("cidade") <> ""   then
stringCidade2 = " and cidade='"&rs("cidade")&"'"
else
stringCidade2 = ""
end if

 '--------------------------Bairro----------------------------








if ( rs("bairro") <> "qualquer um" and  rs("bairro") <> "não informado" and  rs("bairro") <> "") then


 
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
Variavel =  rs("bairro")
Retorno = Split(rs("bairro"),", ")

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

stringBairro3 = ""
stringBairro4 = ""
stringBairro5 = ""


else
stringBairro2 = ""
end if








 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" and rs("tipo") <> ""  then




 
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
VariavelTipo =  rs("tipo")
RetornoTipo = Split(rs("tipo"),", ")

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
if rs("negociacao") = "Compra"  then
vNegocio = "venda"
end if

if rs("negociacao") = "compra" then
vNegocio = "venda"
end if

if rs("negociacao") = "Aluguel" then
vNegocio = "aluguel"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if


if  rs("negociacao") <> "qualquer um" and rs("negociacao") <> "" and rs("negociacao") <> "nqualquer" and rs("negociacao") <> "não informado" and rs("negociacao") <> "" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> int(0) and rs("quartos") <> "" then
stringQuartos2 = " and quartos >="&rs("quartos")&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> int(0) and rs("vagas") <> "" then
stringVagas2 = " and vagas >="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------------Valor-----------------------------------


   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)




'stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

stringValor2 = " and Valor <="& vValorMaior &""



'---------------------------------Condominio-----------------------------------



dim stringCondominio


Porcentual02 = int(rs("condominio"))*10/100
   


   vCondominioMenor = int(rs("condominio")) - int(Porcentual02)
   vCondominioMaior = int(rs("condominio")) + int(Porcentual02)




if  int(rs("condominio")) <> 0 and rs("condominio") <> ""  then

'stringCondominio = " and Condominio >="& vCondominioMenor &" and Condominio <="& vCondominioMaior &""
stringCondominio = "  and Condominio <="& vCondominioMaior &""

else
stringCondominio = ""
end if


'---------------------------------------------------------------------------


'---------------------------------Área Construida-----------------------------------



dim stringAreaConstruida


Porcentual03 = int(rs("area_construida"))*10/100
   


   vAreaConstruidaMenor = int(rs("area_construida")) - int(Porcentual03)
   vAreaConstruidaMaior = int(rs("area_construida")) + int(Porcentual03)



if  int(rs("area_construida")) <> 0 and rs("area_construida") <> "" then
'stringAreaTotal = " and area_total >="& vAreaTotalMenor &" and area_total <="& vAreaTotalMaior &""
stringAreaConstruida = " and area_construida >="& vAreaConstruidaMenor &""


else
stringAreaConstruida = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Suítes-----------------------------------------


dim stringSuites
 
if  rs("suites") <> "suiqualquer" and rs("suites") <> "não" and rs("suites") <> "0" and rs("suites") <>"00" and rs("suites") <>"" then
stringSuites = "  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina
 
if  rs("piscina") <> "pisqualquer" and rs("piscina") <> "não" and rs("piscina") <> "0" and rs("piscina") <>"00" and rs("piscina") <>"" then
stringPiscina = "  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria
 
if  rs("portaria") <> "porqualquer" and rs("portaria") <> "não" and rs("portaria") <> "0" and rs("portaria") <>"00" and rs("portaria") <>"" then
stringPortaria = "  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal
 
if  rs("quintal") <> "quiqualquer" and rs("quintal") <> "não" and rs("quintal") <> "0" and rs("quintal") <>"00" and rs("quintal") <>"" then
stringQuintal = "  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras
 
if  rs("quadras") <> "quaqualquer" and rs("quadras") <> "não" and rs("quadras") <> "0" and rs("quadras") <>"00" and rs("quadras") <>"" then
stringQuadras = "  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula
 
if  rs("edicula") <> "ediqualquer" and rs("edicula") <> "não" and rs("edicula") <> "0" and rs("edicula") <>"00" and rs("edicula") <>"" then
stringEdicula = "  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Ocupação--------------------------------------
dim stringOcupacao
 
if  rs("ocupacao") <> "oqualquer" and rs("ocupacao") <> "não informado" and rs("ocupacao") <> ""  then
stringOcupacao = "  and ocupacao ='"&rs("ocupacao")&"'  and ocupacao IS NOT NULL"




else

stringOcupacao = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

stringStandby = "  and (imovel_em_negociacao like  '"&"imóvel OK"&"' ) "





'---------------------------------------------------------------------------

   ' Set rs444 = Server.CreateObject("ADODB.RecordSet")
'se no cliente ou no servidor.

dim strSQL2
	strSQL2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento FROM imoveis"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio&stringAreaConstruida&stringSuites&stringPiscina&stringPortaria&stringQuintal&stringQuadras&stringEdicula&stringOcupacao&stringStandby
	




'strSQL2 ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis"&stringIndex2&stringCidade2&stringBairro22&stringTipo22&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby&" ORDER  BY indexador_indicacoes DESC"
	


'----------------------------------------------------Fim da instrução SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 
  
  
  '------------------incluir em contas acessadas---------------------
  
   dim EnderecoIP2
	 EnderecoIP2 = request.ServerVariables("REMOTE_ADDR")
	 
	 
	 dim JaComprador
	 
	 JaComprador = request.querystring("JaComprador")
	 
	 if JaComprador <> "" then
	'Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP2 &"','"& now() &"')"
	JaComprador = "JaExiste"
     else
	 
	 'JaComprador = "JaExiste"
	 Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data,atendimento) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP2 &"','"& now() &"','"&varAtendimento02&"')"
	
	JaComprador = "JaExiste"
	 end if
  
  
  
  '-----------------------------------------------------------------
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 



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

if   (rs2.eof or rs2.bof) then
response.write "Não foram imóveis para você."
end if

if  not (rs2.eof or rs2.bof) then
%>




  <table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Veja suas 
    indicações abaixo.</font></strong>
	<br>
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
<% varCodimovel = rs2("COD_Imovel") %>
 
 
   
    
  <table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="170"><table width="784" height="190" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td bgcolor="#e9dca8"><table width="774" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td><table width="774" height="180" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="299" height="180" style="border:1px solid #FFFFFF;" bgcolor="#f7ecbf"> 
                                <% If objFSO.FileExists(Server.MapPath(rs2("foto_pequena"))) = True Then%>
                                <a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')"><img src="<%=rs2("foto_pequena")%>" width="299" height="178" border="0"></img></a> 
                                <%else%>
                                <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><strong>Foto não disponível</strong></a></font></div><%end if%></td>
                              <td width="5" height="180">&nbsp;</td>
                              <td width="469" height="180" bgcolor="#f7ecbf"><table width="469" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="469" height="18"><table width="469" border="0" cellspacing="0" cellpadding="0">
                                        <tr> 
                                          <td width="117" height="18" style="border:1px solid #e9dca8;" >
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs2("cidade")%></font></div> 
                                            </td>
                                          <td width="117" height="18" style="border:1px solid #e9dca8;">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs2("bairro")%></font></div> 
                                            </td>
                                          <td width="117" height="18" style="border:1px solid #e9dca8;">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%if rs2("tipo") <> "tqualquer" then response.write rs2("tipo") else response.write "não informado" end if%></font></div> 
                                            </td>
                                          <td width="117" height="18" style="border:1px solid #e9dca8;">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=formatnumber(rs2("valor"),2)%></font></div> 
                                            </td>
                                        </tr>
                                      </table></td>
                                  </tr>
                                  <tr>
                                    <td width="469" height="164">
									
									<% if rs2("qualidade") = "bom negócio" and rs2("imovel_em_negociacao") <> "Vendido pela Veja" then%>
									
									<table width="469" height="164" border="0" cellpadding="0" cellspacing="0">
        <tr>
                                          <td width="379">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="text-decoration:none;color:#9d9249"><strong><%=rs2("obs_imovel")%></strong></a><br>
                                              <br>
                                              Atualizado em:<strong><%=rs2("data_atualizacao")%></strong><br>
                                              <br>
                                              Referência:<strong><%=rs2("cod_imovel")%></strong></font></div>
											  <br>
											  <%
					 
					
					 
					  SqlCompradores001 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores,compradores.valor FROM compradores where telefone like '"&rs2("telefone")&"' or telefone02 like '"&rs2("telefone")&"' or telefone03 like '"&rs2("telefone")&"' ORDER BY cod_compradores ASC" 

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
                                      este im&oacute;vel para comprar outro no 
                                      valor de <%=formatnumber(rsCompradores001("valor"),2)%>, 
                                      clique no texto acima para saber mais</font> 
                                      <br>
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
	  
	  
	  
	  <% elseif rs2("imovel_em_negociacao") = "Vendido pela Veja" then %>
									
		
		<table width="469" height="164" border="0" cellpadding="0" cellspacing="0">
        <tr>
                                          <td width="379">
<div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><strong><%=rs("obs_imovel")%></strong></a><br>
                                      <br>
                                      Atualizado em:<strong><%=rs2("data_atualizacao")%></strong><br>
                                      <br>
                                      Referência:<strong><%=rs2("cod_imovel")%></strong></font></div>
											  
											  <br>
											  
											   <%
					 
					
					 
					  SqlCompradores001 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores,compradores.valor FROM compradores where telefone like '"&rs2("telefone")&"' or telefone02 like '"&rs2("telefone")&"' or telefone03 like '"&rs2("telefone")&"' ORDER BY cod_compradores ASC" 

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
                                      este im&oacute;vel para comprar outro no 
                                      valor de <%=formatnumber(rsCompradores001("valor"),2)%>, 
                                      clique no texto acima para saber mais</font> 
                                      <br>
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
	  
	                                  <div align="center"><font color="#9d9249" face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=varCodimovel%>')" style="text-decoration:none;color:#9d9249"><strong><%=rs2("obs_imovel")%></strong></a><br>
                                        <br>
                                        Atualizado em:<strong><%=rs2("data_atualizacao")%></strong><br>
                                        <br>Referência:<strong><%=rs2("cod_imovel")%></strong></font></div>
		
		
		                      <%
					 
					
					 
					  SqlCompradores001 = "SELECT compradores.telefone,compradores.telefone02,compradores.telefone03,compradores.cod_compradores,compradores.valor FROM compradores where telefone like '"&rs2("telefone")&"' or telefone02 like '"&rs2("telefone")&"' or telefone03 like '"&rs2("telefone")&"' ORDER BY cod_compradores ASC" 

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
                                propriet&aacute;rio est&aacute; vendendo este 
                                im&oacute;vel para comprar outro no valor de <%=formatnumber(rsCompradores001("valor"),2)%>, 
                                clique no texto acima para saber mais</font><br>
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
	  <tr><td height="20"></td></tr>
	 
  
</table>
<% rs2.movenext %>


<%
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
            <a href="?page=<%=intPage - 1%>&varCodCompradores=<%=varCodCompradores%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>&JaComprador=<%=JaComprador%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" size="1" > 
            </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%> 
            <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
            <a href="?page=<%=intPage + 1%>&varCodCompradores=<%=varCodCompradores%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>&JaComprador=<%=JaComprador%>"><b><font color="#000000" face="Verdana, arial" size="1">Próximo</font></b></a><a href="?page=<%=intPage + 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vValor=<%=session("vValor")%>&vTipo=<%=session("vTipo")%>&vNegociacao=<%=session("vNegociacao")%>"> 
            </a> 
            <%End If%>
            </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>

	  
	<%  EscreveFuncaoJavaScript ( Conexao ) %>  
 
<% response.flush%>
  <%response.clear%>




</body>
</html>
