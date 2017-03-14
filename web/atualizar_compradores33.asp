<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>


<%


dim Conexao
dim rsTarja02



dim vData_futuro_contato_comprador
vData_futuro_contato_comprador = request.form("txt_data_futuro_contato_comprador")

    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Set rsTarja02 = Server.CreateObject("ADODB.RecordSet")
	
	
	Conexao.Open dsn

dim varCodCompradores

varCodCompradores = request.querystring("varCodCompradores")


'--------------Atualizar não contatado-------------------
dim varNaoContatado

varNaoContatado = request.querystring("varNaoContatado")

if varNaoContatado <> "" then


Conexao.execute"update compradores set data_contato='"&now()&"' where cod_compradores="&varCodCompradores
	
 
	 response.Redirect "visualizar_compradores33.asp?varCodCompradores="&varCodCompradores&""
	

end if




'-------------------------------------------------------


dim strSQLTarja02

strSQLTarja02 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02 FROM compradores where cod_compradores="&varCodCompradores
	
	
rsTarja02.CursorLocation = 3
rsTarja02.CursorType = 3

        rsTarja02.Open strSQLTarja02, Conexao

'---------------------------------TARJA DAS DATAS AUTOMÁTICAS----------------------
 
 
 dim vData01_Tarja02
	dim vData02_Tarja02
	dim vTarja02
	
	if UCASE(vData_futuro_contato_comprador) <> UCASE(rsTarja02("data_futuro_contato")) then
	
	vData01_Tarja02 = day(vData_futuro_contato_comprador)
	
	vData02_Tarja02 = day(DateAdd("d", 15, vData_futuro_contato_comprador))
	vTarja02 = "não"
	
	else
	
	vData01_Tarja02 = day(vData_futuro_contato_comprador)
	
	vData02_Tarja02 = day(DateAdd("d", 15, vData_futuro_contato_comprador))
	vTarja02 = "sim"
	end if
	
	
 
 '-----------------------------------------------------------------------------







dim varCod_imovel006

varCod_imovel006 = request.QueryString("varCod_imovel006")


dim varCod_permuta006

varCod_permuta006 = request.QueryString("varCod_permuta006")





dim vPergunta

vPergunta = request.form("txt_pergunta")

if vPergunta = "" then
vPergunta = "não"
end if






dim vPerguntaImovel

vPerguntaImovel = request.querystring("vPerguntaImovel")

dim vPerguntaPermuta  


vPerguntaPermuta = request.querystring("vPerguntaPermuta")


'--------------------------Declarar variáveis referentes ao imóvel---------------------

if ( varCod_imovel006 <> "0")  then

dim vData_futuro_contato_vend
dim vAssunto_futuro_contato_vend
dim  vTipo_vend
dim vNegociacao_vend
dim vData_captacao_vend
dim vOrigem_captacao_vend
dim vCaptacao_vend
dim vAtualizado_por__vend
dim vData_inclusao_vend
dim vData_ultima_atualizacao_vend
dim vResponsavel_cadastramento_vend

dim vChaves_do_imovel_vend
dim vQualidade_vend
dim vMelhor_horario_visita_vend
dim vPlaca_vend
dim vOcupacao_vend
dim vEndereco_vend
dim vValor_vend
dim vCondominio_vend
dim vValor_iptu_vend
dim vValor_outros_vend
dim vSaldo_devedor_vend
dim vJa_pago_devedor_vend
dim vDevendo_devedor_vend
dim vA_total_vend
dim vA_constr_vend
dim vMetros_de_frente_vend
dim vMetros_de_fundo_vend
dim vMetros_lateral_direita_vend
dim vMetros_lateral_esquerda_vend
dim vNome_edificio_vend
dim vPresenca_primeira_vend
dim vTitulo_anuncio_vend
dim vTexto_anuncio_vend
dim vConseguiu_proposta_vend
dim vImovel_em_negociacao_vend
dim vQuartos_vend
dim vOBS_quartos_vend
dim vVagas_vend
dim vOBS_vagas_vend
dim vBanheiros_vend
dim vOBS_banheiros_vend
dim vSuites_vend
dim vOBS_suites_vend
dim vEdicula_vend
dim vOBS_edicula_vend
dim vEntrada_lateral_vend
dim vOBS_entrada_lateral_vend
dim vSalao_de_festas_vend
dim vOBS_salao_de_festas_vend
dim vSalao_de_jogos_vend
dim vOBS_salao_de_jogos_vend
dim vChurrasqueira_vend
dim vOBS_churrasqueira_vend
dim vPiscina_vend
dim vOBS_piscina_vend
dim vQuintal_vend
dim vOBS_quintal_vend
dim vQuadras_vend
dim vOBS_quadras_vend
dim vAndares_edificio_vend
dim vOBS_andares_edificio_vend
dim vQuantidade_elevadores_vend
dim vOBS_quantidade_elevadores_vend
dim vPortaria_vend
dim vOBS_portaria_vend
dim vOBS_proprietario_vend
dim vOBS_imovel_vend
dim vQuem_tirou_foto_vend
dim vOBS_forma_pagamento_vend
dim vRateio_vend


end if
'-----------------------Fim da declaração das variáveis  do imóvel----------------------------------------










'-----------------------Declaração das variáveis de comprador----------------
dim vProprietario_vend
dim vTelefone_vend
dim vTelefone02_vend
dim vTelefone03_vend
dim vEmail_vend





dim vAssunto_futuro_contato_comprador
dim vOrigem
dim vResponsavel_cadastramento_comprador
dim vAtendimento
dim vMelhor_horario_visita_comprador
dim vOcupacao
dim vTipo
dim vQuartos
dim vVagas
dim vExample2
dim vStage22
dim vStandby
dim vDescricao
dim vDescricao_confi


dim vObs_quartos
dim vObs_vagas
dim vSuites
dim vObs_suites
dim vSalao_de_festas
dim vObs_salao_de_festas
dim vSalao_de_jogos
dim vObs_salao_de_jogos
dim vPiscina
dim vObs_piscina
dim vAndares_edificio
dim vObs_andares_edificio
dim vEdicula
dim vObs_edicula
dim vQuintal
dim vObs_quintal
dim vBanheiros
dim vObs_banheiros
dim vEntrada_lateral
dim vObs_entrada_lateral
dim vChurrasqueira
dim vObs_churrasqueira
dim vQuadras
dim vObs_quadras
dim vPortaria
dim vObs_portaria
dim vQuantidade_elevadores
dim vObs_quantidade_elevadores

dim vArea_total
dim vArea_construida
dim vCondominio
dim vCondicoes_pagamento
dim vOBS_forma_pagamento_comp

'-----------------Fim da declaração das variáveis dos compradores--------------------------------------------------------











'------------------Request.form de todos os dados do imóvel-----------------
if ( varCod_imovel006 <> "0")  then

vData_inclusao_vend=now()
vData_ultima_atualizacao_vend=now()
vData_captacao_vend = now()

 vData_futuro_contato_vend = request.form("txt_data_futuro_contato_vend")
 vAssunto_futuro_contato_vend= request.form("txt_assunto_futuro_contato_vend")
 vTipo_vend= request.form("txt_tipo_vend")
 vNegociacao_vend= request.form("txt_negociacao_vend")
 vData_captacao_vend= request.form("txt_data_captacao_vend")
 vOrigem_captacao_vend= request.form("txt_origem_captacao_vend")
 vCaptacao_vend= request.form("txt_captacao_vend")
 vAtualizado_por__vend= request.form("txt_atualizado_por__vend")
 vData_inclusao_vend= request.form("txt_data_inclusao_vend")
 
 vResponsavel_cadastramento_vend= request.form("txt_responsavel_cadastramento_vend")

 vChaves_do_imovel_vend= request.form("txt_chaves_do_imovel_vend")
 vQualidade_vend= request.form("txt_qualidade_vend")
 vMelhor_horario_visita_vend= request.form("txt_melhor_horario_visita_vend")
 vPlaca_vend= request.form("txt_placa_vend")
 vOcupacao_vend= request.form("txt_ocupacao_vend")
 vEndereco_vend= request.form("txt_endereco_vend")
 vValor_vend= request.form("txt_valor_vend")
 vCondominio_vend= request.form("txt_condominio_vend")
 vValor_iptu_vend= request.form("txt_valor_iptu_vend")
 vValor_outros_vend= request.form("txt_valor_outros_vend")
 vSaldo_devedor_vend= request.form("txt_saldo_devedor_vend")
 vJa_pago_devedor_vend= request.form("txt_ja_pago_devedor_vend")
 vDevendo_devedor_vend= request.form("txt_devendo_devedor_vend")
 vA_total_vend= request.form("txt_a_total_vend")
 vA_constr_vend= request.form("txt_a_constr_vend")
 vMetros_de_frente_vend= request.form("txt_metros_de_frente_vend")
 vMetros_de_fundo_vend= request.form("txt_metros_de_fundo_vend")
 vMetros_lateral_direita_vend= request.form("txt_metros_lateral_direita_vend")
 vMetros_lateral_esquerda_vend= request.form("txt_metros_lateral_esquerda_vend")
 vNome_edificio_vend= request.form("txt_nome_edificio_vend")
 vPresenca_primeira_vend= request.form("txt_presenca_primeira_vend")
 vTitulo_anuncio_vend= request.form("txt_titulo_anuncio_vend")
 vTexto_anuncio_vend= request.form("txt_texto_anuncio_vend")
 vConseguiu_proposta_vend= request.form("txt_conseguiu_proposta_vend")
 vImovel_em_negociacao_vend= request.form("txt_imovel_em_negociacao_vend")
 vQuartos_vend= request.form("txt_quartos_vend")
 vOBS_quartos_vend= request.form("txt_obs_quartos_vend")
 vVagas_vend= request.form("txt_vagas_vend")
 vOBS_vagas_vend= request.form("txt_obs_vagas_vend")
 vBanheiros_vend= request.form("txt_banheiros_vend")
 vOBS_banheiros_vend= request.form("txt_obs_banheiros_vend")
 vSuites_vend= request.form("txt_suites_vend")
 vOBS_suites_vend= request.form("txt_obs_suites_vend")
 vEdicula_vend= request.form("txt_edicula_vend")
 vOBS_edicula_vend= request.form("txt_obs_edicula_vend")
 vEntrada_lateral_vend= request.form("txt_entrada_lateral_vend")
 vOBS_entrada_lateral_vend= request.form("txt_obs_entrada_lateral_vend")
 vSalao_de_festas_vend= request.form("txt_salao_de_festas_vend")
 vOBS_salao_de_festas_vend= request.form("txt_obs_salao_de_festas_vend")
 vSalao_de_jogos_vend= request.form("txt_Salao_de_jogos_vend")
 vOBS_salao_de_jogos_vend= request.form("txt_obs_salao_de_jogos_vend")
 vChurrasqueira_vend= request.form("txt_churrasqueira_vend")
 vOBS_churrasqueira_vend= request.form("txt_obs_churrasqueira_vend")
 vPiscina_vend= request.form("txt_piscina_vend")
 vOBS_piscina_vend= request.form("txt_obs_piscina_vend")
 vQuintal_vend= request.form("txt_Quintal_vend")
 vOBS_quintal_vend= request.form("txt_obs_quintal_vend")
 vQuadras_vend= request.form("txt_quadras_vend")
 vOBS_quadras_vend= request.form("txt_obs_quadras_vend")
 vAndares_edificio_vend= request.form("txt_andares_edificio_vend")
 vOBS_andares_edificio_vend= request.form("txt_obs_andares_edificio_vend")
 vQuantidade_elevadores_vend= request.form("txt_quantidade_elevadores_vend")
 vOBS_quantidade_elevadores_vend= request.form("txt_obs_quantidade_elevadores_vend")
 vPortaria_vend= request.form("txt_portaria_vend")
 vOBS_portaria_vend= request.form("txt_obs_portaria_vend")
 vOBS_proprietario_vend = request.form("txt_OBS_proprietario_vend")
 vOBS_imovel_vend = request.form("txt_OBS_imovel_vend")
 vQuem_tirou_foto_vend = request.form("txt_quem_tirou_foto_vend")
 vOBS_forma_pagamento_vend = request.form("txt_obs_forma_pagamento_vend")
 vRateio_vend = request.form("txt_rateio_vend")
end if
'---------------------fim do request.form dos dados do imóvel-------------




'------------------------request.form dos dados do comprador----------------

 vProprietario_vend= request.form("txt_proprietario_vend")
 vTelefone_vend= request.form("txt_telefone_vend")
 vTelefone02_vend= request.form("txt_telefone02_vend")
 vTelefone03_vend= request.form("txt_telefone03_vend")
 vEmail_vend= request.form("txt_email_vend")










 
 vAssunto_futuro_contato_comprador= request.form("txt_assunto_futuro_contato_comprador")
 vOrigem= request.form("txt_origem")
 vResponsavel_cadastramento_comprador= request.form("txt_responsavel_cadastramento_comprador")
 vAtendimento= request.form("txt_atendimento")
 vMelhor_horario_visita_comprador= request.form("txt_melhor_horario_visita_comprador")
 vOcupacao= request.form("txt_ocupacao")
 vTipo= request.form("txt_tipo")
 vQuartos= request.form("txt_quartos")
 vVagas= request.form("txt_vagas")
 vExample2= request.form("example2")
 vStage22= request.form("stage22")
 vStandby= request.form("txt_standby")
 vDescricao= request.form("txt_descricao")
 vDescricao_confi= request.form("txt_descricao_confi")


vObs_quartos = request.form("txt_obs_quartos")
 vObs_vagas = request.form("txt_obs_vagas")
 vSuites = request.form("txt_suites")
 vObs_suites = request.form("txt_obs_suites")
 vSalao_de_festas = request.form("txt_salao_de_festas")
 vObs_salao_de_festas= request.form("txt_obs_salao_de_festas")
 vSalao_de_jogos= request.form("txt_salao_de_jogos")
 vObs_salao_de_jogos= request.form("txt_obs_salao_de_jogos")
 vPiscina= request.form("txt_piscina")
 vObs_piscina= request.form("txt_obs_piscina")
 vAndares_edificio= request.form("txt_andares_edificio")
 vObs_andares_edificio= request.form("txt_obs_andares_edificio")
 vEdicula= request.form("txt_edicula")
 vObs_edicula= request.form("txt_obs_edicula")
 vQuintal= request.form("txt_quintal")
 vObs_quintal= request.form("txt_obs_quintal")
 vBanheiros= request.form("txt_banheiros")
 vObs_banheiros= request.form("txt_obs_banheiros")
 vEntrada_lateral= request.form("txt_entrada_lateral")
 vObs_entrada_lateral= request.form("txt_obs_entrada_lateral")
 vChurrasqueira= request.form("txt_churrasqueira")
 vObs_churrasqueira= request.form("txt_obs_churrasqueira")
 vQuadras= request.form("txt_quadras")
 vObs_quadras= request.form("txt_obs_quadras")
 vPortaria= request.form("txt_portaria")
 vObs_portaria= request.form("txt_obs_portaria")
 vQuantidade_elevadores = request.form("txt_quantidade_elevadores")
 vObs_quantidade_elevadores = request.form("txt_obs_quantidade_elevadores")
 
vArea_total = request.form("txt_area_total")
 vArea_construida = request.form("txt_area_construida")
 vCondominio = request.form("txt_condominio")
 vCondicoes_pagamento = request.form("txt_condicoes_pagamento")
 vOBS_forma_pagamento_comp = request.form("txt_obs_forma_pagamento_comp")






'-------------------------------------------------------------------------------




'---------------pegar os dados de cidade,bairro,vila e fazer uma conexão------------------

	
	
	dim vCidade
	dim vBairro
	
	 vCidade=request.form("combo1")  
    vBairro=request.form("combo2")
	
	
	
	
	
	
	
	if vCidade <> "cqualquer" and vCidade <> "" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 
'---------------------------------------- 
 
 rs2.close
 
 set rs2 = nothing
 
 '--------------------------------------------
 else
 vCidade2 = "não informado"
 end if
 
 
 if ( varCod_imovel006 <> "0")  then
 
 
 dim vCidade_vend,vCidade2_vend
 
 vCidade_vend = request.form("combo3")
 
 if vCidade_vend <> "cqualquer" and vCidade_vend <> "" then
 
 dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 from combo1 where id_combo1 ="& vCidade_vend
 
 rs22.open SQL22,Conexao,2,1

 vCidade2_vend = rs22("nome_combo1")
 
 
 '---------------------------------------- 
 
 rs22.close
 
 set rs22 = nothing
 
 '--------------------------------------------
 
 
 else
 vCidade2_vend = "não informado"
 end if
 
 
 
 end if
 
 
 
 '------------------------------pegar os vários bairros-----------------------
 
 
 
 dim vBairro2
 
 
 if vBairro <> "bqualquer" and vBairro <> "" then
 
 
 dim rsMultiBairros
 dim sqlMultiBairros
 
 Set rsMultiBairros = Server.CreateObject("ADODB.RecordSet")
 

dim Variavel
dim Retorno
dim i
Variavel = vBairro
Retorno = Split(Variavel,", ")

i=0


for i=0 to UBound(Retorno)



sqlMultiBairros = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 from combo2 where id_combo2 ="& Retorno(i)




rsMultiBairros.open sqlMultiBairros,Conexao,2,1



 while not rsMultiBairros.eof

vBairro2 = rsMultiBairros("nome_combo2")&", "&vBairro2 

rsMultiBairros.MoveNext 
Wend



rsMultiBairros.close








next



'---------------------------------------- 
 

 
 set rsMultiBairros = nothing
 
 '--------------------------------------------

 else
 
 vBairro2 = "não informado"
 
 end if
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
' if vBairro <> "bqualquer" then
' dim rs3,SQL3
' Set rs3 = Server.CreateObject("ADODB.RecordSet")
' SQL3 = "select * from combo2 where id_combo2 ="& vBairro
 
 'rs3.open SQL3,Conexao,2,1
' dim vBairro2
' vBairro2 = rs3("nome_combo2")
	'else
	'vBairro2 = "não informado"
	'end if
	   
	   
	   
	   
	   
	   
if ( varCod_imovel006 <> "0")  then	   
	                                      
dim vBairro_vend,vBairro2_vend
vBairro_vend = request.form("combo4")

 if vBairro_vend <> "bqualquer" and vBairro_vend <> "" then
 dim rs33,SQL33
 Set rs33 = Server.CreateObject("ADODB.RecordSet")
 SQL33 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 from combo2 where id_combo2 ="& vBairro_vend
 
 rs33.open SQL33,Conexao,2,1
 
 vBairro2_vend = rs33("nome_combo2")
 
 '---------------------------------------- 
 
 rs33.close
 
 set rs33 = nothing
 
 '--------------------------------------------
 
 
	else
	vBairro2_vend = "não informado"
	end if




end if







		dim vVila
		dim vVila2													  
	
   vVila=request.form("combo5")
	 
 if vVila <> "vlqualquer" and vVila <> "" then
  dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1
 
 vVila2 = rs333("nome_combo3")
 
 '---------------------------------------- 
 
 rs33.close
 
 set rs333 = nothing
 
 '--------------------------------------------
 
 
 
 else
 vVila2 = "não informado"
 end if
 
 
 
 
 
 
 
 
 
 if ( varCod_imovel006 <> "0")  then
 
 
 dim vVila_vend,vVila2_vend
 
  vVila_vend=request.form("combo7")
	 
 if vVila_vend <> "vlqualquer" and vVila_vend <> "" then
  dim rsVL33,SQLVL33
 Set rsVL33 = Server.CreateObject("ADODB.RecordSet")
 SQLVL33 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="&vVila_vend
 
 rsVL33.open SQLVL33,Conexao,2,1
 
 vVila2_vend = rsVL33("nome_combo3")
 
 '---------------------------------------- 
 
 rsVL33.close
 
 set rsVL33 = nothing
 
 '--------------------------------------------
 
 
 else
 vVila2_vend = "não informado"
 end if
 
 
 
 
 
 
 end if
 '------------------------------------------------------


'------------------------------------incluir no banco de dados--------------

dim vimagem

vimagem = "imovel00000.jpg" 



if ( varCod_imovel006 <> "0")  then


'---------------------------Formar indexador para as indicações-----------
dim indexador
dim sqlIndexador
dim rsIndexador

set rsIndexador = Server.CreateObject("ADODB.RecordSet")

sqlIndexador = "Select contador.cod_hits,contador.hits from contador where cod_hits like '1'"

rsIndexador.open sqlIndexador, Conexao

if not rsIndexador.eof then

indexador = rsIndexador("hits")

indexador = int(indexador) + 1
else
indexador = int("0")

end if
'-------------------------------

vData_captacao_vend = now()




'-----------------------------Selecionar dados do imóvel----------------------------------

dim varCod_imovel

varCod_imovel = request.querystring("varCod_imovel")





dim strSQL

dim rs


Set rs = Server.CreateObject("ADODB.RecordSet")




strSQL = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.historico_atual01,compradores.historico_atual02,compradores.historico_atual03,compradores.historico_atual04,compradores.historico_atual05,compradores.historico_atual06,compradores.historico_quem01,compradores.historico_quem02,compradores.historico_quem03,compradores.historico_quem04,compradores.historico_quem05,compradores.historico_quem06,compradores.ocupacao_hist,compradores.endereco_hist,compradores.valor_hist,compradores.quartos_hist,compradores.vagas_hist,compradores.suites_hist,compradores.piscina_hist,compradores.area_total_hist,compradores.area_construida_hist,compradores.edicula_hist,compradores.condominio_hist   FROM compradores where cod_compradores="&varCodcompradores



rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao
	
	
	rs.Open strSQL, Conexao





		
	if vQuartos = "" or vQuartos = "não informado" then
	vQuartos = "0"
	end if
	  
	  if vQuartos_vend = "" or vQuartos_vend = "não informado" then
	vQuartos_vend = "0"
	end if
 
 
    if vVagas = "" or vVagas = "não informado" then
	vVagas = "0"
	end if
	  
	  if vVagas_vend = "" or vVagas_vend = "não informado" then
	vVagas_vend = "0"
	end if
'-------------------------------------------------------------------------------







	'strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria  FROM imoveis where cod_imovel="&varCod_imovel


'RS.CursorLocation = 3
'RS.CursorType = 3

        'rs.Open strSQL, Conexao 
		
		
		
		

'-----------------------------------------------------------------------------------------





'----------------------------------fim da inclusao de imóvel-------------------------




'	 Conexao.execute"update imoveis set proprietario='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"',endereco='"&vEndereco_vend&"',cidade='"&vCidade2_vend&"',bairro='"&vBairro2_vend&"',tipo='"& vTipo_vend&"',area_total='"&va_total_vend&"',area_construida='"&va_constr_vend&"',quartos='"&vQuartos_vend&"',banheiros='"&vBanheiros_vend&"',vagas='"&vVagas_vend &"',negociacao='"&vNegociacao_vend&"',valor='"&int(vValor_vend)&"',data_atualizacao='"&now()&"',obs_imovel='"&vOBS_imovel_vend &"',obs_proprietario='"& vOBS_proprietario_vend &"',presenca_primeira='"&vPresenca_primeira_vend &"',titulo_anuncio='"&vTitulo_Anuncio_vend&"',texto_anuncio='"&vTexto_Anuncio_vend&"',ocupacao='"& vOcupacao_vend&"',vila='"&vVila2_vend&"',qualidade='"&vQualidade_vend&"',captacao='"&vCaptacao_vend&"',placa='"&vPlaca_vend &"',condominio='"&int(vCondominio_vend) &"',data_futuro_contato='"&vData_futuro_contato_vend&"',assunto_futuro_contato='"&vAssunto_futuro_contato_vend&"',telefone02='"&vTelefone02_vend&"',telefone03='"&vTelefone03_vend&"',suites='"&vSuites_vend&"',chaves_do_imovel='"&vChaves_do_imovel_vend&"',melhor_horario_visita='"&vMelhor_horario_visita_vend&"',imovel_em_negociacao='"&vImovel_em_negociacao_vend&"',metros_de_frente='"&vMetros_de_frente_vend&"',metros_de_fundo='"&vMetros_de_fundo_vend&"',metros_lateral_esquerda='"&vMetros_lateral_esquerda_vend&"',metros_lateral_direita='"&vMetros_lateral_direita_vend&"',origem_captacao='"&vOrigem_captacao_vend&"',responsavel_cadastramento='"&vResponsavel_cadastramento_vend&"',saldo_devedor='"&vSaldo_devedor_vend&"',ja_pago_devedor='"&int(vJa_pago_devedor_vend)&"',devendo_devedor='"&int(vDevendo_devedor_vend)&"',quem_atualizou='"& session("nome_id") &"',data_captacao='"& vData_captacao_vend &"',conseguiu_proposta='"& vConseguiu_proposta_vend &"',valor_iptu='"& vValor_iptu_vend &"',valor_outros='"& vValor_outros_vend &"',nome_edificio='"&vNome_edificio_vend &"',edicula='"& vEdicula_vend &"',entrada_lateral='"& vEntrada_lateral_vend &"',piscina='"& vPiscina_vend &"',quintal='"& vQuintal_vend &"',quadras='"& vQuadras_vend &"',andares_edificio='"& vAndares_edificio_vend &"',quantidade_elevadores='"& vQuantidade_elevadores_vend &"',portaria='"& vPortaria_vend &"',salao_de_jogos='"& vSalao_de_jogos_vend &"',salao_de_festas='"& vSalao_de_festas_vend &"',churrasqueira='"& vChurrasqueira_vend &"',obs_quartos='"& vOBS_quartos_vend &"',obs_vagas='"& vOBS_vagas_vend &"',obs_banheiros='"& vOBS_banheiros_vend &"',obs_edicula='"& vOBS_Edicula_vend &"',obs_entrada_lateral='"& vOBS_entrada_lateral_vend &"',obs_salao_de_festas='"& vOBS_salao_de_festas_vend &"',obs_salao_de_jogos='"& vOBS_salao_de_jogos_vend &"',obs_churrasqueira='"& vOBS_churrasqueira_vend &"',obs_piscina='"& vOBS_piscina_vend &"',obs_quintal='"& vOBS_quintal_vend &"',obs_quadras='"& vOBS_quadras_vend &"',obs_andares_edificio='"& vOBS_andares_edificio_vend &"',obs_quantidade_elevadores='"& vOBS_quantidade_elevadores_vend &"',obs_portaria='"& vOBS_portaria_vend &"' where cod_imovel="&varCod_imovel
	


'-----------------------------------------------------------------	
	dim rs444VerificaTarja
	dim strSQL444VerificaTarja
	 Set rs444VerificaTarja = Server.CreateObject("ADODB.RecordSet")
	
	 

 strSQL444VerificaTarja = "SELECT compradores.atendimento,compradores.clique FROM compradores where cod_compradores="&varCodCompradores
	

	rs444VerificaTarja.Open strSQL444VerificaTarja, Conexao
	
	
		if Ucase(rs444VerificaTarja("atendimento")) = UCase(Session("nome_id")) and  (vBairro2 <> "bqualquer" and vBairro2 <> "não informado" and vBairro2 <> "")  then
		 Conexao.execute"update compradores set clique='"&"sim"&"' where cod_compradores="&varCodCompradores
	    end if








'-------------------------------------Inclusão ou atualização de imóveis-------------------





else




end if


'-------------------------------------------------------------------



'----------------------------------gerar indicações-----------------------------------

'-------------------------Atualização das indicações-------------
	
	if ( varCod_imovel006 <> "0")  then
	
	'------------------------Cidade---------------------------
dim stringIndex2

stringIndex2 = " where cod_compradores<>"&"0"&""


dim stringCidade2

if vCidade2_vend <> "qualquer um" and vCidade2_vend <> "não informado"  then
stringCidade2 = " and (cidade='"&vCidade2_vend&"' or cidade='"&"não informado"&"')"
else
stringCidade2 = " and cidade='"&"não informado"&"'"
end if



 '--------------------------Bairro----------------------------

dim stringBairro2


if vBairro2_vend <> "qualquer um" and vBairro2_vend <> "não informado" then
stringBairro2 = " and (Bairro like '%"&vBairro2_vend&"%' or Bairro like '%"&"não informado"&"%')"
else
stringBairro2 = "and Bairro like '%"&"não informado"&"%'"
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

dim stringTipo2


if vTipo_vend <> "qualquer um" and vTipo_vend <> "tqualquer" then
stringTipo2 = " and Tipo like '%"&vTipo_vend&"%'"
else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negociação---------------------------

dim vNegocio
dim stringNegociacao2



vNegocio = "Compra"
if vNegociacao_vend = "venda" then
vNegocio = "compra"
end if

if vNegociacao_vend = "aluguel" then
vNegocio = "aluguel"
end if

if  vNegociacao_vend <> "qualquer um" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------

dim stringQuartos2

if  vQuartos_vend <> 0 and vQuartos_vend <> "" then
stringQuartos2 = " and quartos<="&vQuartos_vend&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------
dim stringVagas2

if  vVagas_vend <> 0 and vVagas_vend <> "" then
stringVagas2 = " and vagas <="&vVagas_vend&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------





'---------------------------------Valor-----------------------------------

dim Porcentual
dim vValorMenor
dim vValorMaior

 
   Porcentual = int(vValor_vend)*10/100
   


   vValorMenor = int(vValor_vend) - int(Porcentual)
   vValorMaior = int(vValor_vend) + int(Porcentual)
  





dim stringValor2

if vValorMenor <> "" and vValorMaior <> "" then

stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

else

stringValor2 = ""
 
 end if













'-----------------------------------Indicações----------------------------------------

if (vPergunta = "sim" and stringVagas2 <> "" and stringQuartos2 <> "") then


dim strSQL444Indicacao


dim rs444Indicacao

dim vAssunto_ligar_urgente


vAssunto_ligar_urgente = " Um  imóvel foi incluído e ocorreu uma indicação, ligue imediatamente para esse comprador."
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

 strSQL444Indicacao = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2



	'rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	' if not rs444Indicacao.eof and vPergunta = "sim" then 
				'     While NOT rs444Indicacao.EoF 
                   
                '    Conexao.execute"update compradores set data_atualizacao='"&now()&"',data_ligar_urgente='"&now()&"', assunto_ligar_urgente='"&vAssunto_ligar_urgente&"' where cod_compradores="&rs444Indicacao("cod_compradores") 
                   
                '   rs444Indicacao.MoveNext 
                    ' Wend 
					
				'	else
					
					'end if



end if


else


end if
'------------------------------Fim das indicações---------------------------------------------------------




'--------------------------------se pergunta é igual a sim então cadastrar permuta-------------------




if ( varCod_imovel006 <> "0")  then

	
	dim varCodImovel444
	varCodImovel444 = "00"
	
	
		
	
	 
 
 
 vimagem = "imovel00000.jpg"
 
 
	
	end if





'--------------------------------------cadastrar permuta---------------------------------



'----------------------------cadastrar como comprador-------------------------------












'---------------------------------------------------------------------





'-------------------------atualizar ou incluir imóvel-----------------------------

dim vCod_imovel02

 vCod_imovel02 = request.form("txt_cod_imovel")

if vPergunta = "sim" and vCod_imovel02 <> "00" and vCod_imovel02 <> "" then


'response.write "teste1"

 Conexao.execute"update imoveis set proprietario='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"',endereco='"&vEndereco_vend&"',cidade='"&vCidade2_vend&"',bairro='"&vBairro2_vend&"',tipo='"& vTipo_vend&"',area_total='"&int(va_total_vend)&"',area_construida='"&int(va_constr_vend)&"',quartos='"&vQuartos_vend&"',banheiros='"&vBanheiros_vend&"',vagas='"&vVagas_vend &"',negociacao='"&vNegociacao_vend&"',valor='"&int(vValor_vend)&"',data_atualizacao='"&now()&"',obs_imovel='"&vOBS_imovel_vend &"',obs_proprietario='"& vOBS_proprietario_vend &"',presenca_primeira='"&vPresenca_primeira_vend &"',titulo_anuncio='"&vTitulo_Anuncio_vend&"',texto_anuncio='"&vTexto_Anuncio_vend&"',ocupacao='"& vOcupacao_vend&"',vila='"&vVila2_vend&"',qualidade='"&vQualidade_vend&"',captacao='"&vCaptacao_vend&"',placa='"&vPlaca_vend &"',condominio='"&int(vCondominio_vend) &"',data_futuro_contato='"&vData_futuro_contato_vend&"',assunto_futuro_contato='"&vAssunto_futuro_contato_vend&"',telefone02='"&vTelefone02_vend&"',telefone03='"&vTelefone03_vend&"',suites='"&vSuites_vend&"',chaves_do_imovel='"&vChaves_do_imovel_vend&"',melhor_horario_visita='"&vMelhor_horario_visita_vend&"',imovel_em_negociacao='"&vImovel_em_negociacao_vend&"',metros_de_frente='"&vMetros_de_frente_vend&"',metros_de_fundo='"&vMetros_de_fundo_vend&"',metros_lateral_esquerda='"&vMetros_lateral_esquerda_vend&"',metros_lateral_direita='"&vMetros_lateral_direita_vend&"',responsavel_cadastramento='"&vResponsavel_cadastramento_vend&"',saldo_devedor='"&vSaldo_devedor_vend&"',ja_pago_devedor='"&int(vJa_pago_devedor_vend)&"',devendo_devedor='"&int(vDevendo_devedor_vend)&"',quem_atualizou='"& session("nome_id") &"',data_captacao='"& vData_captacao_vend &"',conseguiu_proposta='"& vConseguiu_proposta_vend &"',valor_iptu='"& vValor_iptu_vend &"',valor_outros='"& vValor_outros_vend &"',nome_edificio='"&vNome_edificio_vend &"',edicula='"& vEdicula_vend &"',entrada_lateral='"& vEntrada_lateral_vend &"',piscina='"& vPiscina_vend &"',quintal='"& vQuintal_vend &"',quadras='"& vQuadras_vend &"',andares_edificio='"& vAndares_edificio_vend &"',quantidade_elevadores='"& vQuantidade_elevadores_vend &"',portaria='"& vPortaria_vend &"',salao_de_jogos='"& vSalao_de_jogos_vend &"',salao_de_festas='"& vSalao_de_festas_vend &"',churrasqueira='"& vChurrasqueira_vend &"',obs_quartos='"& vOBS_quartos_vend &"',obs_vagas='"& vOBS_vagas_vend &"',obs_banheiros='"& vOBS_banheiros_vend &"',obs_edicula='"& vOBS_Edicula_vend &"',obs_entrada_lateral='"& vOBS_entrada_lateral_vend &"',obs_salao_de_festas='"& vOBS_salao_de_festas_vend &"',obs_salao_de_jogos='"& vOBS_salao_de_jogos_vend &"',obs_churrasqueira='"& vOBS_churrasqueira_vend &"',obs_piscina='"& vOBS_piscina_vend &"',obs_quintal='"& vOBS_quintal_vend &"',obs_quadras='"& vOBS_quadras_vend &"',obs_andares_edificio='"& vOBS_andares_edificio_vend &"',obs_quantidade_elevadores='"& vOBS_quantidade_elevadores_vend &"',obs_portaria='"& vOBS_portaria_vend &"',indexador_indicacoes='"& indexador &"',quem_tirou_foto='"& vQuem_tirou_foto_vend &"',pergunta='"& vPergunta &"',tarja02='"& vTarja02 &"',data01_tarja02='"& vData01_Tarja02 &"',data02_tarja02='"& vData02_Tarja02 &"',data_contato='"& now() &"',obs_forma_pagamento='"& vOBS_forma_pagamento_vend &"',rateio='"& vRateio_vend &"' where cod_imovel="&vCod_imovel02



end if












'Response.write "Até aqui tudo certo..."&varCodCompradores
	 
	 
	 dim vDataAtual022

vDataAtual022 = now()


if vTelefone_vend = "" then

vTelefone_vend = ""

end if




'----------------------------Atualização de histórico--------------------------





'---------------------------------------------------------------------------------









'---------------------atualização de histórico-------------------------------------



dim strSQL99

dim rs99


Set rs99 = Server.CreateObject("ADODB.RecordSet")




strSQL99 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.historico_atual01,compradores.historico_atual02,compradores.historico_atual03,compradores.historico_atual04,compradores.historico_atual05,compradores.historico_atual06,compradores.historico_quem01,compradores.historico_quem02,compradores.historico_quem03,compradores.historico_quem04,compradores.historico_quem05,compradores.historico_quem06,compradores.ocupacao_hist,compradores.endereco_hist,compradores.valor_hist,compradores.quartos_hist,compradores.vagas_hist,compradores.suites_hist,compradores.piscina_hist,compradores.area_total_hist,compradores.area_construida_hist,compradores.edicula_hist,compradores.condominio_hist   FROM compradores where cod_compradores="&varCodcompradores



rs99.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs99.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs99.ActiveConnection = Conexao
	
	
	rs99.Open strSQL99, Conexao




dim vHistorico_quartos



if (vQuartos <> rs99("quartos") ) then

vHistorico_quartos = rs99("quartos")

else

vHistorico_quartos = rs99("quartos_hist")

end if







dim vHistorico_vagas

if vVagas <> rs99("vagas") then

vHistorico_vagas = rs99("vagas")

else

vHistorico_vagas = rs99("vagas_hist")

end if




dim vHistorico_suites

if vSuites <> rs99("suites") then

vHistorico_suites = rs99("suites")

else

vHistorico_suites = rs99("suites_hist")

end if





dim vHistorico_piscina

if vPiscina <> rs99("piscina") then

vHistorico_piscina = rs99("piscina")

else

vHistorico_piscina = rs99("piscina_hist")

end if



dim vHistorico_edicula

if vEdicula <> rs99("edicula") then

vHistorico_edicula = rs99("edicula")

else

vHistorico_edicula = rs99("edicula_hist")

end if



dim vHistorico_area_total

if vArea_total <> rs99("area_total") then

vHistorico_area_total = rs99("area_total")

else

vHistorico_area_total = rs99("area_total_hist")

end if





dim vHistorico_area_construida

if vArea_construida <> rs99("area_construida") then

vHistorico_area_construida = rs99("area_construida")

else

vHistorico_area_construida = rs99("area_construida_hist")

end if




dim vHistorico_ocupacao

if vOcupacao <> rs99("ocupacao") then

vHistorico_ocupacao = rs99("ocupacao")

else

vHistorico_ocupacao = rs99("ocupacao_hist")

end if











dim vHistorico_valor

if vStage22 <> rs99("valor") then

vHistorico_valor = rs99("valor")

else

vHistorico_valor = rs99("valor_hist")

end if




dim vHistorico_condominio

if vCondominio <> rs99("condominio") then

vHistorico_condominio = rs99("condominio")

else

vHistorico_condominio = rs99("condominio_hist")

end if












 if  session("permissao") <> "3" then
	
	
	 
	 
	
	  Conexao.execute"update compradores set nome='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"',cidade='"&vCidade2&"',bairro='"&vbairro2&"',tipo='"&vTipo&"',quartos='"&vQuartos&"',valor='"&int(vStage22)&"',descricao='"&vDescricao&"',descricao_confi='"&vDescricao_confi&"',negociacao='"&vExample2&"',atendimento='"&vAtendimento&"',vila='"&vVila2 &"',vagas='"&vVagas &"',ocupacao='"&vOcupacao &"',standby='"&vStandby &"',origem='"&vOrigem &"',responsavel_cadastramento='"&vResponsavel_cadastramento_comprador &"',data_atualizacao='"& now() &"',data_futuro_contato='"&vData_futuro_contato_comprador &"',assunto_futuro_contato='"&vAssunto_futuro_contato_comprador &"',melhor_horario_visita='"&vMelhor_horario_visita_comprador &"',telefone02='"&vTelefone02_vend &"',telefone03='"&vTelefone03_vend &"',data_ligar_urgente='"&"0/0/2007 00:00:00"&"',assunto_ligar_urgente='"&"Não informado"&"',quem_atualizou='"&session("nome_id")&"',obs_quartos='"&vObs_quartos &"',obs_vagas='"&vObs_vagas &"',suites='"&vSuites &"',obs_suites='"&vObs_suites &"',salao_de_festas='"&vSalao_de_festas &"',obs_salao_de_festas='"&vObs_salao_de_festas &"',salao_de_jogos='"&vSalao_de_jogos &"',obs_salao_de_jogos='"&vObs_salao_de_jogos &"',piscina='"&vPiscina &"',obs_piscina='"&vObs_piscina &"',andares_edificio='"&vAndares_edificio &"',obs_andares_edificio='"&vObs_andares_edificio &"',edicula='"&vEdicula &"',obs_edicula='"&vObs_edicula &"',quintal='"&vQuintal &"',obs_quintal='"&vObs_quintal &"',banheiros='"&vBanheiros &"',obs_banheiros='"&vObs_banheiros &"',entrada_lateral='"&vEntrada_lateral &"',obs_entrada_lateral='"&vObs_entrada_lateral &"',churrasqueira='"&vChurrasqueira &"',obs_churrasqueira='"&vObs_churrasqueira &"',quadras='"&vQuadras &"',obs_quadras='"&vObs_quadras &"',portaria='"&vPortaria &"',obs_portaria='"&vObs_portaria &"',quantidade_elevadores='"&vQuantidade_elevadores &"',obs_quantidade_elevadores='"&vObs_quantidade_elevadores &"' ,area_total='"&vArea_total &"',area_construida='"&vArea_construida &"',condominio='"&int(vCondominio) &"',condicoes_pagamento='"&vCondicoes_pagamento &"',pergunta='"&vPergunta &"',tarja02='"&vTarja02 &"',data01_tarja02='"&vData01_Tarja02 &"',data02_tarja02='"&vData02_Tarja02 &"',data_contato='"&now()&"',obs_forma_pagamento='"&vOBS_forma_pagamento_comp&"',historico_atual01='"& now() &"',historico_atual02='"& rs99("historico_atual01") &"',historico_atual03='"& rs99("historico_atual02") &"',historico_atual04='"& rs99("historico_atual03") &"',historico_atual05='"& rs99("historico_atual04") &"',historico_atual06='"& rs99("historico_atual05") &"',historico_quem01='"& session("nome_id")&"',historico_quem02='"& rs99("historico_quem01") &"',historico_quem03='"& rs99("historico_quem02") &"',historico_quem04='"& rs99("historico_quem03") &"',historico_quem05='"& rs99("historico_quem04") &"',historico_quem06='"& rs99("historico_quem05") &"',ocupacao_hist='"& vHistorico_ocupacao &"',valor_hist='"& vHistorico_valor &"',quartos_hist='"& vHistorico_quartos &"',vagas_hist='"& vHistorico_vagas &"',suites_hist='"& vHistorico_suites &"',piscina_hist='"& vHistorico_piscina &"',area_total_hist='"& vHistorico_area_total &"',area_construida_hist='"& vHistorico_area_construida &"',edicula_hist='"& vHistorico_edicula &"',condominio_hist='"& vHistorico_condominio &"' where cod_compradores="&varCodCompradores
	 
	 
	 
	 else
	
	
	
		
	 Conexao.execute"update compradores set nome='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"',cidade='"&vCidade2&"',bairro='"&vbairro2&"',tipo='"&vTipo&"',quartos='"&vQuartos&"',valor='"&int(vStage22)&"',descricao='"&vDescricao&"',descricao_confi='"&vDescricao_confi&"',negociacao='"&vExample2&"',atendimento='"&vAtendimento&"',data_atualizacao='"&vDataAtual022&"',vila='"&vVila2 &"',vagas='"&vVagas &"',ocupacao='"&vOcupacao &"',standby='"&vStandby &"',origem='"&vOrigem &"',responsavel_cadastramento='"&vResponsavel_cadastramento_comprador &"',data_futuro_contato='"&vData_futuro_contato_comprador &"',assunto_futuro_contato='"&vAssunto_futuro_contato_comprador &"',melhor_horario_visita='"&vMelhor_horario_visita_comprador &"',telefone02='"&vTelefone02_vend &"',telefone03='"&vTelefone03_vend &"',data_ligar_urgente='"&vDataAtual022&"',assunto_ligar_urgente='"&"Cliente novo da internet, transferido para você , ligue imediatamente."&"',quem_atualizou='"&session("nome_id")&"',obs_quartos='"&vObs_quartos &"',obs_vagas='"&vObs_vagas &"',suites='"&vSuites &"',obs_suites='"&vObs_suites &"',salao_de_festas='"&vSalao_de_festas &"',obs_salao_de_festas='"&vObs_salao_de_festas &"',salao_de_jogos='"&vSalao_de_jogos &"',obs_salao_de_jogos='"&vObs_salao_de_jogos &"',piscina='"&vPiscina &"',obs_piscina='"&vObs_piscina &"',andares_edificio='"&vAndares_edificio &"',obs_andares_edificio='"&vObs_andares_edificio &"',edicula='"&vEdicula &"',obs_edicula='"&vObs_edicula &"',quintal='"&vQuintal &"',obs_quintal='"&vObs_quintal &"',banheiros='"&vBanheiros &"',obs_banheiros='"&vObs_banheiros &"',entrada_lateral='"&vEntrada_lateral &"',obs_entrada_lateral='"&vObs_entrada_lateral &"',churrasqueira='"&vChurrasqueira &"',obs_churrasqueira='"&vObs_churrasqueira &"',quadras='"&vQuadras &"',obs_quadras='"&vObs_quadras &"',portaria='"&vPortaria &"',obs_portaria='"&vObs_portaria &"',quantidade_elevadores='"&vQuantidade_elevadores &"',obs_quantidade_elevadores='"&vObs_quantidade_elevadores &"',area_total='"&int(vArea_total) &"',area_construida='"&int(vArea_construida) &"',condominio='"&int(vCondominio) &"',condicoes_pagamento='"&vCondicoes_pagamento &"',pergunta='"&vPergunta &"',tarja02='"&vTarja02 &"',data01_tarja02='"&vData01_Tarja02 &"',data02_tarja02='"&vData02_Tarja02 &"',data_contato='"&now()&"',obs_forma_pagamento='"&vOBS_forma_pagamento_comp&"',historico_atual01='"& now() &"',historico_atual02='"& rs99("historico_atual01") &"',historico_atual03='"& rs99("historico_atual02") &"',historico_atual04='"& rs99("historico_atual03") &"',historico_atual05='"& rs99("historico_atual04") &"',historico_atual06='"& rs99("historico_atual05") &"',historico_quem01='"& session("nome_id")&"',historico_quem02='"& rs99("historico_quem01") &"',historico_quem03='"& rs99("historico_quem02") &"',historico_quem04='"& rs99("historico_quem03") &"',historico_quem05='"& rs99("historico_quem04") &"',historico_quem06='"& rs99("historico_quem05") &"',ocupacao_hist='"& vHistorico_ocupacao &"',valor_hist='"& vHistorico_valor &"',quartos_hist='"& vHistorico_quartos &"',vagas_hist='"& vHistorico_vagas &"',suites_hist='"& vHistorico_suites &"',piscina_hist='"& vHistorico_piscina &"',area_total_hist='"& vHistorico_area_total &"',area_construida_hist='"& vHistorico_area_construida &"',edicula_hist='"& vHistorico_edicula &"',condominio_hist='"& vHistorico_condominio &"' where cod_compradores="&varCodCompradores
	 
	 
	
	
	 end if
	 
	 
	 
	 
	 
	 '-------------cadastrar ou atualizar permuta---------------------

dim vCod_permuta02
 vCod_permuta02 = request.form("txt_cod_permuta")
 
 if vPergunta = "sim" and vCod_permuta02 <> "00" and vCod_permuta02 <> "" then
 
 
 
 
  if vImovel_em_negociacao_vend = "Vendido pela Veja" or vImovel_em_negociacao_vend = "Vendido por outros" or vStandby = "comprador a contatar" then
  
  vStandby = "incluido"
  else
  
  vStandby = "excluido"
  end if 
 
 
 
 
 
  Conexao.execute"update permuta set cod_imovel='"&"0"&"',foto_imovel='"&"imovel00000.jpg"&"',nome='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"',endereco_vend='"&vEndereco_vend&"',link_imovel='"&"não informado"&"',cidade_vend='"&vCidade2_vend&"',cidade_comp='"&vCidade2&"',bairro_vend='"&vbairro2_vend&"',bairro_comp='"&vbairro2&"',tipo_vend='"&vTipo_vend&"',tipo_comp='"&vTipo&"',quartos_vend='"&vQuartos_vend&"',quartos_comp='"&vQuartos&"',valor_vend='"&int(vValor_vend)&"',valor_comp='"&int(vStage22)&"',descricao_vend='"&vOBS_imovel_vend&"',descricao_comp='"&vDescricao&"',atendimento='"&vAtendimento&"',data_atualizacao='"&now()&"',vila_vend='"&vVila2_vend&"',vila_comp='"&vVila2&"',vagas_vend='"&vVagas_vend&"',vagas_comp='"&vVagas&"',standby='"&vStandby&"' where cod_permuta="&vCod_permuta02
	 
	 end if
 
 
 
 
 if vPergunta = "sim" and (vCod_permuta02 = "00" or vCod_permuta02 = "") then
 
 'response.write "teste01"
 
  Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp) values( '"& vimagem &"','"& vProprietario_vend &"','"& vEmail_vend &"','"& vTelefone_vend &"','"& vEndereco_vend &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vOBS_imovel_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vDescricao &"','"& varCodImovel444 &"','"& "00" &"','"& now() &"','"& int(vQuartos_vend) &"','"& int(vQuartos) &"','"& int(vValor_vend) &"','"& int(vStage22) &"','"& vAtendimento &"','"& now() &"','"& vVila2_vend &"','"& vVila2 &"','"& int(vVagas_vend) &"','"& int(vVagas) &"')"
 

	
	end if




if vPergunta = "sim" and (vCod_imovel02 = "00" or vCod_imovel02 = "") then
'response.write "teste01"

 Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,ocupacao,captacao,data_atualizacao,vila,placa,condominio,qualidade,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,data_futuro_contato,assunto_futuro_contato,telefone02,telefone03,suites,chaves_do_imovel,melhor_horario_visita,imovel_em_negociacao,metros_de_frente,metros_de_fundo,metros_lateral_esquerda,metros_lateral_direita,origem_captacao,responsavel_cadastramento,data_ultimo_acesso,saldo_devedor,ja_pago_devedor,devendo_devedor,quem_atualizou,data_captacao,conseguiu_proposta,valor_iptu,valor_outros,nome_edificio,edicula,entrada_lateral,piscina,quintal,quadras,andares_edificio,quantidade_elevadores,portaria,salao_de_jogos,salao_de_festas,churrasqueira,OBS_quartos,OBS_vagas,OBS_banheiros,OBS_edicula,OBS_entrada_lateral,OBS_salao_de_festas,OBS_salao_de_jogos,OBS_churrasqueira,OBS_piscina,OBS_quintal,OBS_quadras,OBS_andares_edificio,OBS_quantidade_elevadores,OBS_portaria,obs_suites,indexador_indicacoes,quem_tirou_foto,cliques_no_imovel,tarja02,data01_tarja02,data02_tarja02,data_contato,obs_forma_pagamento) values( '"& vProprietario_vend &"','"& vEndereco_vend &"','"& vTelefone_vend &"','"& vEmail_vend &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "icon_foto2.gif" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& int(vA_total_vend) &"','"& int(vA_constr_vend) &"','"& vQuartos_vend &"','"& vBanheiros_vend &"','"& vVagas_vend &"','"& vNegociacao_vend &"','"& int(vValor_vend) &"','"& now() &"','"& vOBS_imovel_vend &"','"& vOBS_proprietario_vend &"','"& vPresenca_primeira_vend &"','"& vTitulo_Anuncio_vend &"','"&vTexto_Anuncio_vend &"','"& vOcupacao_vend &"','"& vCaptacao_vend &"','"& now() &"','"& vVila2_vend&"','"& vPlaca_vend &"','"& int(vCondominio_vend)&"','"&vQualidade_vend&"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vData_futuro_contato_vend &"','"& vAssunto_futuro_contato_vend &"','"& vTelefone02_vend &"','"& vTelefone03_vend &"','"& vSuites_vend &"','"& vChaves_do_imovel_vend &"','"& vMelhor_horario_visita_vend &"','"& vImovel_em_negociacao_vend &"','"& vMetros_de_frente_vend &"','"& vMetros_de_fundo_vend &"','"& vMetros_lateral_esquerda_vend &"','"& vMetros_lateral_direita_vend &"','"& vOrigem &"','"& vResponsavel_cadastramento_vend &"','"& now() &"','"& vSaldo_devedor_vend &"','"& int(vJa_pago_devedor_vend) &"','"& int(vDevendo_devedor_vend) &"','"& session("nome_id") &"','"& vData_captacao_vend &"','"& vConseguiu_proposta_vend &"','"& vValor_iptu_vend &"','"& vValor_outros_vend &"','"& vNome_edificio_vend &"','"& vEdicula_vend &"','"& vEntrada_lateral_vend &"','"& vPiscina_vend &"','"& vQuintal_vend &"','"& vQuadras_vend &"','"& vAndares_edificio_vend &"','"& vQuantidade_elevadores_vend &"','"& vPortaria_vend &"','"& vSalao_de_jogos_vend &"','"& vSalao_de_festas_vend &"','"& vChurrasqueira_vend &"','"& vOBS_quartos_vend &"','"& vOBS_vagas_vend &"','"& vOBS_banheiros_vend &"','"& vOBS_edicula_vend &"','"& vOBS_entrada_lateral_vend &"','"& vOBS_salao_de_festas_vend &"','"& vOBS_salao_de_jogos_vend &"','"& vOBS_churrasqueira_vend &"','"& vOBS_piscina_vend &"','"& vOBS_quintal_vend &"','"& vOBS_quadras_vend &"','"& vOBS_andares_edificio_vend &"','"& vOBS_quantidade_elevadores_vend &"','"& vOBS_portaria_vend &"','"& vOBS_suites_vend &"','"& indexador &"','"& vQuem_tirou_foto_vend &"','"& "0" &"','"& vTarja02 &"','"& vData01_Tarja02 &"','"& vData02_Tarja02 &"','"& now() &"','"& vOBS_forma_pagamento_vend &"')"
 else
 
 end if

	
	
	
	
	
	
	
		
	 response.Redirect "visualizar_compradores33.asp?varSucesso_imovel="&vProprietario_vend&"&varCodCompradores="&varCodCompradores&""
	

    ' response.write vPerguntaImovel&"||"&vPerguntaPermuta

%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<%
'response.write vCod_imovel02
%>
</body>
</html>


<% response.flush%>
	   <%response.clear%>
	   
	   <!--#include file="dsn2.asp"-->