<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores03.asp"-->
<% response.buffer=True%>


<%

dim vTarja

if session("permissao") = "2" then

vTarja = "sim"

else
vTarja = ""
 end if
 
 '---------------------------------------------------------------------------------
 
 
 
 
 '---------------------------------TARJA DAS DATAS AUTOMÁTICAS----------------------
 
 
 dim vData01_Tarja02
	dim vData02_Tarja02
	dim vTarja02
	
	vTarja02 ="sim"
	
	
	vData01_Tarja02 = day(now())
	
	vData02_Tarja02 = day(DateAdd("d", 15, now()))
	
	
 
 '-----------------------------------------------------------------------------
 
 
 
 
 
 

'--------------------------Declarar variáveis referentes ao imóvel---------------------


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
dim vProprietario_vend
dim vTelefone_vend
dim vTelefone02_vend
dim vTelefone03_vend
dim vEmail_vend
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
'-----------------------Fim da declaração das variáveis  do imóvel----------------------------------------










'-----------------------Declaração das variáveis de comprador----------------

dim vData_futuro_contato_comprador
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
 vData_ultima_atualizacao_vend= request.form("txt_data_ultima_atualizacao_vend")
 vResponsavel_cadastramento_vend= request.form("txt_responsavel_cadastramento_vend")
 vProprietario_vend= request.form("txt_proprietario_vend")
 vTelefone_vend= request.form("txt_telefone_vend")
 vTelefone02_vend= request.form("txt_telefone02_vend")
 vTelefone03_vend= request.form("txt_telefone03_vend")
 vEmail_vend= request.form("txt_email_vend")
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
 vPortaria_vend= request.form("txt_Portaria_vend")
 vOBS_portaria_vend= request.form("txt_obs_portaria_vend")
 vOBS_proprietario_vend = request.form("txt_OBS_proprietario_vend")
 vOBS_imovel_vend = request.form("txt_OBS_imovel_vend")
 vQuem_tirou_foto_vend = request.form("txt_quem_tirou_foto_vend")
 vOBS_forma_pagamento_vend = request.form("txt_obs_forma_pagamento_vend")
 vRateio_vend = request.form("txt_rateio_vend")
'---------------------fim do request.form dos dados do imóvel-------------




'------------------------request.form dos dados do comprador----------------




 vData_futuro_contato_comprador = request.form("txt_data_futuro_contato_comprador")
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
dim Conexao

    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	dim vCidade
	dim vBairro
	
	 vCidade=request.form("combo1")  
    vBairro=request.form("combo2")
	
	
	Conexao.Open dsn
	
	
	
	
	if vCidade <> "cqualquer" then
	
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
 
 
 
 dim vCidade_vend,vCidade2_vend
 
 vCidade_vend = request.form("combo3")
 
 if vCidade_vend <> "cqualquer" then
 
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
 
 
 
 
 
 
 
 '------------------------------pegar os vários bairros-----------------------
 
 
 
 dim vBairro2
 
 
 if vBairro <> "bqualquer" then
 
 
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
	   
	   
	   
	   
	   
	   
	   
	                                      
dim vBairro_vend,vBairro2_vend
vBairro_vend = request.form("combo4")

 if vBairro_vend <> "bqualquer" then
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


		dim vVila
		dim vVila2													  
	
   vVila=request.form("combo5")
	 
 if vVila <> "vlqualquer" then
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
 
 
 
 
 
 dim vVila_vend,vVila2_vend
 
  vVila_vend=request.form("combo7")
	 
 if vVila_vend <> "vlqualquer" then
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
 
 
 '------------------------------------------------------


'------------------------------------incluir no banco de dados--------------

dim vimagem

vimagem = "imovel00000.jpg" 

vAssunto_ligar_urgente = "Cliente novo para você, via recepção, ligar urgente !!"

Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,descricao_confi,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,origem,responsavel_cadastramento,data_ultimo_acesso,data_futuro_contato,assunto_futuro_contato,melhor_horario_visita,telefone02,telefone03,quem_atualizou,data_ligar_urgente,assunto_ligar_urgente,obs_quartos,obs_vagas,suites,obs_suites,salao_de_festas,obs_salao_de_festas,salao_de_jogos,obs_salao_de_jogos,piscina,obs_piscina,andares_edificio,obs_andares_edificio,edicula,obs_edicula,quintal,obs_quintal,banheiros,obs_banheiros,entrada_lateral,obs_entrada_lateral,churrasqueira,obs_churrasqueira,quadras,obs_quadras,portaria,obs_portaria,quantidade_elevadores,obs_quantidade_elevadores,area_total,area_construida,condominio,condicoes_pagamento,pergunta,clique,tarja02,data01_tarja02,data02_tarja02,data_contato,obs_forma_pagamento,origem_franquia) values( '"& vProprietario_vend &"','"& vTelefone_vend &"','"& vEmail_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vQuartos &"','"& vExample2 &"','"& int(vStage22) &"','"& now() &"','"& vDescricao &"','"& vDescricao_confi &"','"& vAtendimento &"','"& now() &"','"& vVila2 &"','"& vVagas &"','"& vOcupacao &"','"& vStandBy &"','"& vOrigem &"','"& vResponsavel_cadastramento_comprador &"','"& now() &"','"& vData_futuro_contato_comprador &"','"& vAssunto_futuro_contato_comprador &"','"& vMelhor_horario_visita_comprador &"','"& vTelefone03_vend &"','"& vTelefone02_vend &"','"& session("nome_id") &"','"& now() &"','"& vAssunto_ligar_urgente &"','"&vObs_quartos&"','"&vObs_vagas&"','"&vSuites&"','"&vObs_suites&"','"&vSalao_de_festas&"','"&vObs_salao_de_festas&"','"&vSalao_de_jogos&"','"&vObs_salao_de_jogos&"','"&vPiscina&"','"&vObs_piscina&"','"&vAndares_edificio&"','"&vObs_andares_edificio&"','"&vEdicula&"','"&vObs_edicula&"','"&vQuintal&"','"&vObs_quintal&"','"&vBanheiros&"','"&vObs_banheiros&"','"&vEntrada_lateral&"','"&vObs_entrada_lateral&"','"&vChurrasqueira&"','"&vObs_churrasqueira&"','"&vQuadras&"','"&vObs_quadras&"','"&vPortaria&"','"&vObs_portaria&"','"&vQuantidade_elevadores&"','"&vObs_quantidade_elevadores&"','"&int(vArea_total)&"','"&int(vArea_construida)&"','"&int(vCondominio)&"','"&vCondicoes_pagamento&"','"&vPergunta&"','"& vTarja &"','"& vTarja02 &"','"& vData01_Tarja02 &"','"& vData02_Tarja02 &"','"& now() &"','"& vOBS_forma_pagamento_comp &"','"& session("vOrigem_Franquia") &"')"
	
 





dim vPergunta

vPergunta = request.form("txt_pergunta")


if vPergunta = "sim" then


'---------------------------Formar indexador para as indicações-----------
dim indexador
'dim sqlIndexador
'dim rsIndexador

'set rsIndexador = Server.CreateObject("ADODB.RecordSet")

'sqlIndexador = "Select contador.cod_hits,contador.hits from contador where cod_hits = '1'"



'rsIndexador.open sqlIndexador, Conexao

'if rsIndexador.eof then
'indexador = rsIndexador("hits")

'indexador = int(indexador) + 1
'-------------------------------
'else
'indexador = "0"
'end if

indexador = "0"

vData_captacao_vend = now()


 Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,ocupacao,captacao,data_atualizacao,vila,placa,condominio,qualidade,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,data_futuro_contato,assunto_futuro_contato,telefone02,telefone03,suites,chaves_do_imovel,melhor_horario_visita,imovel_em_negociacao,metros_de_frente,metros_de_fundo,metros_lateral_esquerda,metros_lateral_direita,origem_captacao,responsavel_cadastramento,data_ultimo_acesso,saldo_devedor,ja_pago_devedor,devendo_devedor,quem_atualizou,data_captacao,conseguiu_proposta,valor_iptu,valor_outros,nome_edificio,edicula,entrada_lateral,piscina,quintal,quadras,andares_edificio,quantidade_elevadores,portaria,salao_de_jogos,salao_de_festas,churrasqueira,OBS_quartos,OBS_vagas,OBS_banheiros,OBS_edicula,OBS_entrada_lateral,OBS_salao_de_festas,OBS_salao_de_jogos,OBS_churrasqueira,OBS_piscina,OBS_quintal,OBS_quadras,OBS_andares_edificio,OBS_quantidade_elevadores,OBS_portaria,obs_suites,indexador_indicacoes,quem_tirou_foto,cliques_no_imovel,clique,tarja02,data01_tarja02,data02_tarja02,data_contato,obs_forma_pagamento,rateio,origem_franquia) values( '"& vProprietario_vend &"','"& vEndereco_vend &"','"& vTelefone_vend &"','"& vEmail_vend &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "icon_foto2.gif" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& int(vA_total_vend )&"','"& int(vA_constr_vend) &"','"& vQuartos_vend &"','"& vBanheiros_vend &"','"& vVagas_vend &"','"& vNegociacao_vend &"','"& int(vValor_vend) &"','"& now() &"','"& vOBS_imovel_vend &"','"& vOBS_proprietario_vend &"','"& vPresenca_primeira_vend &"','"& vTitulo_Anuncio_vend &"','"&vTexto_Anuncio_vend &"','"& vOcupacao_vend &"','"& vCaptacao_vend &"','"& now()&"','"& vVila2_vend&"','"& vPlaca_vend &"','"& int(vCondominio_vend)&"','"&vQualidade_vend&"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vData_futuro_contato_vend &"','"& vAssunto_futuro_contato_vend &"','"& vTelefone02_vend &"','"& vTelefone03_vend &"','"& vSuites_vend &"','"& vChaves_do_imovel_vend &"','"& vMelhor_horario_visita_vend &"','"& vImovel_em_negociacao_vend &"','"& vMetros_de_frente_vend &"','"& vMetros_de_fundo_vend &"','"& vMetros_lateral_esquerda_vend &"','"& vMetros_lateral_direita_vend &"','"& vOrigem_captacao_vend &"','"& vResponsavel_cadastramento_vend &"','"& now() &"','"& vSaldo_devedor_vend &"','"& int(vJa_pago_devedor_vend) &"','"& int(vDevendo_devedor_vend) &"','"& session("nome_id") &"','"& vData_captacao_vend &"','"& vConseguiu_proposta_vend &"','"& vValor_iptu_vend &"','"& vValor_outros_vend &"','"& vNome_edificio_vend &"','"& vEdicula_vend &"','"& vEntrada_lateral_vend &"','"& vPiscina_vend &"','"& vQuintal_vend &"','"& vQuadras_vend &"','"& vAndares_edificio_vend &"','"& vQuantidade_elevadores_vend &"','"& vPortaria_vend &"','"& vSalao_de_jogos_vend &"','"& vSalao_de_festas_vend &"','"& vChurrasqueira_vend &"','"& vOBS_quartos_vend &"','"& vOBS_vagas_vend &"','"& vOBS_banheiros_vend &"','"& vOBS_edicula_vend &"','"& vOBS_entrada_lateral_vend &"','"& vOBS_salao_de_festas_vend &"','"& vOBS_salao_de_jogos_vend &"','"& vOBS_churrasqueira_vend &"','"& vOBS_piscina_vend &"','"& vOBS_quintal_vend &"','"& vOBS_quadras_vend &"','"& vOBS_andares_edificio_vend &"','"& vOBS_quantidade_elevadores_vend &"','"& vOBS_portaria_vend &"','"& vOBS_suites_vend &"','"& indexador &"','"& vQuem_tirou_foto_vend &"','"& "0" &"','"& vTarja &"','"& vTarja02 &"','"& vData01_Tarja02 &"','"& vData02_Tarja02 &"','"& now() &"','"& vOBS_forma_pagamento_vend &"','"& vRateio_vend &"','"& session("vOrigem_Franquia") &"')"
 


end if
'----------------------------------fim da inclusao de imóvel-------------------------




	' Conexao.execute"update imoveis set proprietario='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"',endereco='"&vEndereco_vend&"',cidade='"&vCidade2_vend&"',bairro='"&vBairro2_vend&"',tipo='"& vTipo_vend&"',area_total='"&va_total_vend&"',area_construida='"&va_constr_vend&"',quartos='"&vQuartos_vend&"',banheiros='"&vBanheiros_vend&"',vagas='"&vVagas_vend &"',negociacao='"&vNegociacao_vend&"',valor='"&int(vValor_vend)&"',data_atualizacao='"&now()&"',obs_imovel='"&vOBS_imovel_vend &"',obs_proprietario='"& vOBS_proprietario_vend &"',presenca_primeira='"&vPresenca_primeira_vend &"',titulo_anuncio='"&vTitulo_Anuncio_vend&"',texto_anuncio='"&vTexto_Anuncio_vend&"',ocupacao='"& vOcupacao_vend&"',vila='"&vVila2_vend&"',qualidade='"&vQualidade_vend&"',captacao='"&vCaptacao_vend&"',placa='"&vPlaca_vend &"',condominio='"&int(vCondominio_vend) &"',data_futuro_contato='"&vData_futuro_contato_vend&"',assunto_futuro_contato='"&vAssunto_futuro_contato_vend&"',telefone02='"&vTelefone02_vend&"',telefone03='"&vTelefone03_vend&"',suites='"&vSuites_vend&"',chaves_do_imovel='"&vChaves_do_imovel_vend&"',melhor_horario_visita='"&vMelhor_horario_visita_vend&"',imovel_em_negociacao='"&vImovel_em_negociacao_vend&"',metros_de_frente='"&vMetros_de_frente_vend&"',metros_de_fundo='"&vMetros_de_fundo_vend&"',metros_lateral_esquerda='"&vMetros_lateral_esquerda_vend&"',metros_lateral_direita='"&vMetros_lateral_direita_vend&"',origem_captacao='"&vOrigem_captacao_vend&"',responsavel_cadastramento='"&vResponsavel_cadastramento_vend&"',saldo_devedor='"&vSaldo_devedor_vend&"',ja_pago_devedor='"&int(vJa_pago_devedor_vend)&"',devendo_devedor='"&int(vDevendo_devedor_vend)&"',quem_atualizou='"& session("nome_id") &"',data_captacao='"& vData_captacao_vend &"',conseguiu_proposta='"& vConseguiu_proposta_vend &"',valor_iptu='"& vValor_iptu_vend &"',valor_outros='"& vValor_outros_vend &"',nome_edificio='"&vNome_edificio_vend &"',edicula='"& vEdicula_vend &"',entrada_lateral='"& vEntrada_lateral_vend &"',piscina='"& vPiscina_vend &"',quintal='"& vQuintal_vend &"',quadras='"& vQuadras_vend &"',andares_edificio='"& vAndares_edificio_vend &"',quantidade_elevadores='"& vQuantidade_elevadores_vend &"',portaria='"& vPortaria_vend &"',salao_de_jogos='"& vSalao_de_jogos_vend &"',salao_de_festas='"& vSalao_de_festas_vend &"',churrasqueira='"& vChurrasqueira_vend &"',obs_quartos='"& vOBS_quartos_vend &"',obs_vagas='"& vOBS_vagas_vend &"',obs_banheiros='"& vOBS_banheiros_vend &"',obs_edicula='"& vOBS_Edicula_vend &"',obs_entrada_lateral='"& vOBS_entrada_lateral_vend &"',obs_salao_de_festas='"& vOBS_salao_de_festas_vend &"',obs_salao_de_jogos='"& vOBS_salao_de_jogos_vend &"',obs_churrasqueira='"& vOBS_churrasqueira_vend &"',obs_piscina='"& vOBS_piscina_vend &"',obs_quintal='"& vOBS_quintal_vend &"',obs_quadras='"& vOBS_quadras_vend &"',obs_andares_edificio='"& vOBS_andares_edificio_vend &"',obs_quantidade_elevadores='"& vOBS_quantidade_elevadores_vend &"',obs_portaria='"& vOBS_portaria_vend &"' where cod_imovel="&3282
	





'----------------------------------gerar indicações-----------------------------------

'-------------------------Atualização das indicações-------------
	
	
	
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

if  vQuartos_vend <> 0 then
stringQuartos2 = " and quartos<="&vQuartos_vend&""
else
stringQuartos2 = " and quartos<="&vQuartos_vend&""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------
dim stringVagas2

if  vVagas_vend <> 0 then
stringVagas2 = " and vagas <="&vVagas_vend&""
else
stringVagas2 = " and vagas <="&vVagas_vend&""
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


stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""















'---------------------------------------------------------------------------

dim strSQL444Indicacao







dim rs444Indicacao

dim vAssunto_ligar_urgente

vAssunto_ligar_urgente = " Um  imóvel foi incluído e ocorreu indicação , ligue imediatamente para o cliente ."
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

 strSQL444Indicacao = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2

	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	 if not rs444Indicacao.eof and vPergunta ="sim" then 
				     While NOT rs444Indicacao.EoF 
                   
                    Conexao.execute"update compradores set data_atualizacao='"&now()&"',data_ligar_urgente='"&now()&"', assunto_ligar_urgente='"&vAssunto_ligar_urgente&"' where cod_compradores="&rs444Indicacao("cod_compradores") 
                   
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if





'------------------------------Fim das indicações---------------------------------------------------------




'--------------------------------se pergunta é igual a sim então cadastrar permuta-------------------




if vPergunta = "sim" then

	
	dim varCodImovel444
	varCodImovel444 = "00"
	
	
		
	
	 
 
 if vValor_vend = "não informado" or vValor_vend = "" then
 vValor_vend = "0"
 end if
 
 
 Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,origem_franquia) values( '"& vimagem &"','"& vProprietario_vend &"','"& vEmail_vend &"','"& vTelefone_vend &"','"& vEndereco_vend &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vOBS_imovel_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vDescricao &"','"& varCodImovel444 &"','"& "00" &"','"& now() &"','"& vQuartos_vend &"','"& vQuartos &"','"& int(vValor_vend) &"','"& int(vStage22) &"','"& vAtendimento &"','"& now() &"','"& vVila2_vend &"','"& vVila2 &"','"& vVagas_vend &"','"& vVagas &"','"& session("vOrigem_Franquia") &"')"
 

	
	
	end if





'--------------------------------------cadastrar permuta---------------------------------



'----------------------------cadastrar como comprador-------------------------------





 
 
 
 
 










'---------------------------------------------------------------------


'-------------------Avisar quando tiver indicação------------------------







'------------------------Cidade---------------------------









stringIndex2 = " where cod_imovel<>"&"0"&""


if vCidade2 <> "qualquer um" and vCidade2 <> "não informado" and vCidade2 <> ""  then
stringCidade2 = " and cidade='"&vCidade2&"'"
else
stringCidade2 = ""
end if

 '--------------------------Bairro----------------------------








if ( vBairro2 <> "qualquer um" and  vBairro2 <> "não informado" and  vBairro2 <> "") then


 
dim Numero_Indicacoes
dim Numero_Indicacoes02




Numero_Indicacoes = 0
Numero_Indicacoes02 = 0


dim soma02
dim soma

soma = 0
soma02 = 0

dim Variavel23
dim Retorno23
dim contar23
Variavel23 =  vBairro2
Retorno23 = Split(vBairro2,", ")

contar23=0

dim stringBairro3
dim stringBairro4
dim stringBairro5

for contar23=0 to UBound(Retorno23)

stringBairro3 = "and ( "
stringBairro4 = " Bairro='"&Retorno23(contar23)&"'or  " &stringBairro4

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

if vTipo <> "qualquer um" and vTipo <> "tqualquer" and vTipo <> ""  then




 
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
VariavelTipo =  vTipo
RetornoTipo = Split(vTipo,", ")

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
if vExample2 = "Compra"  then
vNegocio = "venda"
end if

if vExample2 = "compra" then
vNegocio = "venda"
end if

if vExample2 = "Aluguel" then
vNegocio = "aluguel"
end if

if vExample2 = "aluguel" then
vNegocio = "aluguel"
end if


if  vExample2 <> "qualquer um" and vExample2 <> "" and vExample2 <> "nqualquer" and vExample2 <> "não informado" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  vQuartos <> int(0) and vQuartos <> "" then
stringQuartos2 = " and quartos >="&vQuartos&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if vVagas  <> int(0) and vVagas  <> "" then
stringVagas2 = " and vagas >="&vVagas &""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------------Valor-----------------------------------


   Porcentual = int(vStage22)*10/100
   


   vValorMenor = int(vStage22) - int(Porcentual)
   vValorMaior = int(vStage22) + int(Porcentual)






stringValor2 = "  and Valor <="& vValorMaior &""



'---------------------------------Condominio-----------------------------------



dim stringCondominio
dim Porcentual02


Porcentual02 = int(vCondominio)*10/100
   
dim vCondominioMenor
dim vCondominioMaior

   vCondominioMenor = int(vCondominio) - int(Porcentual02)
   vCondominioMaior = int(vCondominio) + int(Porcentual02)




if  int(vCondominio) <> 0 and vCondominio <> ""  then
stringCondominio = "  and Condominio <="& vCondominioMaior &""
else
stringCondominio = ""
end if


'---------------------------------------------------------------------------


'---------------------------------Área Total-----------------------------------



dim stringAreaTotal
dim Porcentual03

Porcentual03 = int(vArea_total)*10/100
   
dim vAreaTotalMenor
dim vAreaTotalMaior

   vAreaTotalMenor = int(vArea_total) - int(Porcentual03)
   vAreaTotalMaior = int(vArea_total) + int(Porcentual03)



if  int(vArea_total) <> 0 and vArea_total <> "" then
stringAreaTotal = " and area_total >="& vAreaTotalMenor &""
else
stringAreaTotal = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Suítes-----------------------------------------


dim stringSuites
 
if  vSuites <> "suiqualquer" and vSuites <> "não" and vSuites <> "0" and vSuites <>"00" and vSuites <>"" then
stringSuites = "  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina
 
if  vPiscina <> "pisqualquer" and vPiscina <> "não" and vPiscina <> "0" and vPiscina <>"00" and vPiscina <>"" then
stringPiscina = "  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria
 
if  vPortaria <> "porqualquer" and vPortaria <> "não" and vPortaria <> "0" and vPortaria <>"00" and vPortaria <>"" then
stringPortaria = "  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal
 
if  vQuintal <> "quiqualquer" and vQuintal <> "não" and vQuintal <> "0" and vQuintal <>"00" and vQuintal <>"" then
stringQuintal = "  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras
 
if  vQuadras <> "quaqualquer" and vQuadras <> "não" and vQuadras <> "0" and vQuadras <>"00" and vQuadras <>"" then
stringQuadras = "  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula
 
if  vEdicula <> "ediqualquer" and vEdicula <> "não" and vEdicula <> "0" and vEdicula <>"00" and vEdicula <>"" then
stringEdicula = "  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Ocupação--------------------------------------
dim stringOcupacao
 
if  vOcupacao <> "oqualquer" and vOcupacao <> "não informado"  then
stringOcupacao = "  and ocupacao ='"&vOcupacao&"'  and ocupacao IS NOT NULL"




else

stringOcupacao = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

stringStandby = " and ( standby like 'comprador OK') and origem_franquia like '"&session("vOrigem_Franquia")&"' "





'---------------------------------------------------------------------------
dim rs444
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
'se no cliente ou no servidor.

dim strSQL444
	strSQL444 = "SELECT imoveis.cod_imovel FROM imoveis"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio&stringAreaTotal&stringSuites&stringPiscina&stringPortaria&stringQuintal&stringQuadras&stringEdicula&stringOcupacao&stringStandby
	
	
	
	
	dim varIndicacaoCidade 
	 dim varIndicacaoBairro 
	 dim varIndicacaoNegociacao 
	 dim varIndicacaoTipo 
	 dim varIndicacaoQuartos 
	 dim varIndicacaoVagas 
	 dim varIndicacaoValor 
	
	 dim varIndicacaoCondominio 
	 dim varIndicacaoAreaTotal 
	 dim varIndicacaoSuites 
	 dim varIndicacaoPiscina 
	 dim varIndicacaoPortaria 
	 dim varIndicacaoQuintal 
	 dim varIndicacaoQuadras 
	  dim varIndicacaoEdicula 
	  dim varIndicacaoOcupacao 
	
	
	
	
	
	
 varIndicacaoCidade = vCidade2
	 varIndicacaoBairro = vBairro2
	 varIndicacaoNegociacao = vExample2
	 varIndicacaoTipo = vTipo
	 varIndicacaoQuartos = vQuartos
	 varIndicacaoVagas = vVagas
	 varIndicacaoValor = vStage22
	
	 varIndicacaoCondominio = vCondominio
	 varIndicacaoAreaTotal = vArea_total
	 varIndicacaoSuites = vSuites
	 varIndicacaoPiscina = vPiscina
	 varIndicacaoPortaria = vPortaria
	 varIndicacaoQuintal = vQuintal
	 varIndicacaoQuadras = vQuadras
	  varIndicacaoEdicula = vEdicula
	  varIndicacaoOcupacao = vOcupacao
	
	 dim varCodIndicacao
	 
	varCodIndicacao = "'"&strSQL444&"'"
	 
		
Rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
Rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.
 
	 
	 rs444.Open strSQL444,Conexao 
	 
	 
	 
			'-----------------Gerar indicações para permuta--------------   
		   
'------------------------Sua Cidade--------------------------
dim stringIndex202

stringIndex202 = " where cod_permuta<>"&"0"&""
 
 dim stringCidadeVend202
 
  if   vCidade2_vend = "não informado" or vCidade2_vend  = "" or vCidade2_vend  = "cqualquer" or  vCidade2_vend  = "qualquer um" then
	stringCidadeVend202 = ""
 else

stringCidadeVend202 = " and (Cidade_comp='"&vCidade2_vend &"' or Cidade_comp='"&"não informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend202

 if   vBairro2_vend = "não informado" or vBairro2_vend = "" or vBairro2_vend = "bqualquer" or  vBairro2_vend = "qualquer um" then
	stringBairroVend202 = ""
 else
'stringBairroVend = ""
stringBairroVend202 = " and (Bairro_comp like'%"&vBairro2_vend&"%' or Bairro_comp like'%"&"não informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend202

'" and (Vila_comp='"&rs("vila_vend")&"' or Vila_comp='"&"não informado"&"' or Vila_comp='"&"vlqualquer"&"' or Vila_comp='"&"qualquer um"&"' )"

 if   vVila2_vend = "não informado" or vVila2_vend = "" or vVila2_vend = "vlqualquer" or vVila2_vend = "qualquer um" then
	stringVilaVend202 =  ""
 else

stringVilaVend202 = ""

end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend202
 
 
 if vTipo_vend = "não informado" or vTipo_vend = "" or vTipo_vend = "tqualquer" or vTipo_vend = "qualquer um"  then

stringTipoVend202 = ""

else
stringTipoVend202 = " and Tipo_comp like '%"&vTipo_vend&"%'"
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend202
 
 
 

stringQuartosVend202 = " and Quartos_comp <="&int(vQuartos_vend)&""

 


 '-----------------------Número de Vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend202
 
 
 



stringVagasVend202 = " and vagas_comp <="&int(vVagas_vend )&""

 




 
 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
dim PorcentualVend202

dim vValorMenorVend202
dim vValorMaiorVend202

PorcentualVend202 = int(vValor_vend)*20/100

   


   vValorMenorVend202 = int(vValor_vend) - int(PorcentualVend202)
   vValorMaiorVend202 = int(vValor_vend) + int(PorcentualVend202)

 
 
 
 
 
	 dim stringValorVend202
  
	
	
	
	stringValorVend202 = " and Valor_comp >="&  vValorMenorVend202 &" and Valor_comp <="& vValorMaiorVend202&""
 
 
 
 
 
 
 

 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp202
  if vCidade2 ="não informado" or vCidade2 ="" or vCidade2 ="cqualquer" or vCidade2  = "qualquer um" then
	stringCidadeComp202 = ""
	else
	
	stringCidadeComp202 = " and Cidade_vend ='"& vCidade2  &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp202

	if vBairro2 = "não informado" or  vBairro2 = "" or  vBairro2 = "bqualquer" or vBairro2 = "qualquer um" then
	
	
	
	
	
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
Variavel202 = vBairro2
Retorno202 = Split(vBairro2,", ")

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

	if  vVila2 <> "não informado" and  vVila2 <> "" and  vVila2 <> "vlqualquer" and  vVila2 <> "qualquer um" then
	stringVilaComp202 = ""
	else
	
	stringVilaComp202 = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp
  'if rs("tipo_comp")="não informado" or rs("tipo_comp")="" or rs("tipo_comp")="tqualquer" or rs("tipo_comp") = "qualquer um" then
	'stringTipoComp = ""
	'else
	
	
	'stringTipoComp = " and Tipo_vend ='"& rs("tipo_comp")&"'"
	'end if
	
	
	
	'--------------------------Tipo----------------------------

if  vTipo <> "qualquer um" and  vTipo <> "não informado" then




 
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
VariavelTipoComp202 =   vTipo
RetornoTipoComp202 = Split( vTipo,", ")

contarTipoComp202=0

dim stringTipo3Comp202
dim stringTipo4Comp202
dim stringTipo5Comp202
dim stringTipo2Comp202

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
  
	
	stringQuartosComp202 = " and Quartos_vend >="& int(vQuartos) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp202
 
	
	stringVagasComp202 = " and vagas_vend >="& int(vVagas) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------

dim PorcentualComp202

dim vValorMenorComp202
dim vValorMaiorComp202

PorcentualComp202 = int(vStage22)*20/100

   


   vValorMenorComp202 = int(vStage22) - int(PorcentualComp202)
   vValorMaiorComp202 = int(vStage22) + int(PorcentualComp202)


	 dim stringValorComp202
  
	
	
	stringValorComp202 = " and Valor_vend >="& vValorMenorComp202 &" and Valor_vend <="& vValorMaiorComp202 &""
	
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	'strSQL444 = "SELECT * FROM permuta"
	'&stringVilaVend
	'&stringVilaComp
	
	dim varIndicacaoCodigo202
	
	'varIndicacaoCodigo202=rs202("cod_permuta")
	
	dim strSQL444202

	
	strSQL444202 = "SELECT permuta.cod_permuta   FROM permuta"&stringIndex202&stringCidadeVend202&stringBairroVend202&stringVilaVend202&stringTipoVend202&stringQuartosVend202&stringVagasVend202&stringValorVend202&stringCidadeComp202&stringBairroComp202&stringVilaComp202&stringTipo2Comp202&stringQuartosComp202&stringVagasComp202&stringValorComp202
	
	
	
	
	
   dim varIndicacaoCidadeVend202
 dim varIndicacaoBairroVend202
 dim varIndicacaoVilaVend202
 dim varIndicacaoQuartosVend202
 dim varIndicacaoVagasVend202
 dim varIndicacaoValorVend202
 dim varIndicacaoTipoVend202


 dim varIndicacaoCidadeComp202
 dim varIndicacaoBairroComp202
 dim varIndicacaoVilaComp202
 dim varIndicacaoQuartosComp202
 dim varIndicacaoVagasComp202
 dim varIndicacaoValorComp202
 dim varIndicacaoTipoComp202
	
	
	
	
	
	 varIndicacaoCidadeVend202=vCidade2_vend 
 varIndicacaoBairroVend202=vBairro2_vend 
 varIndicacaoVilaVend202=vVila2_vend 
 varIndicacaoQuartosVend202=vQuartos_vend 
 varIndicacaoVagasVend202=vVagas_vend 
 varIndicacaoValorVend202=vValor_vend 
 varIndicacaoTipoVend202=vTipo_vend 


 varIndicacaoCidadeComp202=vCidade2
 varIndicacaoBairroComp202=vBairro2
 varIndicacaoVilaComp202=vVila2
 varIndicacaoQuartosComp202=vQuartos
 varIndicacaoVagasComp202=vVagas
 varIndicacaoValorComp202=vStage22
 varIndicacaoTipoComp202=vTipo
	
	
 dim rs444202
 Set rs444202 = Server.CreateObject("ADODB.RecordSet")	
	
	 
rs444202.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444202.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444202.ActiveConnection = Conexao
	 
	 rs444202.Open strSQL444202,Conexao 
	   
	   
	   
    







	 if (rs444.eof and rs444202.eof) then
	 'response.Redirect "form_incluir_compradores33.asp?varSucesso_imovel="&vProprietario_vend&""
	
     else
	 
	 
	
	
	 end if


%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>inclusão de compradores</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow43652(abrejanela43652) {
   openWindow43652 = window.open(abrejanela43652,'openWin43652','width=610,height=500,resizable=yes,scrollbars=yes,Left=0,Top=0')
   openWindow43652.focus( )
   }

</SCRIPT>




<body bgcolor="<%=escuro%>">
<center>

<%

dim SqlUltimoComprador
dim rsUltimoComprador

sqlUltimoComprador="SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.historico_atual01,compradores.historico_atual02,compradores.historico_atual03,compradores.historico_atual04,compradores.historico_atual05,compradores.historico_atual06,compradores.historico_quem01,compradores.historico_quem02,compradores.historico_quem03,compradores.historico_quem04,compradores.historico_quem05,compradores.historico_quem06,compradores.ocupacao_hist,compradores.endereco_hist,compradores.valor_hist,compradores.quartos_hist,compradores.vagas_hist,compradores.suites_hist,compradores.piscina_hist,compradores.area_total_hist,compradores.area_construida_hist,compradores.edicula_hist,compradores.condominio_hist  FROM compradores order By cod_compradores DESC"
   
Set rsUltimoComprador = Server.CreateObject("ADODB.RecordSet")

	rsUltimoComprador.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsUltimoComprador.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsUltimoComprador.ActiveConnection = Conexao
	
	
	rsUltimoComprador.Open sqlUltimoComprador, Conexao



if not rsUltimoComprador.eof then
'response.write rsUltimoComprador("cod_compradores")
else
 'response.write "0"
end if

%>
<br>
<strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow43652('visualizar_compradores33.asp?varCodCompradores=<%=rsUltimoComprador("cod_compradores")%>')" style="color:#000000 ; text-decoration:none">Visualizar 
o comprador</a></font></strong><br>
<br>
<br>

<%
rsUltimoComprador.close

set rsUltimoComprador = nothing


%>



<br>
<br>
<table width="500" height="120" border="0" cellpadding="0" cellspacing="0">
  <tr>
  
      <td width="225"><table width="250" height="120" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="30" bgcolor="<%=claro%>"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Indica&ccedil;&otilde;es</font></strong></div></td>
          </tr>
          <tr>
            <td bgcolor="<%=medio%>"> 
              <% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6"  then %>
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow22('indicacao_compradores22.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>&varIndicacaoCondominio=<%=varIndicacaoCondominio%>&varIndicacaoAreaTotal=<%=varIndicacaoAreaTotal%>&varIndicacaoSuites=<%=varIndicacaoSuites%>&varIndicacaoPiscina=<%=varIndicacaoPiscina%>&varIndicacaoPortaria=<%=varIndicacaoPortaria%>&varIndicacaoQuintal=<%=varIndicacaoQuintal%>&varIndicacaoQuadras=<%=varIndicacaoQuadras%>&varIndicacaoEdicula=<%=varIndicacaoEdicula%>&varIndicacaoOcupacao=<%=varIndicacaoOcupacao%>')" style="color:#FFFFFF ; text-decoration:none"><%=rs444.RecordCount%></a></strong></font> 
                <%else%>
                <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444.RecordCount%></strong></font> 
                <%end if%>
                <%
	 
 do while not rs444.eof 

 
 
 rs444.movenext
loop
 
 rs444.close
  
 
 









%>
              </div></td>
          </tr>
        </table></td>
      <td width="225"><table width="250" height="120" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="30" bgcolor="<%=claro%>"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Permuta</font></strong></div></td>
          </tr>
          <tr>
            <td bgcolor="<%=medio%>">
		
		 
            <% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %>
        
        <div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow43652('indicacao_permuta22.asp?varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend202%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend202%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend202%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend202%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend202%>&varIndicacaoVagasVend=<%=varIndicacaoVagasVend202%>&varIndicacaoValorVend=<%=varIndicacaoValorVend202%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp202%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp202%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp202%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp202%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp202%>&varIndicacaoVagasComp=<%=varIndicacaoVagasComp202%>&varIndicacaoValorComp=<%=varIndicacaoValorComp202%>&varIndicacaoCodigo=<%=varIndicacaoCodigo202%>')" style="color:#FFFFFF ; text-decoration:none"><strong><%=rs444202.RecordCount%></strong></a></font></div>
          <%else%>
          <div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444202.RecordCount%></strong></font></div>
          <%end if%>
          <%
	 if not rs444202.eof then
 do while not rs444202.eof 

 
 
 rs444202.movenext
loop
 
 rs444202.close
  
 
 
else
%>

<%



end if



%>

		
		
		
		
		
		
		</div> 
             
		    </td>
          </tr>
        </table></td>
  </tr>
</table>

</center>
</body>
</html>


<% response.flush%>
	   <%response.clear%>
	   
	   <!--#include file="dsn2.asp"-->