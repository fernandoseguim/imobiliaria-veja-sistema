<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>


<%

dim vPerguntaPermuta

vPerguntaPermuta = request.QueryString("vPerguntaPermuta")


dim vPerguntaCompradores

vPerguntaCompradores = request.QueryString("vPerguntaCompradores")





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

'-----------------------Fim da declaração das variáveis  do imóvel----------------------------------------










'-----------------------Declaração das variáveis de comprador----------------

if (vPerguntaCompradores <> "sim" or  vPerguntaPermuta <> "sim")  then



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


end if


'-----------------Fim da declaração das variáveis dos compradores--------------------------------------------------------











'------------------Request.form de todos os dados do imóvel-----------------


vData_inclusao_vend=now()
vData_ultima_atualizacao_vend=now()
vData_captacao_vend = request.Form("txt_data_captacao_vend")

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
 vSalao_de_jogos_vend= request.form("txt_salao_de_jogos_vend")
 vOBS_salao_de_jogos_vend= request.form("txt_obs_salao_de_jogos_vend")
 vChurrasqueira_vend= request.form("txt_churrasqueira_vend")
 vOBS_churrasqueira_vend= request.form("txt_obs_churrasqueira_vend")
 vPiscina_vend= request.form("txt_piscina_vend")
 vOBS_piscina_vend= request.form("txt_obs_piscina_vend")
 vQuintal_vend= request.form("txt_quintal_vend")
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


'---------------------fim do request.form dos dados do imóvel-------------




'------------------------request.form dos dados do comprador----------------


if (vPerguntaCompradores <> "sim" or  vPerguntaPermuta <> "sim")  then

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

end if









'-------------------------------------------------------------------------------




'---------------pegar os dados de cidade,bairro,vila e fazer uma conexão------------------
dim Conexao

    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	if (vPerguntaCompradores <> "sim" or  vPerguntaPermuta <> "sim")  then
	dim vCidade
	dim vBairro
	
	 vCidade=request.form("combo1")  
    vBairro=request.form("combo2")
	
	end if
	
	
	
	Conexao.Open dsn
	
	
	if (vPerguntaCompradores <> "sim" or  vPerguntaPermuta <> "sim")  then
	
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
 
 end if
 
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
 
 
 
 
 
 
 
 '------------------------------pegar os vários bairros-----------------------
 
 if (vPerguntaCompradores <> "sim" or  vPerguntaPermuta <> "sim")  then
 
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



if (vPerguntaCompradores <> "sim" or  vPerguntaPermuta <> "sim")  then

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
 
 
 
 end if
 
 
 
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
 
 
 '------------------------------------------------------


'------------------------------------incluir no banco de dados--------------

dim vimagem

vimagem = "imovel00000.jpg" 

' Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,ocupacao,captacao,data_atualizacao,vila,placa,condominio,qualidade,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,data_futuro_contato,assunto_futuro_contato,telefone02,telefone03,suites,chaves_do_imovel,melhor_horario_visita,imovel_em_negociacao,metros_de_frente,metros_de_fundo,metros_lateral_esquerda,metros_lateral_direita,origem_captacao,responsavel_cadastramento,data_ultimo_acesso,saldo_devedor,ja_pago_devedor,devendo_devedor,quem_atualizou,data_captacao,conseguiu_proposta,valor_iptu,valor_outros,nome_edificio,edicula,entrada_lateral,piscina,quintal,quadras,andares_edificio,quantidade_elevadores,portaria,salao_de_jogos,salao_de_festas,churrasqueira,OBS_quartos,OBS_vagas,OBS_banheiros,OBS_edicula,OBS_entrada_lateral,OBS_salao_de_festas,OBS_salao_de_jogos,OBS_churrasqueira,OBS_piscina,OBS_quintal,OBS_quadras,OBS_andares_edificio,OBS_quantidade_elevadores,OBS_portaria,obs_suites) values( '"& vProprietario_vend &"','"& vEndereco_vend &"','"& vTelefone_vend &"','"& vEmail_vend &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "icon_foto2.gif" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vA_total_vend &"','"& vA_constr_vend &"','"& vQuartos_vend &"','"& vBanheiros_vend &"','"& vVagas_vend &"','"& vNegociacao_vend &"','"& int(vValor_vend) &"','"& now() &"','"& vOBS_imovel_vend &"','"& vOBS_proprietario_vend &"','"& vPresenca_primeira_vend &"','"& vTitulo_Anuncio_vend &"','"&vTexto_Anuncio_vend &"','"& vOcupacao_vend &"','"& vCaptacao_vend &"','"& now()&"','"& vVila2_vend&"','"& vPlaca_vend &"','"& vCondominio_vend&"','"&vQualidade_vend&"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vData_futuro_contato_vend &"','"& vAssunto_futuro_contato_vend &"','"& vTelefone02_vend &"','"& vTelefone03_vend &"','"& vSuites_vend &"','"& vChaves_do_imovel_vend &"','"& vMelhor_horario_visita_vend &"','"& vImovel_em_negociacao_vend &"','"& vMetros_de_frente_vend &"','"& vMetros_de_fundo_vend &"','"& vMetros_lateral_esquerda_vend &"','"& vMetros_lateral_direita_vend &"','"& vOrigem_captacao_vend &"','"& vResponsavel_cadastramento_vend &"','"& now() &"','"& vSaldo_devedor_vend &"','"& int(vJa_pago_devedor_vend) &"','"& int(vDevendo_devedor_vend) &"','"& session("nome_id") &"','"& vData_captacao_vend &"','"& vConseguiu_proposta_vend &"','"& vValor_iptu_vend &"','"& vValor_outros_vend &"','"& vNome_edificio_vend &"','"& vEdicula_vend &"','"& vEntrada_lateral_vend &"','"& vPiscina_vend &"','"& vQuintal_vend &"','"& vQuadras_vend &"','"& vAndares_edificio_vend &"','"& vQuantidade_elevadores_vend &"','"& vPortaria_vend &"','"& vSalao_de_jogos_vend &"','"& vSalao_de_festas_vend &"','"& vChurrasqueira_vend &"','"& vOBS_quartos_vend &"','"& vOBS_vagas_vend &"','"& vOBS_banheiros_vend &"','"& vOBS_edicula_vend &"','"& vOBS_entrada_lateral_vend &"','"& vOBS_salao_de_festas_vend &"','"& vOBS_salao_de_jogos_vend &"','"& vOBS_churrasqueira_vend &"','"& vOBS_piscina_vend &"','"& vOBS_quintal_vend &"','"& vOBS_quadras_vend &"','"& vOBS_andares_edificio_vend &"','"& vOBS_quantidade_elevadores_vend &"','"& vOBS_portaria_vend &"','"& vOBS_suites_vend &"')"
 



'-----------------------------Selecionar dados do imóvel----------------------------------

dim varCod_imovel

varCod_imovel = request.querystring("varCod_imovel")
dim strSQL

dim rs


Set rs = Server.CreateObject("ADODB.RecordSet")



	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.indexador_indicacoes,imoveis.quem_tirou_foto,imoveis.historico_atual01,imoveis.historico_atual02,imoveis.historico_atual03,imoveis.historico_atual04,imoveis.historico_atual05,imoveis.historico_atual06,imoveis.historico_quem01,imoveis.historico_quem02,imoveis.historico_quem03,imoveis.historico_quem04,imoveis.historico_quem05,imoveis.historico_quem06 FROM imoveis where cod_imovel="&varCod_imovel


RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
		
		

'-----------------------------------------------------------------------------------------





'----------------------------------fim da inclusao de imóvel-------------------------



'---------------------------Formar indexador para as indicações-----------



dim indexador
dim sqlIndexador
dim rsIndexador

set rsIndexador = Server.CreateObject("ADODB.RecordSet")

sqlIndexador = "Select contador.cod_hits,contador.hits from contador where cod_hits = '1'"

rsIndexador.open sqlIndexador, Conexao

indexador = rsIndexador("hits")

indexador = int(indexador) + 1
'-------------------------------

if FormatNumber(rs("valor"),2) <> FormatNumber(vValor_vend,2) then
 indexador = int(indexador)
 else
 indexador = int(rs("indexador_indicacoes"))
 
 end if
 




	 Conexao.execute"update imoveis set cidade='"&vCidade2_vend&"',bairro='"&vBairro2_vend&"',tipo='"& vTipo_vend&"',quartos='"&vQuartos_vend&"',vagas='"&vVagas_vend &"',negociacao='"&vNegociacao_vend&"',valor='"&int(vValor_vend)&"',data_atualizacao='"&now()&"',obs_imovel='"&vOBS_imovel_vend &"',ocupacao='"& vOcupacao_vend&"' where cod_imovel="&varCod_imovel
	





'----------------------------------gerar indicações-----------------------------------

''-------------------------Atualização das indicações-------------
	
	







'---------------------------------------------------------------------






	 
	 response.Redirect "conta_imovel01.asp?varSucesso_imovel="&vProprietario_vend&"&varCod_imovel="&varCod_imovel&""
	



%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

</body>
</html>


<% response.flush%>
	   <%response.clear%>
	   
	   <!--#include file="dsn2.asp"-->