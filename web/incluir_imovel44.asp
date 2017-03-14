<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>


<%



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
 vPortaria_vend= request.form("txt_portaria_vend")
 vOBS_portaria_vend= request.form("txt_obs_portaria_vend")
 vOBS_proprietario_vend = request.form("txt_OBS_proprietario_vend")
 vOBS_imovel_vend = request.form("txt_OBS_imovel_vend")
 vQuem_tirou_foto_vend = request.form("txt_quem_tirou_foto_vend")


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










'-------------------------------------------------------------------------------




'---------------pegar os dados de cidade,bairro,vila e fazer uma conexão------------------
dim Conexao

    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	'----------------------Verificar se existe esse cliente e dar origem da------ 
'---------------------captação para a Imobiliária Veja-----------------------




	dim VerificaInclusao
	dim rsVerificaInclusao
	dim SQLVerificaInclusao
	Set rsVerificaInclusao = Server.CreateObject("ADODB.RecordSet")
	
	
	
	
if  vTelefone_vend <>"" then	

SQLVerificaInclusao = "select imoveis.cod_imovel,imoveis.captacao,imoveis.proprietario,imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.endereco  from imoveis where "
do until instr(vTelefone_vend, " ") = 0

		SQLVerificaInclusao = SQLVerificaInclusao & "telefone like '%" _
			& left(vTelefone_vend, instr(vTelefone_vend," ") - 1) & "%' or "
		vTelefone_vend = Right(vTelefone_vend, len(vTelefone_vend) - instr(vTelefone_vend," "))
	loop
	if len(vTelefone_vend) > 1 then
		SQLVerificaInclusao = SQLVerificaInclusao & "telefone like '%" & vTelefone_vend & "%'"
	    SQLVerificaInclusao = SQLVerificaInclusao & " or telefone02 like '%" & vTelefone_vend & "%' or telefone03 like '%" & vTelefone_vend & "%'"
	    
		else
		
		SQLVerificaInclusao = SQLVerificaInclusao & " or telefone02 like '%" & vTelefone_vend & "%' or telefone03 like '%" & vTelefone_vend & "%'"
		SQLVerificaInclusao = left(SQLVerificaInclusao, len(SQLVerificaInclusao) - 4)
	
	
	end if
	
	
	
	 rsVerificaInclusao.open SQLVerificaInclusao,Conexao,2,1
	
	if not rsVerificaInclusao.eof then
	
	vOrigem_captacao_vend = "internet"
	
	end if
	
	else
	
	vTelefone_vend = "00"
	
	end if
	
	
	
	'----------------------Fim de verificação de inclusão---------------------------
	
	
	





	'----------------------Verificar se existe esse cliente e dar origem da------ 
'---------------------captação para a Imobiliária Veja-----------------------




	dim VerificaInclusao02
	dim rsVerificaInclusao02
	dim SQLVerificaInclusao02
	Set rsVerificaInclusao02 = Server.CreateObject("ADODB.RecordSet")
	
	
	
	
if  vTelefone02_vend <>"" then	

SQLVerificaInclusao02 = "select imoveis.cod_imovel,imoveis.captacao,imoveis.proprietario,imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.endereco  from imoveis where "
do until instr(vTelefone02_vend, " ") = 0

		SQLVerificaInclusao02 = SQLVerificaInclusao02 & "telefone like '%" _
			& left(vTelefone02_vend, instr(vTelefone02_vend," ") - 1) & "%' or "
		vTelefone02_vend = Right(vTelefone02_vend, len(vTelefone02_vend) - instr(vTelefone02_vend," "))
	loop
	if len(vTelefone02_vend) > 1 then
		SQLVerificaInclusao02 = SQLVerificaInclusao02 & "telefone like '%" & vTelefone02_vend & "%'"
	    SQLVerificaInclusao02 = SQLVerificaInclusao02 & " or telefone02 like '%" & vTelefone02_vend & "%' or telefone03 like '%" & vTelefone02_vend & "%'"
	    
		else
		
		SQLVerificaInclusao02 = SQLVerificaInclusao02 & " or telefone02 like '%" & vTelefone02_vend & "%' or telefone03 like '%" & vTelefone02_vend & "%'"
		SQLVerificaInclusao02 = left(SQLVerificaInclusao02, len(SQLVerificaInclusao02) - 4)
	
	
	end if
	
	
	
	 rsVerificaInclusao02.open SQLVerificaInclusao02,Conexao,2,1
	
	if not rsVerificaInclusao02.eof then
	
	vOrigem_captacao_vend = "internet"
	
	end if
	
	else
	
	vTelefone02_vend = "00"
	
	end if
	
	
	
	'----------------------Fim de verificação de inclusão---------------------------
	







'----------------------Verificar se existe esse cliente e dar origem da------ 
'---------------------captação para a Imobiliária Veja-----------------------




	dim VerificaInclusao03
	dim rsVerificaInclusao03
	dim SQLVerificaInclusao03
	Set rsVerificaInclusao03 = Server.CreateObject("ADODB.RecordSet")
	
	
	
	
if  vTelefone03_vend <>"" then	

SQLVerificaInclusao03 = "select imoveis.cod_imovel,imoveis.captacao,imoveis.proprietario,imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.endereco  from imoveis where "
do until instr(vTelefone03_vend, " ") = 0

		SQLVerificaInclusao03 = SQLVerificaInclusao03 & "telefone like '%" _
			& left(vTelefone03_vend, instr(vTelefone03_vend," ") - 1) & "%' or "
		vTelefone03_vend = Right(vTelefone03_vend, len(vTelefone03_vend) - instr(vTelefone03_vend," "))
	loop
	if len(vTelefone03_vend) > 1 then
		SQLVerificaInclusao03 = SQLVerificaInclusao03 & "telefone like '%" & vTelefone03_vend & "%'"
	    SQLVerificaInclusao03 = SQLVerificaInclusao03 & " or telefone02 like '%" & vTelefone03_vend & "%' or telefone03 like '%" & vTelefone03_vend & "%'"
	    
		else
		
		SQLVerificaInclusao03 = SQLVerificaInclusao03 & " or telefone02 like '%" & vTelefone03_vend & "%' or telefone03 like '%" & vTelefone03_vend & "%'"
		SQLVerificaInclusao03 = left(SQLVerificaInclusao03, len(SQLVerificaInclusao03) - 4)
	
	
	end if
	
	
	
	 rsVerificaInclusao03.open SQLVerificaInclusao03,Conexao,2,1
	
	if not rsVerificaInclusao03.eof then
	
	vOrigem_captacao_vend = "internet"
	
	end if
	
	else
	
	vTelefone03_vend = "00"
	
	end if
	
	
	
	'----------------------Fim de verificação de inclusão---------------------------
	







	
	
	
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

 Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,ocupacao,captacao,data_atualizacao,vila,placa,condominio,qualidade,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,data_futuro_contato,assunto_futuro_contato,telefone02,telefone03,suites,chaves_do_imovel,melhor_horario_visita,imovel_em_negociacao,metros_de_frente,metros_de_fundo,metros_lateral_esquerda,metros_lateral_direita,origem_captacao,responsavel_cadastramento,data_ultimo_acesso,saldo_devedor,ja_pago_devedor,devendo_devedor,quem_atualizou,data_captacao,conseguiu_proposta,valor_iptu,valor_outros,nome_edificio,edicula,entrada_lateral,piscina,quintal,quadras,andares_edificio,quantidade_elevadores,portaria,salao_de_jogos,salao_de_festas,churrasqueira,OBS_quartos,OBS_vagas,OBS_banheiros,OBS_edicula,OBS_entrada_lateral,OBS_salao_de_festas,OBS_salao_de_jogos,OBS_churrasqueira,OBS_piscina,OBS_quintal,OBS_quadras,OBS_andares_edificio,OBS_quantidade_elevadores,OBS_portaria,obs_suites,indexador_indicacoes,quem_tirou_foto) values( '"& vProprietario_vend &"','"& vEndereco_vend &"','"& vTelefone_vend &"','"& vEmail_vend &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& "icon_foto2.gif" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vA_total_vend &"','"& vA_constr_vend &"','"& vQuartos_vend &"','"& vBanheiros_vend &"','"& vVagas_vend &"','"& vNegociacao_vend &"','"& int(vValor_vend) &"','"& now() &"','"& vOBS_imovel_vend &"','"& vOBS_proprietario_vend &"','"& vPresenca_primeira_vend &"','"& vTitulo_Anuncio_vend &"','"&vTexto_Anuncio_vend &"','"& vOcupacao_vend &"','"& vCaptacao_vend &"','"& now()&"','"& vVila2_vend&"','"& vPlaca_vend &"','"& vCondominio_vend&"','"&vQualidade_vend&"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vimagem &"','"& vData_futuro_contato_vend &"','"& vAssunto_futuro_contato_vend &"','"& vTelefone02_vend &"','"& vTelefone03_vend &"','"& vSuites_vend &"','"& vChaves_do_imovel_vend &"','"& vMelhor_horario_visita_vend &"','"& vImovel_em_negociacao_vend &"','"& vMetros_de_frente_vend &"','"& vMetros_de_fundo_vend &"','"& vMetros_lateral_esquerda_vend &"','"& vMetros_lateral_direita_vend &"','"& vOrigem_captacao_vend &"','"& vResponsavel_cadastramento_vend &"','"& now() &"','"& vSaldo_devedor_vend &"','"& int(vJa_pago_devedor_vend) &"','"& int(vDevendo_devedor_vend) &"','"& session("nome_id") &"','"& vData_captacao_vend &"','"& vConseguiu_proposta_vend &"','"& vValor_iptu_vend &"','"& vValor_outros_vend &"','"& vNome_edificio_vend &"','"& vEdicula_vend &"','"& vEntrada_lateral_vend &"','"& vPiscina_vend &"','"& vQuintal_vend &"','"& vQuadras_vend &"','"& vAndares_edificio_vend &"','"& vQuantidade_elevadores_vend &"','"& vPortaria_vend &"','"& vSalao_de_jogos_vend &"','"& vSalao_de_festas_vend &"','"& vChurrasqueira_vend &"','"& vOBS_quartos_vend &"','"& vOBS_vagas_vend &"','"& vOBS_banheiros_vend &"','"& vOBS_edicula_vend &"','"& vOBS_entrada_lateral_vend &"','"& vOBS_salao_de_festas_vend &"','"& vOBS_salao_de_jogos_vend &"','"& vOBS_churrasqueira_vend &"','"& vOBS_piscina_vend &"','"& vOBS_quintal_vend &"','"& vOBS_quadras_vend &"','"& vOBS_andares_edificio_vend &"','"& vOBS_quantidade_elevadores_vend &"','"& vOBS_portaria_vend &"','"& vOBS_suites_vend &"','"& indexador &"','"& vQuem_tirou_foto_vend &"')"
 

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

vAssunto_ligar_urgente = " Um novo imóvel foi incluído e ocorreu uma indicação, ligue imediatamente para esse comprador."
   
   ' Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

 'strSQL444Indicacao = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2

	'rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	' if not rs444Indicacao.eof then 
				   '  While NOT rs444Indicacao.EoF 
                   
                   ' Conexao.execute"update compradores set data_atualizacao='"&now()&"',data_ligar_urgente='"&now()&"', assunto_ligar_urgente='"&vAssunto_ligar_urgente&"' where cod_compradores="&rs444Indicacao("cod_compradores") 
                   
                 '  rs444Indicacao.MoveNext 
                    ' Wend 
					
					'else
					
					'end if





'------------------------------Fim das indicações---------------------------------------------------------




'--------------------------------se pergunta é igual a sim então cadastrar permuta-------------------


dim vPergunta

vPergunta = request.form("txt_pergunta")

if vPergunta = "sim" then

	
	dim varCodImovel444
	varCodImovel444 = "00"
	
	
		
	
	 
 
 
 
 'Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp) values( '"& vimagem &"','"& vProprietario_vend &"','"& vEmail_vend &"','"& vTelefone_vend &"','"& vEndereco_vend &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vOBS_imovel_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vDescricao &"','"& varCodImovel444 &"','"& "00" &"','"& now() &"','"& vQuartos_vend &"','"& vQuartos &"','"& int(vValor_vend) &"','"& int(vStage22) &"','"& vAtendimento &"','"& now() &"','"& vVila2_vend &"','"& vVila2 &"','"& vVagas_vend &"','"& vVagas &"')"
 

	
	
	end if





'--------------------------------------cadastrar permuta---------------------------------



'----------------------------cadastrar como comprador-------------------------------


if vPergunta = "sim" then
	
	
 
 
	'Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,descricao_confi,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,origem,responsavel_cadastramento,data_ultimo_acesso,data_futuro_contato,assunto_futuro_contato,melhor_horario_visita,telefone02,telefone03,quem_atualizou) values( '"& vProprietario_vend &"','"& vTelefone_vend &"','"& vEmail_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vQuartos &"','"& vExample2 &"','"& int(vStage22) &"','"& now() &"','"& vDescricao &"','"& vDescricao_confi &"','"& vAtendimento &"','"& now() &"','"& vVila2 &"','"& vVagas &"','"& vOcupacao &"','"& vStandBy &"','"& vOrigem &"','"& vResponsavel_cadastramento_comprador &"','"& now() &"','"& vData_futuro_contato_comprador &"','"& vAssunto_futuro_contato_comprador &"','"& vMelhor_horario_visita_comprador &"','"& vTelefone02_vend &"','"& vTelefone03_vend &"','"& session("nome_id") &"')"
	
 
 
 
 end if









'---------------------------------------------------------------------






	 
	 response.Redirect "sucesso_duplicacao01.asp?varSucesso_duplicacao="&vProprietario_vend&""
	



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