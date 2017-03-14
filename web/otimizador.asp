<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Documento sem t&iacute;tulo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<%



'----------------------------------------------------------


Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs333.ActiveConnection = Conexao3
	
	
	rs333.Open Sql333, Conexao3
	
	
	
	'----------------------------------------------------------------





'------------------compradores-------------------





"   compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou    "




'----------------------------------------------------------------



'-------------------------imóveis---------------------------------



" imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou    "





'-----------------------------------------------------------------


'------------------------permuta----------------------------------





"  permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais    "






'------------------------------------------------------------







'-------------------------Senha--------------------------------------



" senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id       "




'--------------------------------------------------------




'------------------------------------Tipo---------------------

" tipo.id_tipo,tipo.tipo,tipo.data_tipo         "




'---------------------------------------------




'----------------------------------combo1--------------


" combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1    "


'--------------------------------------------------------



'-----------------------combo2----------------------------

" combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2      "





'---------------------------------------------------------





'-----------------------------Combo3------------------------


" combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3    "





'---------------------------------------------------------



'---------------------------imóveis procurados -----------------------





"  imoveis_procurados.cod_procurados,imoveis_procurados.cidade,imoveis_procurados.bairro,imoveis_procurados.tipo,imoveis_procurados.negociacao,imoveis_procurados.valor,imoveis_procurados.data,imoveis_procurados.enderecoIP,imoveis_procurados.quartos,imoveis_procurados.vagas,imoveis_procurados.nome,imoveis_procurados.telefone,imoveis_procurados.email      "





'-------------------------------------------------------------

'------------------------------referência procurados-----------------


"    referencia_procurados.cod_ReferenciaProcurados,referencia_procurados.referencia,referencia_procurados.enderecoIP,referencia_procurados.data    "



'----------------------------------------------------------------------



'-----------------------------compradores procurados---------------------


"        compradores_procurados.cod_CompradoresProcurados,compradores_procurados.cidade,compradores_procurados.bairro,compradores_procurados.tipo,compradores_procurados.negociacao,compradores_procurados.valor,compradores_procurados.enderecoIP,compradores_procurados.data,compradores_procurados.quartos,compradores_procurados.vagas,compradores_procurados.nome,compradores_procurados.telefone       "




'---------------------------------------------------------------------




'-----------------------------contas procuradas-------------------------



   "     contas_procuradas.cod_conta,contas_procuradas.cod_conta,contas_procuradas.nome,contas_procuradas.telefone,contas_procuradas.codigo_conta,contas_procuradas.tipo_conta,contas_procuradas.endereco_ip,contas_procuradas.data     "




'-----------------------------------------------------------------------






'------------------------------email--------------------------

" email.cod_email,email.nome,email.email,email.assunto,email.mensagem,email.data,email.cod_imovel,email.telefone   "



'---------------------------------------------------------


'---------------------------proposta-------------------


" proposta.Cod_proposta,proposta.proposta_proposta,proposta.foto_proposta,proposta.foto_proposta,proposta.nome_proposta,proposta.telefone_proposta,proposta.email_proposta,proposta.data_proposta,proposta.horario_proposta,proposta.interesse_proposta,proposta.cod_imovel_proposta    "





'-------------------------------------------------------



%>
</body>
</html>
