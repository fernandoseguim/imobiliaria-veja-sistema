


<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>



<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vValor,vNegociacao,vFoto
Dim vdata2
Dim vTitulo_anuncio,vTexto_anuncio
dim vVila,vVila2
dim vStandBy
dim vOrigem

vOrigem = request.form("txt_origem")




'----------------------itens do sistema de fideliza��o----------

dim vResponsavel_cadastramento_comprador
dim vData_ultimo_acesso_comprador
dim vData_futuro_contato_comprador
dim vAssunto_futuro_contato_comprador
dim vMelhor_horario_visita_comprador



vResponsavel_cadastramento_comprador = request.form("txt_responsavel_cadastramento_comprador")

vData_futuro_contato_comprador = request.form("txt_data_futuro_contato_comprador")
vAssunto_futuro_contato_comprador = request.form("txt_assunto_futuro_contato_comprador")
vMelhor_horario_visita_comprador = request.form("txt_melhor_horario_visita_comprador")
vData_ultimo_acesso_comprador = now()





dim vTelefone02
dim vTelefone03

vTelefone02 = request.form("txt_telefone02")
vTelefone03 = request.form("txt_telefone03")

dim vData_futuro_contato
dim vAssunto_futuro_contato
dim vSuites
dim vChaves_do_imovel
dim vMelhor_horario_visita
dim vImovel_em_negociacao
dim vMetros_de_frente
dim vMetros_de_fundo
dim vMetros_lateral_esquerda
dim vMetros_lateral_direita
dim vOrigem_captacao
dim vResponsavel_cadastramento
dim vData_ultimo_acesso



 vData_futuro_contato = request.form("txt_data_futuro_contato")
 vAssunto_futuro_contato = request.form("txt_assunto_futuro_contato")
 vSuites = request.form("txt_suites")
 vChaves_do_imovel = request.form("txt_chaves_do_imovel")
 vMelhor_horario_visita = request.form("txt_melhor_horario_visita")
 vImovel_em_negociacao = request.form("txt_imovel_em_negociacao")
 vMetros_de_frente = request.form("txt_metros_de_frente")
 vMetros_de_fundo = request.form("txt_metros_de_fundo")
 vMetros_lateral_esquerda = request.form("txt_metros_lateral_esquerda")
 vMetros_lateral_direita = request.form("txt_metros_lateral_direita")
 vOrigem_captacao = request.form("txt_origem_captacao")
 vResponsavel_cadastramento = request.form("txt_responsavel_cadastramento")
 vData_ultimo_acesso = now()



'------------------------------------------------------------------------------------





dim varCod_imovel

varCod_imovel = request.querystring("varCod_imovel")

vdata2 = now()

vdata = now()

 vStandBy = request.form("txt_standby")
 
 
   
  vProprietario=request.form("txt_proprietario")
 
 
 
  
  
	
	vEmail=request.form("txt_email") 
	
	if vEmail = "" then
	vEmail = "N�o informado"
	end if
	
	 
    vTelefone=request.form("txt_telefone")
	
	if vTelefone = "" then
	vTelefone = "N�o informado"
	end if
	
	   
      
	 
	 
	  
    vCidade=request.form("combo1")  
    vBairro=request.form("combo2")   
     vTipo=request.form("txt_tipo")   
	 
	
    vQuartos=request.form("txt_quartos")   
    
	
	 if vQuartos = "n�o informado" then
 vQuartos = "0"
 end if
	
	 
	 vNegociacao=request.form("example2")
	 
	 if vNegociacao = "nqualquer" then
	 vNegociacao = "qualquer um" 
	 end if
	    
     vValor=request.form("stage22") 
	 
	 if vValor = "vqualquer" then
	 vValor = "0" 
	 end if 
	    
	 
	 dim vDescricao
		
		vDescricao=request.form("txt_descricao")
		
		if vDescricao = "" then
		vDescricao = "n�o informado"
		end if
		
		
		 
	 dim vDescricao_confi
		
		vDescricao_confi=request.form("txt_descricao_confi")
		
		if vDescricao_confi = "" then
		vDescricao_confi = "n�o informado"
		end if
		
		
		
		
		
		dim vAtendimento
		
		vAtendimento=request.form("txt_atendimento")
		
		if vAtendimento = "" then
		vAtendimento = "n�o informado"
		end if
			
			
		dim vCaptacao
		
		vCaptacao=request.form("txt_captacao")
		
		if vCaptacao = "" then
		vCaptacao = "n�o informado"
		end if	
			
			
		
		dim vVagas
		
		vVagas=request.form("txt_vagas")
		
		 
		
		
		
		if vVagas = "" then
		vVagas = "0"
		end if
		
		
		
		if vVagas = "n�o informado" then
 vVagas = "0"
 end if
		
		
	
 
 dim vOcupacao
		
		vOcupacao=request.form("txt_ocupacao")
		
		if vOcupacao = "" then
		vOcupacao = "n�o informado"
		end if						
		






'--------------------------Descri��o do im�vel que vai ser parte do pagamento----------------------------
'----------------------------------------------------------
 
 
 
 
 
 
 dim vEndereco_vend
 dim vPresenca_primeira_vend
 dim vTitulo_Anuncio_vend
 dim vTexto_Anuncio_vend
 dim vPlaca_vend
 dim vCondominio_vend
 Dim vpergunta
 dim vTipo_vend
 dim v_a_total_vend
 dim v_a_constr_vend
 dim vQuartos_vend
 dim vBanheiros_vend
 dim vVagas_vend
 dim vNegociacao_vend
 dim vValor_vend
 dim vStandby_vend
 dim vOcupacao_vend
 dim vQualidade_vend
 dim vOBS_imovel_vend
 dim vOBS_proprietario_vend
 
 
 dim vSaldo_devedor
 dim vJa_pago_devedor
 dim vDevendo_devedor
 
 
 
 
 
 
 vEndereco_vend = request.form("txt_endereco_vend")
 vPresenca_Primeira_vend = request.form("txt_presenca_primeira")

 vTitulo_Anuncio_vend = request.form("txt_titulo_anuncio_vend")
 
 if vTitulo_Anuncio_vend = "" then
 vTitulo_Anuncio_vend = "n�o informado"
 end if
 
 
 
 vTexto_Anuncio_vend = request.form("txt_texto_anuncio_vend")
 
  if vTexto_Anuncio_vend = "" then
 vTexto_Anuncio_vend = "n�o informado"
 end if
 
 
 
 vPlaca_vend = request.form("txt_placa_vend")
 vCondominio_vend = request.form("txt_condominio_vend")
 
 if vCondominio_vend = "" then
 vCondominio_vend = "0"
 end if
 
 	
 vpergunta = request.form("txt_pergunta")
 vTipo_vend = request.form("txt_tipo_vend")
 
 v_a_total_vend = request.form("txt_a_total_vend")
 
 if v_a_total_vend = "" then
 v_a_total_vend = "00"
 end if
 
 
 v_a_constr_vend = request.form("txt_a_constr_vend")
 
 if v_a_constr_vend = "" then
 v_a_constr_vend = "00"
 end if
 
 
 
 vQuartos_vend = request.form("txt_quartos_vend")
 
 if vQuartos_vend = "n�o informado" then
 vQuartos_vend = "0"
 end if
 
 vBanheiros_vend = request.form("txt_banheiros_vend")
 
  if vBanheiros_vend = "n�o informado" then
 vBanheiros_vend = "0"
 end if
 
 
 vVagas_vend = request.form("txt_vagas_vend")

  if vVagas_vend = "n�o informado" then
 vVagas_vend = "0"
 end if
 
 
  vNegociacao_vend = request.form("txt_negociacao_vend")
 vValor_vend = request.form("txt_valor_vend")
 vStandby_vend = request.form("txt_standby_vend")
 vOcupacao_vend = request.form("txt_ocupacao_vend")
 vQualidade_vend = request.form("txt_qualidade_vend")
 vOBS_imovel_vend = request.form("obs_imovel_vend")
 vOBS_proprietario_vend = request.form("obs_proprietario_vend")
 
	
 vSaldo_devedor = request.form("txt_saldo_devedor")
vJa_pago_devedor = request.form("txt_ja_pago_devedor")
vDevendo_devedor = request.form("txt_devendo_devedor")
	
 






 

   
   
	 
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	if vCidade <> "cqualquer" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="& vCidade
 
 
   
         rs2.CursorLocation = 3
         rs2.CursorType = 3

           rs2.ActiveConnection = Conexao
 
 
 
 rs2.open SQL2,Conexao,2,1
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 
'-------------------------------
   rs2.close
  
  
  set rs2 = nothing
 '------------------------------
 
 else
 vCidade2 = "n�o informado"
 end if
 
 
 
 dim vCidade_vend,vCidade2_vend
 
 vCidade_vend = request.form("combo3")
 
 if vCidade_vend <> "cqualquer" then
 
 dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="& vCidade_vend
 
 
     rs22.CursorLocation = 3
        rs22.CursorType = 3

            rs22.ActiveConnection = Conexao
 
 
 
 
 rs22.open SQL22,Conexao,2,1

 vCidade2_vend = rs22("nome_combo1")
 
 '-------------------------------
   rs22.close
  
  
  set rs22 = nothing
 '------------------------------
 
 
 
 else
 vCidade2_vend = "n�o informado"
 end if
 
 
 '-------------------------pegar v�rios bairros--------------------------
 
 
 
 dim vBairro2
 
 
 if vBairro <> "bqualquer" then
 
 
 dim rsMultiBairros
 dim sqlMultiBairros
 
 Set rsMultiBairros = Server.CreateObject("ADODB.RecordSet")
 
 
     rsMultiBairros.CursorLocation = 3
     rsMultiBairros.CursorType = 3

     rsMultiBairros.ActiveConnection = Conexao
 

dim Variavel
dim Retorno
dim i
Variavel = vBairro
Retorno = Split(Variavel,", ")

i=0


for i=0 to UBound(Retorno)



sqlMultiBairros = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& Retorno(i)




rsMultiBairros.open sqlMultiBairros,Conexao,2,1



 while not rsMultiBairros.eof

vBairro2 = rsMultiBairros("nome_combo2")&", "&vBairro2 

rsMultiBairros.MoveNext 
Wend



rsMultiBairros.close








next

'-------------------------------
   
  
  
  set rsMultiBairros = nothing
 '------------------------------
 else
 
 vBairro2 = "n�o informado"
 
 end if
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------
	                                      
dim vBairro_vend,vBairro2_vend
vBairro_vend = request.form("combo4")

 if vBairro_vend <> "bqualquer" then
 dim rs33,SQL33
 Set rs33 = Server.CreateObject("ADODB.RecordSet")
 SQL33 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& vBairro_vend
 
 
   rs33.CursorLocation = 3
        rs33.CursorType = 3

            rs33.ActiveConnection = Conexao
 
 
 rs33.open SQL33,Conexao,2,1
 
 vBairro2_vend = rs33("nome_combo2")
 
 '-------------------------------
   rs33.close
  
  
  set rs33 = nothing
 '------------------------------
 
 
	else
	vBairro2_vend = "n�o informado"
	end if


															  
	
   vVila=request.form("combo5")
	 
 if vVila <> "vlqualquer" then
  dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="&vVila
 
 
   rs333.CursorLocation = 3
        rs333.CursorType = 3

            rs333.ActiveConnection = Conexao
 
 
 
 
 rs333.open SQL333,Conexao,2,1
 
 vVila2 = rs333("nome_combo3")
 
 '-------------------------------
   rs333.close
  
  
  set rs333 = nothing
 '------------------------------
 
 
 else
 vVila2 = "n�o informado"
 end if
 
 
 
 
 
 dim vVila_vend,vVila2_vend
 
  vVila_vend=request.form("combo7")
	 
 if vVila_vend <> "vlqualquer" then
  dim rsVL33,SQLVL33
 Set rsVL33 = Server.CreateObject("ADODB.RecordSet")
 SQLVL33 = "select combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  from combo3 where id_combo3 ="&vVila_vend
 
 
   rsVL33.CursorLocation = 3
        rsVL33.CursorType = 3

            rsVL33.ActiveConnection = Conexao
 
 
 
 
 rsVL33.open SQLVL33,Conexao,2,1
 
 vVila2_vend = rsVL33("nome_combo3")
 
 
 
 '-------------------------------
   rsVL33.close
  
  
  set rsVL33 = nothing
 '------------------------------
 else
 vVila2_vend = "n�o informado"
 end if
 
 
 
 
 
 
 
  dim CompradoresExiste
 dim PermutaExiste
 
 CompradoresExiste = "n�o"
 PermutaExiste = "n�o"
 
 'vamos verificar se o telefone desse im�vel existe em compradores
 
 if  vPergunta = "sim" then
 
 
 
 
 
 
 
 
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,descricao_confi,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,origem,responsavel_cadastramento,data_ultimo_acesso,data_futuro_contato,assunto_futuro_contato,melhor_horario_visita,telefone02,telefone03,quem_atualizou) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vQuartos &"','"& vNegociacao &"','"& int(vValor) &"','"& vdata &"','"& vDescricao &"','"& vDescricao_confi &"','"& vAtendimento &"','"& vData &"','"& vVila2 &"','"& vVagas &"','"& vOcupacao &"','"& vStandBy &"','"& vOrigem &"','"& vResponsavel_cadastramento_comprador &"','"& now() &"','"& vData_futuro_contato_comprador &"','"& vAssunto_futuro_contato_comprador &"','"& vMelhor_horario_visita_comprador &"','"& vTelefone02 &"','"& vTelefone03 &"','"& session("nome_id") &"')"
	
	

	
	
 end if
 
 
 
 
 
   
	 
 
 
 


 
if vPergunta = "sim" then

	dim vimagem
	dim vLink
	
	vimagem = "imovel00000.jpg"
 vLink= "n�o informado"
 
 
 Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp) values( '"& vimagem &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& "n�o informado" &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vOBS_imovel_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vDescricao &"','"& vLink &"','"& vData2 &"','"& vQuartos_vend &"','"& vQuartos &"','"& int(vValor_vend) &"','"& int(vValor) &"','"& vAtendimento &"','"& vData2 &"','"& vVila2_vend &"','"& vVila2 &"','"& vVagas_vend &"','"& vVagas &"')"
 
 
 
 
 
 

	
	
	end if
	
	
	
	'----------------------------Buscar refer�ncia do �ltimo im�vel--------------------
	
	
	
	
	
	
	
'-------------------------------------------------------------------------------	
	
	
	'----------------------------Buscar refer�ncia do �ltimo comprador--------------------
	
	
	
	
	
'-------------------------------------------------------------------------------	
	
	
	
	'----------------------------Buscar refer�ncia do �ltimo comprador--------------------
	
	
	
'-------------------------------------------------------------------------------	
	
	
	'---------------------------Atualizar Refer�ncias-------------
	
	
	
	
	
	'-------------------------Atualiza��o das indica��es-------------
	
	
	dim rs9
	dim strSQL9
	
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	
	
	
	 strSQL9 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where cod_imovel="&varCod_imovel
	   rs9.CursorLocation = 3
        rs9.CursorType = 3

            rs9.ActiveConnection = Conexao
	  
	  
	 rs9.Open strSQL9, Conexao
	
	
	
	if FormatNumber(rs9("valor"),2) <> FormatNumber(vValor_vend,2) then
	
	
	'------------------------Cidade---------------------------
dim stringIndex2

stringIndex2 = " where cod_compradores<>"&"0"&""


dim stringCidade2

if vCidade2_vend <> "qualquer um" and vCidade2_vend <> "n�o informado"  then
stringCidade2 = " and (cidade='"&vCidade2_vend&"' or cidade='"&"n�o informado"&"')"
else
stringCidade2 = " and cidade='"&"n�o informado"&"'"
end if



 '--------------------------Bairro----------------------------

dim stringBairro2


if vBairro2_vend <> "qualquer um" and vBairro2_vend <> "n�o informado" then
stringBairro2 = " and (Bairro like '%"&vBairro2_vend&"%' or Bairro like '%"&"n�o informado"&"%')"
else
stringBairro2 = "and Bairro like '%"&"n�o informado"&"%'"
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

dim stringTipo2


if vTipo_vend <> "qualquer um" and vTipo_vend <> "tqualquer" then
stringTipo2 = " and Tipo='"&vTipo_vend&"'"
else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negocia��o---------------------------

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






dim stringStandby2

stringStandby2 = " and standby = '"&"excluido"&"'"







'---------------------------------------------------------------------------

dim strSQL444Indicacao



dim rs444Indicacao



dim vAssunto_ligar_urgente

vAssunto_ligar_urgente = " Um novo im�vel foi atualizado e ocorreu uma indica��o, ligue imediatamente para esse comprador"
  
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444Indicacao = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby2


         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3

            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	 if not rs444Indicacao.eof then 
				     While NOT rs444Indicacao.EoF 
                   
                    Conexao.execute"update compradores set data_atualizacao='"&now()&"',data_ligar_urgente='"&now()&"', assunto_ligar_urgente='"&vAssunto_ligar_urgente&"' where cod_compradores="&rs444Indicacao("cod_compradores") 
                   
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	'-------------------------------
   rs444Indicacao.close
  
  
  set rs444Indicacao = nothing
 '------------------------------
	
	
	
	
	
end if
	
	
'-------------------------------
   rs9.close
  
  
  set rs9 = nothing
 '------------------------------	
	

	 Conexao.execute"update imoveis set proprietario='"&vProprietario&"',telefone='"&vTelefone&"',email='"&vEmail&"',endereco='"&vEndereco_vend&"',cidade='"&vCidade2_vend&"',bairro='"&vBairro2_vend&"',tipo='"& vTipo_vend&"',area_total='"&v_a_total_vend&"',area_construida='"&v_a_constr_vend&"',quartos='"&vQuartos_vend&"',banheiros='"&vBanheiros_vend&"',vagas='"&vVagas_vend &"',negociacao='"&vNegociacao_vend&"',valor='"&int(vValor_vend)&"',data_atualizacao='"&vdata2&"',obs_imovel='"&vOBS_imovel_vend &"',obs_proprietario='"& vOBS_proprietario_vend &"',presenca_primeira='"&vPresenca_primeira_vend &"',titulo_anuncio='"&vTitulo_Anuncio_vend&"',texto_anuncio='"&vTexto_Anuncio_vend&"',standby='"&vStandby_vend&"',ocupacao='"& vOcupacao_vend&"',vila='"&vVila2_vend&"',qualidade='"&vQualidade_vend&"',captacao='"&vCaptacao&"',placa='"&vPlaca_vend &"',condominio='"&int(vCondominio_vend) &"',data_futuro_contato='"&vData_futuro_contato&"',assunto_futuro_contato='"&vAssunto_futuro_contato&"',telefone02='"&vTelefone02&"',telefone03='"&vTelefone03&"',suites='"&vSuites&"',chaves_do_imovel='"&vChaves_do_imovel&"',melhor_horario_visita='"&vMelhor_horario_visita&"',imovel_em_negociacao='"&vImovel_em_negociacao&"',metros_de_frente='"&vMetros_de_frente&"',metros_de_fundo='"&vMetros_de_fundo&"',metros_lateral_esquerda='"&vMetros_lateral_esquerda&"',metros_lateral_direita='"&vMetros_lateral_direita&"',origem_captacao='"&vOrigem_captacao&"',responsavel_cadastramento='"&vResponsavel_cadastramento&"',saldo_devedor='"&vSaldo_devedor&"',ja_pago_devedor='"&int(vJa_pago_devedor)&"',devendo_devedor='"&int(vDevendo_devedor)&"',quem_atualizou='"& session("nome_id") &"' where cod_imovel="&varCod_imovel
	 	
	
	
	
	
	
	
	
	
	
	 
	 response.Redirect "visualizar_imovel22.asp?varSucesso_imovel="&vProprietario&"&varCod_imovel="&varCod_imovel&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sugest�o inclu�da</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"  ><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
<td width="590" height="105" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    
          <td width="202" height="156"><img src="sorriso_sugestao.jpg" width="202" height="156" border="0"></img></td>   
          <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36"></img></td>

</tr>


</table>







 
 <%
 
     
 
 
           
           Conexao.Close
		   Set Conexao = nothing
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
