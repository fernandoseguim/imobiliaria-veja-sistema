<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_imoveis.asp"-->
<%response.Buffer = true %>

<%


if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "S?o Bernardo"
end if


dim Conexao3


Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn



'----------------Pegar os dados do cliente--------------------------------

session("nome") = request.querystring("txt_nome")

if session("nome") = "" then


session("nome") = request.querystring("nome")


end if





session("telefone") = request.querystring("txt_telefone")

if session("telefone") = "" then


session("telefone") = request.querystring("telefone")


end if





session("email") = request.querystring("txt_email")

if session("email") = "" then


session("email") = request.querystring("email")


end if









'----------------Fazer listagem de cidade atual-------------
dim rs333
dim Sql333


Set rs333 = Server.CreateObject("ADODB.RecordSet")
'Abrindo a tabela MARCAS!
Sql333 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



 rs333.CursorLocation = 3
rs333.CursorType = 3

rs333.ActiveConnection = Conexao3



        rs333.Open sql333, Conexao3




'------------------------------------------------------------





'------------------------Listagem do tipo atual -------------------

dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 



 rs444Tipo23.CursorLocation = 3
rs444Tipo23.CursorType = 3

rs444Tipo23.ActiveConnection = Conexao3



       
	 
	 
	 
	 
	 
	 
	 rs444Tipo23.Open strSQL444Tipo23, Conexao3






'--------------------------------------------------------------------























'------------------------Listagem da cidade pretendida------------

dim rs555
dim Sql555



Set rs555 = Server.CreateObject("ADODB.RecordSet")

Sql555 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


 rs555.CursorLocation = 3
rs555.CursorType = 3

rs555.ActiveConnection = Conexao3



        rs555.Open Sql555, Conexao3




'-------------------------------------------------------------------






'--------------------------Listagem do tipo pretendido-------------
dim rs444Tipo24,strSQL444Tipo24
   
    Set rs444Tipo24 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo24 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	



 rs444Tipo24.CursorLocation = 3
rs444Tipo24.CursorType = 3

rs444Tipo24.ActiveConnection = Conexao3



        
	
	
	
	
	
	
	 rs444Tipo24.Open strSQL444Tipo24, Conexao3




'-------------------------------------------------------------------



%>







































<%

dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2
dim vValor_vend,vValor_vend1,vValor_vend2
dim vValor_comp,vValor_comp1,vValor_comp2
dim vCidade_vend,vCidade_comp
dim stringIndex
 
  vValor_vend=request.querystring("txt_valor_vend")
  
  session("vValor_vend") = vValor_vend
  
  session("vValor_vend1")=left(vValor_vend,10)
   session("vValor_vend2")=right(vValor_vend,10)
 
  
  
  vValor_comp=request.querystring("txt_valor_comp")
  
  if vValor_comp = "vqualquer" then
  vValor_comp = "0000000000 0000000000"
  end if
  
  
  session("vValor_comp") = vValor_comp
  
   session("vValor_comp1")=left(vValor_comp,10)
   session("vValor_comp2")=right(vValor_comp,10)
   
   dim vNome,vTelefone
   
   vNome = request.querystring("txt_nome")
   vTelefone = request.querystring("txt_telefone")
   
  
 
  
  
 
 '---------------------------Buscar Cidades-------------------------------------
 

 
 vCidade_vend2 = request.querystring("combo1")
 
session("vCidade_vend2") = vCidade_vend2
  
   
   if session("vCidade_vend2") = "" then
session("vCidade_vend2") = request.querystring("vCidade_vend2")
end if
   
   
    
	
	
	
	
	 

	  
	
	if session("vCidade_vend2") = "" then
	session("vCidade_vend2") = request.QueryString("vCidade_vend2")
	end if
	
	
	if session("vCidade_vend2") <> "cqualquer" and session("vCidade_vend2") <> "" then
	
	dim rs222,SQL222
 Set rs222 = Server.CreateObject("ADODB.RecordSet")
 SQL222 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade_vend2")
 
 rs222.open SQL222,Conexao3,2,1
 
 vCidade_vend = rs222("nome_combo1")
 
 rs222.close
 
 set rs222 = nothing
 
 else
 vCidade_vend = vCidade_vend2
 end if

	session("vCidade_vend")= vCidade_vend
	
	if session("vCidade_vend") = "" then
	session("vCidade_vend") = request.querystring("vCidade_vend")
	end if
	
	
	
	dim vBairro_vend2
	 vBairro_vend2=request.Querystring("combo2")
	 session("vBairro_vend2") = vBairro_vend2
	 if session("vBairro_vend2") = "" then
session("vBairro_vend2") = request.querystring("vBairro_vend2")

end if
	 
	 if session("vBairro_vend2") <> "bqualquer" and session("vBairro_vend2") <> ""  then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& session("vBairro_vend2")
 
 rs3.open SQL3,Conexao3,2,1

 vBairro_vend = rs3("nome_combo2")
 
 rs3.close
 
 set rs3 = nothing
 
 
 
 
 else
 vBairro_vend = vBairro_vend2
	end if                                      
									
	 
	 
	 
	 session("vBairro_vend")= vBairro_vend
	 
	 if session("vBairro_vend") = "" then
	session("vBairro_vend") = request.querystring("vBairro_vend")
	end if
	 
	 
	 
	 
	
 '----------------------------Cidade e bairro comp-------------------------
 
 
 
 
  
  dim vCidade_comp2
 
   
   vCidade_comp2=request.querystring("combo3")
   
   session("vCidade_comp2") = vCidade_comp2
   
   if session("vCidade_comp2") = "" then
session("vCidade_comp2") = request.querystring("vCidade_comp2")
end if
   
   
    
	
	
	
	
	 

	 
	
	if session("vCidade_comp2") <> "cqualquer" and session("vCidade_comp2") <> ""  then
	
	dim rs22,SQL22
 Set rs22 = Server.CreateObject("ADODB.RecordSet")
 SQL22 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade_comp2")
 
 rs22.open SQL22,Conexao3,2,1
 
 vCidade_comp = rs22("nome_combo1")
 
 rs22.close
 
 set rs22 = nothing
 
 else
 vCidade_comp = vCidade_comp2
 end if

	session("vCidade_comp")= vCidade_comp
	
	if session("vCidade_comp") = "" then
	session("vCidade_comp") = request.querystring("vCidade_comp")
	end if
	
	
	dim vBairro_comp2
	 vBairro_comp2=request.Querystring("combo4")
	 session("vBairro_comp2") = vBairro_comp2
	 if session("vBairro_comp2") = "" then
session("vBairro_comp2") = request.querystring("vBairro_comp2")

end if
	 
	 if session("vBairro_comp2") <> "bqualquer" and session("vBairro_comp2") <> ""  then
	  dim rs33,SQL33
 Set rs33 = Server.CreateObject("ADODB.RecordSet")
 SQL33 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   from combo2 where id_combo2 ="& session("vBairro_comp2")
 
 rs33.open SQL33,Conexao3,2,1

 vBairro_comp = rs33("nome_combo2")
 
 rs33.close
 
 set rs33 = nothing
 
 else
 vBairro_comp = vBairro_comp2
	end if                                      
									
	 
	 
	 
	 session("vBairro_comp")= vBairro_comp
	 
	 
	 if session("vBairro_comp") = "" then
	session("vBairro_comp") = request.querystring("vBairro_comp")
	end if
	 
	
	 
	 

 '-----------------------------Buscando os tipos de imóveis---------------------
 
 
 
 dim vTipo_vend,vTipo_comp
 
 vTipo_vend = request.Querystring("txt_Tipo_vend")
 session("vTipo_vend") = vTipo_vend
 
 if session("vTipo_vend") = "" then
 
 session("vTipo_vend") = request.querystring("vTipo_vend")
 
 end if
 
 
 
 
 vTipo_comp = request.querystring("txt_Tipo_comp")
 session("vTipo_comp") = vTipo_comp
 
 if session("vTipo_comp") = "" then
 
 session("vTipo_comp") = request.querystring("vTipo_comp")
 
 end if
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 '------------------------------N?meros de quartos--------------------------------

 
 
 dim vQuartos_vend,vQuartos_comp
 
 vQuartos_vend = request.querystring("txt_Quartos_vend")
 session("vQuartos_vend") = vQuartos_vend
 
 if session("vQuartos_vend") = "" then
 
 session("vQuartos_vend") = request.querystring("vQuartos_vend")
 
 end if
 
 
 
 
 vQuartos_comp = request.querystring("txt_Quartos_comp")
 session("vQuartos_comp") = vQuartos_comp
 
 if session("vQuartos_comp") = "" then
 
 session("vQuartos_comp") = request.querystring("vQuartos_comp")
 
 end if
 
 
 
 
 '------------------------------N?meros de Vagas--------------------------------

 
 
 dim vVagas_vend,vVagas_comp
 
 vVagas_vend = request.querystring("txt_vagas_vend")
 session("vVagas_vend") = vVagas_vend
 
 if session("vVagas_vend") = "" then
 
 session("vVagas_vend") = request.querystring("vVagas_vend")
 
 end if
 
 
 
 
 vVagas_comp = request.querystring("txt_vagas_comp")
 session("vVagas_comp") = vVagas_comp
 
 if session("vVagas_comp") = "" then
 
 session("vVagas_comp") = request.querystring("vVagas_comp")
 
 end if
 
 
 
 
 
 
 
 
 
 
 
 '------------------------Sua Cidade--------------------------

stringIndex = " where cod_permuta<>"&"0"&"" 

if  session("vCidade_vend") <> "cqualquer" and session("vCidade_vend") <> "qualquer um" and session("vCidade_vend") <> "n?o informado" and session("vCidade_vend") <> "" then
stringCidadeVend = " and (cidade_comp='"& session("vCidade_vend")&"' or cidade_comp='"& "n?o informado" &"' or cidade_comp='"& "cqualquer" &"' or cidade_comp='"&"qualquer um"&"')"
else
stringCidadeVend = ""
end if
 
 
 
 	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend

 if   session("vBairro_vend") <> "bqualquer"  and session("vBairro_vend") <> "" and session("vBairro_vend") <> "n?o informado" and session("vBairro_vend") <> "qualquer um" then
	stringBairroVend = " and (Bairro_comp like '%"&session("vBairro_vend")&"%' or Bairro_comp like '%"&"n?o informado"&"%' or Bairro_comp like'%"&"bqualquer"&"%' or Bairro_comp like'%"&"qualquer um"&"%')"
 else

stringBairroVend = ""

end if







 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend
 
 
 if session("vTipo_vend") <> "tqualquer" and session("vTipo_vend") <> "" then

stringTipoVend = " and Tipo_comp like '%"&session("vTipo_vend")&"%'"

else
stringTipoVend = ""
 
 end if


 
 '-----------------------N?mero de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend
 
 
 if session("vQuartos_vend") <> "qqualquer" and session("vQuartos_vend") <> "" then

stringQuartosVend = " and Quartos_comp <="&session("vQuartos_vend")&""
else
stringQuartosVend = ""
 end if
 


 
 
 '-----------------------N?mero de Vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend
 
 
 if session("vVagas_vend") <> "vgqualquer" and session("vVagas_vend") <> "" then

stringVagasVend = " and vagas_comp <="&session("vVagas_vend")&""
else

stringVagasVend = ""
 end if
 


 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
	 dim stringValorVend
	
	
	 if session("vValor_vend") = "" then
	 session("vValor_vend")= request.QueryString("vValor_vend")
	 end if
	 
	  
	   if session("vValor_vend1") = "" then
	 session("vValor_vend1")= request.QueryString("vValor_vend1")
	 end if
	 
	  if session("vValor_vend2") = "" then
	 session("vValor_vend2")= request.QueryString("vValor_vend2")
	 end if
	 
  if session("vValor_vend")<>"vqualquer" and session("vValor_vend")<>"" then
	stringValorVend = " and Valor_comp >="& session("vValor_vend1") &" and Valor_comp <="& session("vValor_vend2") &""
	else	
	stringValorVend = ""
  end if
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp
  if session("vCidade_comp")<>"cqualquer" and session("vCidade_comp")<>"cqualquer" and session("vCidade_comp")<>"n?o informado" and session("vCidade_comp")<>"qualquer um" and session("vCidade_comp")<>"" then
	stringCidadeComp = " and Cidade_vend ='"& session("vCidade_comp") &"'"
	else
	
	stringCidadeComp = ""
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp

	if session("vBairro_comp") <> "bqualquer" and session("vBairro_comp") <> "bqualquer" and session("vBairro_comp") <> "n?o informado" and session("vBairro_comp") <> "qualquer um" and session("vBairro_comp") <> "" then
	stringBairroComp = " and Bairro_vend ='"& session("vBairro_comp") &"'"
	else
	
	stringBairroComp = ""
	end if
	
	
	
	 
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	 dim stringTipoComp
  if session("vTipo_comp")<>"tqualquer" and session("vTipo_comp")<>"" then
	stringTipoComp = " and Tipo_vend ='"& session("vTipo_comp") &"'"
	else
	
	
	stringTipoComp = ""
	end if
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp
  if session("vQuartos_comp")<>"qqualquer" and session("vQuartos_comp")<>"" then
	stringQuartosComp = " and Quartos_vend >="& session("vQuartos_comp") &""
	else
	
	stringQuartosComp = ""
	end if
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp
  if session("vVagas_comp") <> "vgqualquer" and session("vVagas_comp") <> "" then
	stringVagasComp = " and vagas_vend >="& session("vVagas_comp") &""
	else	
	stringVagasComp = ""
	end if
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 '----------------------------Valor pretendido----------------------------



	 dim stringValorComp
	 
	 
	 if session("vValor_comp") = "" then
	 session("vValor_comp")= request.QueryString("vValor_comp")
	 end if
	 
	 if session("vValor_comp1") = "" then
	 session("vValor_comp1")= request.QueryString("vValor_comp1")
	 end if
	 
	 if session("vValor_comp2") = "" then
	 session("vValor_comp2")= request.QueryString("vValor_comp2")
	 end if
	 
  if session("vValor_comp")<>"vqualquer" and session("vValor_comp")<>"" then
	'stringValorComp = " and Valor_vend >="& session("vValor_comp1") &" and Valor_vend <="& session("vValor_comp2") &""
	stringValorComp = " and Valor_vend <="& session("vValor_comp2") &""
	
	else
	
	
	stringValorComp = ""
	end if
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	if session("vCidade_vend") <> "" then
	
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringTipoComp&stringQuartosComp&stringVagasComp&stringValorComp
	
	else
	
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where cod_permuta =0"
	
	end if
	
	
	
	
	
	
	
	
	
	
	
	
	if vNome = "" then
	vNome = "n?o informado"
	end if
	
	if vTelefone = "" then
	vTelefone = "n?o informado"
	end if
	
	
	 dim vEnderecoIP , vData
  vData = now()
  
 
 vEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	
	
	
dim rs444VerificaConta2,strSQL444VerificaConta2
 dim rs444VerificaConta3,strSQL444VerificaConta3
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where telefone like '%"&session("telefone")&"%' or telefone02 like'%"&session("telefone")&"%' or telefone03 like'%"&session("telefone")&"%'" 
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rs444VerificaConta2.ActiveConnection = Conexao3
	
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao3	
	
	
	 Set rs444VerificaConta3 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta3 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou FROM compradores where telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	 rs444VerificaConta3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

 rs444VerificaConta3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

 rs444VerificaConta3.ActiveConnection = Conexao3
	
	
	
	
	
	 rs444VerificaConta3.Open strSQL444VerificaConta3, Conexao3	
	
	
	
	
	
	
	if rs444VerificaConta2.eof and vTipo <> "" and rs444VerificaConta3.eof then
	
	'Conexao.execute"Insert into permuta_procurados(Nome,Telefone,cidade_vend, bairro_vend,tipo_vend,quartos_vend,valor_vend,cidade_comp,bairro_comp,tipo_comp,quartos_comp,valor_comp,enderecoIP,data,vagas_comp,vagas_vend)values( '"& vNome &"','"& vTelefone &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& vQuartos_vend &"','"& vValor_vend &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& vValor_comp &"','"& vEnderecoIP &"','"& vData &"','"& session("vVagas_comp") &"','"& session("vVagas_vend") &"')" 
	
	
	 if session("vVagas_comp") <> "vgqualquer" then
	session("vVagas_comp") = session("vVagas_comp")
	else	
	session("vVagas_comp") = "0"
	end if
	
	if vQuartos_comp <> "qqualquer"   then
	vQuartos_comp = vQuartos_comp
	else
	vQuartos_comp = "0"
	end if
	
	if vValor_comp <> "vqualquer" then
	vValor_comp = vValor_comp
	else
	vValor_comp = "0"
	end if
	
	
	
	
	 if vVagas_vend <> "vgqualquer" then
	vVagas_vend = vVagas_vend
	else	
	vVagas_vend = "0"
	end if
	
	if vQuartos_vend <> "qqualquer" then
	vQuartos_vend = vQuartos_vend
	else
	vQuartos_vend = "0"
	end if
	
	if vValor_vend <> "vqualquer" then
	vValor_vend = vValor_vend
	else
	vValor_vend = "0"
	end if
	
	
	
	if vCidade_comp <> "cqualquer" then
	vCidade_comp = vCidade_comp
	else
	vCidade_comp = "n?o informado"
	end if
	
	
	
	if vBairro_comp <> "bqualquer" then
	vBairro_comp = vCidade_comp
	else
	vBairro_comp = "n?o informado"
	end if
	
	
	
	
	
	'Conexao3.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& int(varValorMedioComp) &"','"& now() &"','"& "n?o informado" &"','"& "internet" &"','"& now() &"','"& "n?o informado" &"','"& session("vVagas_comp") &"','"& "n?o informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "N?o informado" &"','"& "internet" &"')"
	
	
				
'Conexao3.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade,indexador_indicacoes) values( '"& session("nome") &"','"& "n?o informado" &"','"& session("telefone") &"','"& session("email") &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "icon_foto2.gif" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "0" &"','"& "0" &"','"& vQuartos_vend &"','"& "n?o informado" &"','"& vVagas_vend&"','"& "venda" &"','"& int(varValorMedioVend) &"','"& now() &"','"& "n?o informado" &"','"& "n?o informado" &"','"& "excluido" &"','"& "n?o informado" &"','"& "n?o informado" &"','"& "excluido" &"','"& "n?o informado" &"','"& "internet" &"','"& now() &"','"& "n?o informado" &"','"& "Sem Placa"&"','"& "0" &"','"& "0" &"','"& "0" &"','"& "neg?cio comum" &"','"&"0"&"')"	 

	
	
	
  end if
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
   dim rs444VerificaConta02,strSQL444VerificaConta02
   
    Set rs444VerificaConta02 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta02 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone ='"&session("telefone")&"'" 
	
	

	rs444VerificaConta02.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444VerificaConta02.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rs444VerificaConta02.ActiveConnection = Conexao3
	
	
	
	 rs444VerificaConta02.Open strSQL444VerificaConta02, Conexao3
	


  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 
if vTipo_vend <> "" then
	
	Conexao3.execute"Insert into permuta_procurados(Nome,Telefone,cidade_vend, bairro_vend,tipo_vend,quartos_vend,valor_vend,cidade_comp,bairro_comp,tipo_comp,quartos_comp,valor_comp,enderecoIP,data,vagas_comp,vagas_vend,origem_franquia)values( '"& vNome &"','"& vTelefone &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& vQuartos_vend &"','"& vValor_vend &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& vValor_comp &"','"& vEnderecoIP &"','"& vData &"','"& session("vVagas_comp") &"','"& session("vVagas_vend") &"','"& session("vOrigem_Franquia") &"')" 
	
  end if
%>




<%

'--------------------------Fazer listagem do bairro atual--------
dim rs444
dim strSQL444



 Set rs444 = Server.CreateObject("ADODB.RecordSet")



if session("vCidade_vend2") <> "cqualquer" and session("vCidade_vend2") <> "" then

strSQL444 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = '"&session("vCidade_vend2")&"'  ORDER BY nome_combo2" 

else

strSQL444 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 



end if

 rs444.CursorLocation = 3
rs444.CursorType = 3

rs444.ActiveConnection = Conexao3



		
		rs444.Open strSQL444, Conexao3



'-------------------------------------------------------------------


'-----------------------Listagem do bairro pretendido----------------
dim rs666
dim strSQL666

Set rs666 = Server.CreateObject("ADODB.RecordSet")



if session("vCidade_comp2") <> "cqualquer" and session("vCidade_comp2") <> "" then

 strSQL666 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = '"&session("vCidade_comp2")&"'  ORDER BY nome_combo2" 
	
else

strSQL666 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 =4  ORDER BY nome_combo2" 
		
end if


 rs666.CursorLocation = 3
rs666.CursorType = 3

rs666.ActiveConnection = Conexao3



        
		
		
		
		
		
		rs666.Open strSQL666, Conexao3



'-------------------------------------------------------------------






%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Procurar permutante</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script>

// Verifica se somentenúmeros foram digitados no campo
function isValidDigitNumber2 (doublecombo2) 
{


{

if (doublecombo2.txt_valor_vend.value == "vqualquer") {
		alert("Por favor, coloque o valor do seu imóvel.");
		doublecombo2.txt_valor_vend.focus();
		
		return false;
}


if (doublecombo2.txt_valor_comp.value == "vqualquer") {
		alert("Por favor, coloque o valor do imóvel que pretende adquirir.");
		doublecombo2.txt_valor_comp.focus();
		
		return false;
}

}
}


</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>



</head>

<body topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">


<form name="doublecombo2"  onSubmit="return isValidDigitNumber2(this);"  method="get" action="procurar_permuta001.asp">

<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="106"><img src="top01.jpg" width="794" height="106"></td>
  </tr>
  <tr>
      <td height="60" bgcolor="#e0a94e"><div align="center"><font color="#e0a94e" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000" size="3">Est? 
          Esta é a página para você encontrar uma pessoa para trocar o imóvel 
          com você</font></strong></font></div></td>
  </tr>
  <tr>
    <td height="260" bgcolor="#e0a94e"> 
      <table width="784" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="784" height="250" bgcolor="#e6dca9"><table width="774" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><table width="764" height="230" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="764" height="230" bgcolor="#e6dca9"><table width="754" height="220" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="754" height="220"><table width="754" height="220" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td width="372" height="220"><table width="372" border="0" cellspacing="0" cellpadding="0">
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                              <input name="txt_nome" type="text" class="inputBox"  id="txt_nome" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" onfocus="doublecombo2.txt_nome.value=''" value="<% if session("nome") <> "" then response.write session("nome") else response.write "Seu nome:" end if%>" size="30" maxlength="30">
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                              <input name="txt_telefone" type="text" class="inputBox"  id="txt_nome7" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" onfocus="doublecombo2.txt_telefone.value=''" value="<% if session("telefone") <> "" then response.write session("telefone") else response.write "Seu telefone:" end if%>" size="30" maxlength="30">
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                              <input name="txt_email" type="text" class="inputBox"  id="txt_nome8" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" onfocus="doublecombo2.txt_email.value=''" value="<% if session("email") <> "" then response.write session("email") else response.write "Seu email:" end if%>" size="30" maxlength="30">
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="combo1"  id="combo1" class="inputBox" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" onChange="javascript:atualizacarros(this.form);">
                                              <option value="cqualquer" selected>Qual a cidade do seu imóvel ?</option>
            <option value="cqualquer">Qualquer cidade</option>
           
		   
		    <% if not rs333.eof then %>
            <% While NOT Rs333.EoF %>
           <option value="<% = Rs333("id_combo1") %>"<%if session("vCidade_vend2")<> "cqualquer" then%><%if int(rs333("id_combo1")) = int(session("vCidade_vend2")) then response.write "selected" else response.write "" end if %><%end if%>> 
		   
            <% = Rs333("nome_combo1") %>
            </option>
            </option>
            <% Rs333.MoveNext %>
            <% Wend %>
            <%else%>
            <option value=""></option>
            <%end if%>
          </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="combo2" id="combo2"  class="inputBox"  style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                             <option value="bqualquer" selected>Qual o bairro do seu imóvel ?</option>
              <option value="bqualquer">Qualquer bairro</option>
             
			 
			  <% if not rs444.eof then%>
              <% While NOT Rs444.EoF %>
             
			  <option value="<% = Rs444("id_combo2") %>" <% if session("vBairro_vend2") <> "bqualquer" then if Rs444("id_combo2") = int(session("vBairro_vend2"))  then response.write "selected" else response.write "" end if end if %>> 
                
            
			
			<% = Rs444("nome_combo2") %>
            </option>
             
			  <% Rs444.MoveNext %>
              <% Wend %>
              <% else %>
              <option value=""></option>
              <% end if %>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                              <option value="<%=session("vTipo_vend")%>" selected><%if session("vTipo_vend") <> "tqualquer" and session("vTipo_vend") <> "" then  response.write session("vTipo_vend") else response.write "Qual o tipo do seu imóvel?" end if%></option>
				 
                  <option value="tqualquer">Qualquer um</option>
				  	<% if not rs444Tipo23.eof then%>
					<% While NOT rs444Tipo23.EoF %>
                    <option value="<% = rs444Tipo23("tipo") %>">
                    <% =rs444Tipo23("tipo") %>
                    </option>
                    <% rs444Tipo23.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
                </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_quartos_vend" size="1" id="txt_quartos_vend" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                           <option value="<%=session("vQuartos_vend")%>"><% if session("vQuartos_vend") <> "qqualquer" and session("vQuartos_vend") <> "" then response.write session("vQuartos_vend") else response.write "Quantos quartos tem o seu imóvel?" end if%></option>
										 
			   <option value="qqualquer">Qualquer um</option>
			 			 
              <option value="01" >01</option>
              <option value="02">02 </option>
              <option value="03">03</option>
              <option value="04">04</option>
              <option value="05">05</option>
              <option value="06">06</option>
			  <option value="04">07</option>
              <option value="05">08</option>
              <option value="06">09</option>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_vagas_vend" size="1" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                             
											  <option value="<%=session("vVagas_vend")%>"><% if session("vVagas_vend") <> "vgqualquer" and session("vVagas_vend") <> "" then response.write session("vVagas_vend") else response.write "Quantas vagas na garagem tem o seu imóvel" end if%></option>
									   
			  
			  <option value="vgqualquer">Qualquer um</option>
			 
              <option value="01" >01</option>
              <option value="02">02 </option>
              <option value="03">03</option>
              <option value="04">04</option>
              <option value="05">05</option>
              <option value="06">06</option>
			  <option value="04">07</option>
              <option value="05">08</option>
              <option value="06">09</option>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_valor_vend" size="1" id="txt_valor_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                             
				  
           
			 
			   <option value="<%=session("vValor_vend")%>" selected><% if session("vValor_vend") <> "vqualquer" and session("vValor_vend") <> "" then response.write FormatNumber(session("vValor_vend1"),2)&" Até "&FormatNumber(session("vValor_vend2"),2) else response.write "Qual o valor do seu imóvel?" end if%></option>
			 
			  <option value="vqualquer">Qualquer um</option>
			   <option value="0000000000 0000020000">Até 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 Até 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 Até 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 Até 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 Até 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 Até 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 Até 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 Até 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 Até 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 Até 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
			  
			   
			    </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center"></div></td>
                                      </tr>
                                    </table></td>
                                  <td width="10" height="220">&nbsp;</td>
                                  <td width="372" height="220"><table width="372" border="0" cellspacing="0" cellpadding="0">
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" onChange="javascript:atualizacarros2(this.form);">
                                              <option value="cqualquer" selected>Em qual cidade o sr(a) quer adquirir um imóvel ?</option>
              <option value="cqualquer" >Qualquer cidade</option>
             
			 
			  <% if not rs555.eof then %>
              <% While NOT Rs555.EoF %>
             
			 <option value="<% = Rs555("id_combo1") %>"<%if session("vCidade_comp2")<> "cqualquer" then%><%if int(rs555("id_combo1")) = int(session("vCidade_comp2")) then response.write "selected" else response.write "" end if %><%end if%>> 
		   
            <% = Rs555("nome_combo1") %>
            </option>
			  
			  
			  <% Rs555.MoveNext %>
              <% Wend %>
              <%else%>
              <option value=""></option>
              <%end if%>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="combo4" size="1"  class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                              <option value="bqualquer" >Em qual bairro o sr(a) quer adquirir um imóvel ?</option>
               <option value="bqualquer">Qualquer bairro</option>
              
			  
			  <% if not rs666.eof then%>
              <% While NOT Rs666.EoF %>
             
			   <option value="<% = Rs666("id_combo2") %>" <% if session("vBairro_comp2") <> "bqualquer" then if Rs666("id_combo2") = int(session("vBairro_comp2"))  then response.write "selected" else response.write "" end if end if %>> 
                
            
			
			<% = Rs666("nome_combo2") %>
            </option>
              
			  <% Rs666.MoveNext %>
              <% Wend %>
              <% else %>
              <option value=""></option>
              <% end if %>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_tipo_comp" size="1" id="txt_tipo_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                             
											   <option value="<%=session("vTipo_comp")%>" selected><%if session("vTipo_comp") <> "tqualquer" and session("vTipo_comp") <> "" then  response.write session("vTipo_comp") else response.write "Qual o tipo do  imóvel que o sr(a) deseja?" end if%></option>
				 
											   <option value="tqualquer" >Qualquer tipo</option>
			  	<% if not rs444Tipo24.eof then%>
					<% While NOT rs444Tipo24.EoF %>
                    <option value="<% = rs444Tipo24("tipo") %>">
                    <% =rs444Tipo24("tipo") %>
                    </option>
                    <% rs444Tipo24.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_quartos_comp" size="1" id="txt_quartos_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                             <option value="<%=session("vQuartos_comp")%>"><% if session("vQuartos_comp") <> "qqualquer" and session("vQuartos_comp") <> "" then response.write session("vQuartos_comp") else response.write "Quantos quartos tem o imóvel que sr(a) deseja?" end if%></option>
										 
			  <option value="qqualquer">Qualquer um</option>			  
             
			  <option value="01" >01</option>
              <option value="02">02 </option>
              <option value="03">03</option>
              <option value="04">04</option>
              <option value="05">05</option>
              <option value="06">06</option>
			  <option value="04">07</option>
              <option value="05">08</option>
              <option value="06">09</option>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_vagas_comp" size="1" id="txt_vagas_comp" class="inputBox"  style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                              <option value="<%=session("vVagas_comp")%>"><% if session("vVagas_comp") <> "vgqualquer" and session("vVagas_comp") <> "" then response.write session("vVagas_comp") else response.write "Quantas vagas na garagem tem o  imóvel que sr(a) deseja?" end if%></option>
									   
			  
			  
			  <option value="vgqualquer">Qualquer um</option>			  
             
			 
			  <option value="01" >01</option>
              <option value="02">02 </option>
              <option value="03">03</option>
              <option value="04">04</option>
              <option value="05">05</option>
              <option value="06">06</option>
			  <option value="04">07</option>
              <option value="05">08</option>
              <option value="06">09</option>
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center">
                                            <select name="txt_valor_comp" size="1" id="txt_valor_comp" class="inputBox"style="HEIGHT: 18px; WIDTH: 372px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                              
								 <option value="<%=session("vValor_comp")%>" selected><% if session("vValor_comp") <> "vqualquer" and session("vValor_comp") <> "" then response.write FormatNumber(session("vValor_comp1"),2)&" Até "&FormatNumber(session("vValor_comp2"),2) else response.write "Qual o valor do imóvel que sr(a) deseja comprar ou alugar?" end if%></option>
			 			  
											  
                  <option value="0000000000 0000020000">Até 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 Até 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 Até 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 Até 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 Até 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 Até 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 Até 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 Até 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 Até 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 Até 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
               
            </select>
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"> <div align="right">
                                              <input name="image22" type="image"  src="bt_procurar505.jpg" width="201" height="18"  border="0">
                                            </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center"> 
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center"> 
                                          </div></td>
                                      </tr>
                                      <tr> 
                                        <td width="372" height="22"><div align="center"></div></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  
  <%





Set RS = Server.CreateObject("ADODB.Recordset")
'um objeto recordset ? inst?nciado.

Dim LinkTemp
'essa variável vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
'as vari?veis acima s?o usadas para trocar a cor das tabelas que conter?o os valores
'dos recordsets.






dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeirapágina.

dim intPageCount
'Essa variável vai receber o valor da quantidade depáginas do recordset.

dim intRecordCount
'Essa variável vai receber onúmero de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a variável intPage recebe o valor "1" na primeirapágina.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conex?o o recordset utilizar?.
	
RS.Open strSQL, Conn, 1, 3
'o recordset ? aberto
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros porpágina.

RS.CacheSize = RS.PageSize
'o Cache tamb?m conter? 20 registros porpágina.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor donúmero depágina do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor donúmero de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%>
  
  
  
  
  
  <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount end if
'se intPage ? maior que onúmero depáginas ent?o intPage ? igual aonúmero depáginas.

	If CInt(intPage) <= 0 Then intPage = 1 end if
	'se intPage ? menor ou igual a zero ent?o intPage igual a "1"
	'a variável intPage sempre vai ser for?ada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados ent?o.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina apágina exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posi??o exata do primeiro registro dapágina correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual aonúmero depáginas no recordset , estamos na ?ltima 
			'p?gina ent?o.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor donúmero do ?ltimo recordset.
				'intFinish corresponde ao valor do ?ltimo registro dapágina correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'donúmero de registros napágina menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros ent?o
		For intRecord = 1 to RS.PageSize
		'um contador inRecord ? colocado Até onúmero de registros napágina.

dim varCodPermuta




	


%>
<% varCodPermuta =RS("cod_permuta") %>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="794" height="190"><table width="784" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td style="border:1px solid #ddddc5;"><table width="774" height="170" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e9dca8"><div align="center"><font face="Verdana, arial" size="2" color="FFFFFF"><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')" style="color:#e0a94e;text-decoration:none;"><strong>Ol?, 
                      meu nome &eacute; <%=rs("nome")%> ,o sitema VEJA analizou 
                      os dados do seu e do meu imóvel e dectetou a possibilidade 
                      de efetuarmos uma permuta entre nossos imóveis. Lique já
                      para 4123-72-44 e fale com meu atendente sr(a) <%=rs("atendimento")%>. 
                      para que cada um de n?s visitemos os imóveis de um e de 
                      outro, para ver mais detalhes do meu imóvel clique aqui, 
                      muito obrigado.</strong></a> </font></div></td>
              </tr>
            </table></td>
        </tr>
      </table> </tr>
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
 'acima ? feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
	  
	  
	  
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><font color="#000000" size="1" face="Verdana, arial"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se apágina atual for maior que "1" ent?o o link anteriro ? colocado na 
			  na tela .-->
               <a href="?page=<%=intPage - 1%>&vCidade_vend=<%=session("vCidade_vend")%>&vCidade_vend2=<%=session("vCidade_vend2")%>&vBairro_vend=<%=session("vBairro_vend")%>&vBairro_vend2=<%=session("vBairro_vend2")%>&vVila_vend=<%=session("vVila_vend")%>&vVila_vend2=<%=session("vVila_vend2")%>&vTipo_vend=<%=session("vTipo_vend")%>&vNegociacao_vend=<%=session("vNegociacao_vend")%>&vQuartos_vend=<%=session("vQuartos_vend")%>&vVagas_vend=<%=session("vVagas_vend")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_vend1=<%=session("vValor_vend1")%>&vValor_vend2=<%=session("vValor_vend2")%>&vCidade_comp=<%=session("vCidade_comp")%>&vCidade_comp2=<%=session("vCidade_comp2")%>&vBairro_comp=<%=session("vBairro_comp")%>&vBairro_comp2=<%=session("vBairro_comp2")%>&vVila_comp=<%=session("vVila_comp")%>&vVila_comp2=<%=session("vVila_comp2")%>&vTipo_comp=<%=session("vTipo_comp")%>&vNegociacao_comp=<%=session("vNegociacao_comp")%>&vQuartos_comp=<%=session("vQuartos_comp")%>&vVagas_comp=<%=session("vVagas_comp")%>&vValor_comp1=<%=session("vValor_comp1")%>&vValor_comp2=<%=session("vValor_comp2")%>&vValor_comp=<%=session("vValor_comp")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="color:#000000;text-decoration:none;">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td>
              
			  <!-- sepágina atual ? menor que o total depáginas e intPage maior que um
			  ou seja, se n?o estiver na primeirapágina e nem na ?ltima ent?o. -->
			  <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- sepágina atual ? menor que o total depáginas e intPage maior que um
			  ou seja, se n?o estiver na primeirapágina e nem na ?ltima ent?o. -->
       página <%=cInt(intPage)%> de 
        <%=cInt(intPageCount)%> </font> 
        <%End If%></font>
        </div>
             
             </td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage ? menor que onúmero depáginas ent?o colocar o bot?o pr?ximo -->
                <a href="?page=<%=intPage + 1%>&vCidade_vend=<%=session("vCidade_vend")%>&vCidade_vend2=<%=session("vCidade_vend2")%>&vBairro_vend=<%=session("vBairro_vend")%>&vBairro_vend2=<%=session("vBairro_vend2")%>&vVila_vend=<%=session("vVila_vend")%>&vVila_vend2=<%=session("vVila_vend2")%>&vTipo_vend=<%=session("vTipo_vend")%>&vNegociacao_vend=<%=session("vNegociacao_vend")%>&vQuartos_vend=<%=session("vQuartos_vend")%>&vVagas_vend=<%=session("vVagas_vend")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_vend1=<%=session("vValor_vend1")%>&vValor_vend2=<%=session("vValor_vend2")%>&vCidade_comp=<%=session("vCidade_comp")%>&vCidade_comp2=<%=session("vCidade_comp2")%>&vBairro_comp=<%=session("vBairro_comp")%>&vBairro_comp2=<%=session("vBairro_comp2")%>&vVila_comp=<%=session("vVila_comp")%>&vVila_comp2=<%=session("vVila_comp2")%>&vTipo_comp=<%=session("vTipo_comp")%>&vNegociacao_comp=<%=session("vNegociacao_comp")%>&vQuartos_comp=<%=session("vQuartos_comp")%>&vVagas_comp=<%=session("vVagas_comp")%>&vValor_comp1=<%=session("vValor_comp1")%>&vValor_comp2=<%=session("vValor_comp2")%>&vValor_comp=<%=session("vValor_comp")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="color:#000000;text-decoration:none;"><b>Pr?ximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  
  <%end if%>
  
  
  
  
</table>


<% else%>
  
 
  <table width="794" border="0" cellspacing="0" cellpadding="0">
    <tr><td height="20"></td></tr>
	<tr>
      <td height="300" bgcolor="#e0a94e"> 
        <div align="center">
          <table width="785" height="290" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td bgcolor="#e6dca9"><div align="center"><font color="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
			  
			  <% if vCidade_vend <> "" then %>
                  N?o foi encontrado nenhum permutante para você. 
                  <%else%>
                  Preencha os dados acima para encontrar alguém para permutar(Trocar) de imóvel com você. 
                  <%end if%>
				  </strong></font>
		   
		    </tr>
          </table>
        </div></td>
  </tr>
</table>
  
  
  
  <%end if%>

</form>




<%



dim varValorMedioComp
	
	varValorMedioComp= (int(session("vValor_comp1")) + int(session("vValor_comp2")))/2
	
	
	dim varValorMedioVend
	
	varValorMedioVend= (int(session("vValor_vend1")) + int(session("vValor_vend2")))/2
	


dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais FROM permuta where telefone ='"&session("telefone")&"' " 
	
	
	rs444VerificaConta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444VerificaConta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rs444VerificaConta.ActiveConnection = Conexao3
	
	
	
	
	
	
	 rs444VerificaConta.Open strSQL444VerificaConta, Conexao3
	

if  rs444VerificaConta.eof and vTipo_vend <> "" then


Conexao3.execute"Insert into permuta(Foto_imovel,Nome,Email,Telefone,endereco_vend,cidade_vend,bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby,origem_franquia)values( '"& "imovel00000.jpg" &"','"& session("nome") &"','"& session("email") &"','"& session("telefone") &"','"& "n?o informado" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "n?o informado" &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& "n?o informado" &"','"& "0" &"','"& "n?o informado" &"','"& now() &"','"& vQuartos_vend &"','"& vQuartos_comp &"','"& int(vValorMedio_vend) &"','"& int(vValorMedio_comp) &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"&vVila_comp&"','"& vVagas_vend &"','"& vVagas_comp &"','"& "excluido" &"','"& session("vOrigem_Franquia") &"')" 
	


 end if
%>










<%


'Cadastrar em compradores

dim rs444VerificaConta22,strSQL444VerificaConta22
   
    Set rs444VerificaConta22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta22 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	rs444VerificaConta22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444VerificaConta22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rs444VerificaConta22.ActiveConnection = Conexao3
	
	
	
	
	
	
	 rs444VerificaConta22.Open strSQL444VerificaConta22, Conexao3
	

if  rs444VerificaConta22.eof and vTipo_vend <> "" then


Conexao3.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem,area_total,area_construida,condominio,condicoes_pagamento,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& int(varValorMedioComp) &"','"& now() &"','"& "n?o informado" &"','"& "internet" &"','"& now() &"','"& vVila_comp &"','"& vVagas_comp &"','"& "n?o informado" &"','"& "comprador a contatar" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "N?o informado" &"','"& "internet" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "n?o informado" &"','"& session("vOrigem_Franquia") &"')"


'Conexao3.execute"Insert into permuta(Foto_imovel,Nome,Email,Telefone,endereco_vend,cidade_vend,bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby)values( '"& "imovel00000.jpg" &"','"& session("nome") &"','"& session("email") &"','"& session("telefone") &"','"& "n?o informado" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "n?o informado" &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& "n?o informado" &"','"& "0" &"','"& "n?o informado" &"','"& now() &"','"& vQuartosConta_vend &"','"& vQuartosConta_comp &"','"& vValorMedio_vend &"','"& vValorMedio_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"&vVila_comp&"','"& vVagasConta_vend &"','"& vVagasConta_comp &"','"& "excluido" &"')" 
		

 end if
%>





<%


'Cadastrar em compradores

dim rs444VerificaConta23,strSQL444VerificaConta23
   
    Set rs444VerificaConta23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta23 = "SELECT imoveis.cod_imovel,imoveis.telefone,imoveis.telefone02,imoveis.telefone03  FROM imoveis  where telefone like'%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%'" 
	
	
	rs444VerificaConta23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444VerificaConta23.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rs444VerificaConta23.ActiveConnection = Conexao3
	
	
	
	
	
	
	 rs444VerificaConta23.Open strSQL444VerificaConta23, Conexao3
	

if  rs444VerificaConta23.eof and vTipo_vend <> "" then

Conexao3.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade,indexador_indicacoes,origem_captacao,data_captacao,cliques_no_imovel,tarja02,data01_tarja02,data02_tarja02,imovel_em_negociacao,origem_franquia) values( '"& session("nome") &"','"& "n?o informado" &"','"& session("telefone") &"','"& session("email") &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "icon_foto2.gif" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "0" &"','"& "0" &"','"& vQuartos_vend &"','"& "n?o informado" &"','"& vVagas_vend&"','"& "venda" &"','"& int(varValorMedioVend) &"','"& now() &"','"& "n?o informado" &"','"& "n?o informado" &"','"& "excluido" &"','"& "n?o informado" &"','"& "n?o informado" &"','"& "excluido" &"','"& "n?o informado" &"','"& "internet" &"','"& now() &"','"& "n?o informado" &"','"& "Sem Placa"&"','"& "0" &"','"& "0" &"','"& "0" &"','"& "neg?cio comum" &"','"&"0"&"','"&"Busca de permutante"&"','"& now()&"','"& "0"&"','"& "sim"&"','"& day(now())&"','"& day(DateAdd("d", 15, now()))&"','"& "imóvel n?o contatado" &"','"& session("vOrigem_Franquia") &"')"	 
	

'Conexao3.execute"Insert into permuta(Foto_imovel,Nome,Email,Telefone,endereco_vend,cidade_vend,bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp,standby)values( '"& "imovel00000.jpg" &"','"& session("nome") &"','"& session("email") &"','"& session("telefone") &"','"& "n?o informado" &"','"& vCidade_vend &"','"& vBairro_vend &"','"& vTipo_vend &"','"& "n?o informado" &"','"& vCidade_comp &"','"& vBairro_comp &"','"& vTipo_comp &"','"& "n?o informado" &"','"& "0" &"','"& "n?o informado" &"','"& now() &"','"& vQuartosConta_vend &"','"& vQuartosConta_comp &"','"& vValorMedio_vend &"','"& vValorMedio_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"&vVila_comp&"','"& vVagasConta_vend &"','"& vVagasConta_comp &"','"& "excluido" &"')" 
		

 end if
%>




<% response.flush%>
  <%response.clear%>

<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receber? uma conexao aberta!
'Em funcoes, geralmente n?o criamos objetos do tipo conex?es!
'Opte por sempre deixar sua fun??o o mais compAtével poss?vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo2) {" & vbcrlf

'Essa fun??o JavaScript recebe o form em que est?o os campos a serem atualizados!
'Veja na chamada da fun??o no m?todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo2.combo1.options[doublecombo2.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op??es de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 


Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")


 rsMarcas3.CursorLocation = 3
rsMarcas3.CursorType = 3

rsMarcas3.ActiveConnection = Conexao3



        rsMarcas3.Open sqlMarcas3, Conexao3





While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo2.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"



Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")


 rsCarros3.CursorLocation = 3
rsCarros3.CursorType = 3

rsCarros3.ActiveConnection = Conexao3



        rsCarros3.Open sqlCarros3, Conexao3



'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo2.combo2.options[" & i  & "] = new Option('" & "Qual o bairro do seu imóvel ?" & "','" & "bqualquer" & "');"
i = 1 
While NOT rsCarros3.EoF

Response.Write "doublecombo2.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo2.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Pr?xima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da fun??o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas3.close

set rsMarcas3 = nothing

rsCarros3.close

set rsCarros3 = nothing



End Function
%> 

<%
Function EscreveFuncaoJavaScript222 ( Conexao3 )
'O parametro conexao receber? uma conexao aberta!
'Em funcoes, geralmente n?o criamos objetos do tipo conex?es!
'Opte por sempre deixar sua fun??o o mais compAtével poss?vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo2) {" & vbcrlf

'Essa fun??o JavaScript recebe o form em que est?o os campos a serem atualizados!
'Veja na chamada da fun??o no m?todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo2.combo3.options[doublecombo2.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op??es de carro!
SqlMarcas444 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 




Set rsMarcas444 = Server.CreateObject("ADODB.RecordSet")


 rsMarcas444.CursorLocation = 3
rsMarcas444.CursorType = 3

rsMarcas444.ActiveConnection = Conexao3



        rsMarcas444.Open sqlMarcas444, Conexao3





While NOT rsMarcas444.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas444("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo2.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros444 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas444("id_combo1")&" order by nome_combo2"




Set rsCarros444 = Server.CreateObject("ADODB.RecordSet")


 rsCarros444.CursorLocation = 3
rsCarros444.CursorType = 3

rsCarros444.ActiveConnection = Conexao3



        rsCarros444.Open sqlCarros444, Conexao3




'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
 i = 0 
Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & "Em qual bairro o sr(a) quer adquirir um imóvel ?" & "','" & "bqualquer" & "');" & vbcrlf
i = 1
While NOT rsCarros444.EoF


Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & rsCarros444("nome_combo2") & "','" & rsCarros444("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros444.MoveNext
Wend
Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" & vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Pr?xima marca! 
rsMarcas444.MoveNext 
Wend 

'Fecha chaves do switch e da fun??o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas444.close

set rsMarcas444 = nothing

rsCarros444.close

set rsCarros444 = nothing


End Function
%> 
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript ( Conexao3) %>

<%
'response.write session("vValor_comp1")&"||"&vValor_comp&"||"&varValorMedioComp
'response.write (int(session("vValor_comp1")) + int(session("vValor_comp2")))
'response.write (int(session("vValor_comp1")) + int(session("vValor_comp2")))/2

'session("vValor_comp1")=left(vValor_comp,10)
  ' session("vValor_comp2")=right(vValor_comp,10)

'response.write int(varValorMedioComp)&"<br>"&(int(session("vValor_comp1")) + int(session("vValor_comp2")))/2

%>
<%'=strSQL%>
</body>
</html>
