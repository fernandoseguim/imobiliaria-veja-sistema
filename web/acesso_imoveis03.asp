<!--#include file="dsn.asp"-->
<%
Function EscreveFuncaoJavaScript ( Conexao3 )
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

rsMarcas3.ActiveConnection = Conexao3
	
	
	rsMarcas3.Open SqlMarcas3, Conexao3



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

rsCarros3.ActiveConnection = Conexao3
	
	
	rsCarros3.Open SqlCarros3, Conexao3



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
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open Sql3, Conexao3

%> 


<%
Function EscreveFuncaoJavaScript222 ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros222 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo2.options[doublecombo.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 

Set rsMarcas333 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas333.ActiveConnection = Conexao3
	
	
	rsMarcas333.Open SqlMarcas333,Conexao3


While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where id_combo2 =" & rsMarcas333("id_combo2")&""

Set rsCarros333 = Server.CreateObject("ADODB.RecordSet")

	rsCarros333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros333.ActiveConnection = Conexao3
	
	
	rsCarros333.Open SqlCarros333, Conexao3
'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1   
While NOT rsCarros333.EoF

Response.Write "doublecombo.combo5.options[" & i & "] = new Option('" & rsCarros333("nome_combo3") & "','" & rsCarros333("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros333.MoveNext
Wend


Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas333.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 


rsMarcas333.close

set rsMarcas333 = nothing


rsCarros333.close

set rsCarros333 = nothing



End Function
%> 





<%


Sql333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 



Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs333.ActiveConnection = Conexao3
	
	
	rs333.Open Sql333, Conexao3
	



%> 











<!--#include file="cores.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel,objFSO
Dim rs2,strSQL2,varCodImovel
dim nome,telefone

	
dim varNome
varNome=request.form("nome")

if varNome = "" then
varNome = request.querystring("varNome")
end if

	
dim varTelefone
varTelefone=request.form("telefone")

if varTelefone = "" then
varTelefone = request.querystring("varTelefone")
end if


'--------------------atualizar ultimo acesso-------------------------




dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone ='"&varTelefone&"'" 
	 
	 
	 rs444VerificaConta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta.ActiveConnection = Conexao3
	 
	 
	 rs444VerificaConta.Open strSQL444VerificaConta, Conexao3
	

if  not rs444VerificaConta.eof then




	 Conexao3.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta("cod_compradores")
	 
	
	 


end if





'--------------------------------------------------------------------








varCod_imovel = request.form("CodImovel")

if varCod_imovel = "" then
varCod_imovel = request.querystring("varCod_imovel")
end if

	
	dim Conexao9,rs9

	Set rs9 = Server.CreateObject("ADODB.RecordSet") 
	
	
	 
	 if varNome <> "" and varTelefone <> "" then
	 strSQL9 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where proprietario='"&varNome&"' and telefone='"&varTelefone&"' and cod_imovel="&varCod_imovel
	else
	strSQL9 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where  cod_imovel="&varCod_imovel
	end if
	
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	  
	  
	 rs9.Open strSQL9, Conexao3
	 
	 if not rs9.eof then

 dim EnderecoIP
	 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
	 
	Conexao3.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs9("proprietario") &"','"& rs9("telefone") &"','"& rs9("cod_imovel") &"','"& "Imóveis" &"','"& EnderecoIP &"','"& now() &"')"
	


if varCod_imovel = "" then
varCod_imovel = request.querystring("varCod_imovel")
end if


varSucesso_imovel = request.form("CodComprador")
   
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
   Set rs2 = Server.CreateObject("ADODB.RecordSet")
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where cod_imovel="&varCod_imovel
	 
	 

   
RS.CursorLocation = 3
RS.CursorType = 3

RS2.CursorLocation = 3
RS2.CursorType = 3

        rs.Open strSQL, Conexao3 
		rs2.Open strSQL2, Conexao3
		
	
	dim Sql4,rs4
	  Set rs4 = Server.CreateObject("ADODB.RecordSet")
Sql4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 where nome_combo2 like '"& rs("bairro") &"' and cidade_combo2 like '"& rs("cidade") &"' ORDER BY nome_combo2" 


Set rs4 = Server.CreateObject("ADODB.RecordSet")

	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao3
	
	
	rs4.Open Sql4, Conexao3



dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 FROM combo3 where nome_combo3 ='"&rs("vila")&"' and cidade_combo3 ='"&rs("cidade")&"' and bairro_combo3 ='"&rs("bairro")&"'  ORDER BY nome_combo3" 
	
	rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444.ActiveConnection = Conexao3
	
	
	 rs444.Open strSQL444, Conexao3		







dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 where nome_combo1 ='"&rs("cidade")&"'  ORDER BY nome_combo1" 
	
	
	rs555.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs555.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs555.ActiveConnection = Conexao3
	
	
	
	
	
	 rs555.Open strSQL555, Conexao3		




'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 
	 rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao3
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao3










'-------------------------------------------------------------------------------------------------




'----------------------------verifica imóvel---------------------------
	
	
	dim rsVerifica2
	dim strSQLVerifica2
	
 Set rsVerifica2 = Server.CreateObject("ADODB.RecordSet")
    
	
	dim varCod_imovelVerifica
	
	if rs("cod_imovel") <> "" and rs("cod_imovel") <> "0" then
	
	varCod_imovelVerifica = rs("cod_imovel")
	else
	varCod_imovelVerifica = "0"
	end if
	
	
	
	strSQLVerifica2 = "SELECT proposta.Cod_proposta,proposta.proposta_proposta,proposta.foto_proposta,proposta.foto_proposta,proposta.nome_proposta,proposta.telefone_proposta,proposta.email_proposta,proposta.data_proposta,proposta.horario_proposta,proposta.interesse_proposta,proposta.cod_imovel_proposta  FROM proposta where cod_imovel_proposta='"&varCod_imovelVerifica&"'"
	 
   
   
rsVerifica2.CursorLocation = 3
rsVerifica2.CursorType = 3

        rsVerifica2.Open strSQLVerifica2, Conexao3 	
		
		
	
	


varCodImovel = rs("cod_imovel")






'----------------------------adicionar acesso--------------------------------



if session("telefone") <> "" and session("acessos") = "" then



dim rs444VerificaConta022,strSQL444VerificaConta022
   
    Set rs444VerificaConta022 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta022 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&session("telefone")&"%' or telefone02 like'%"&session("telefone")&"%' or telefone03 like'%"&session("telefone")&"%'" 
	 
	 
	 
	

	rs444VerificaConta022.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta022.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta022.ActiveConnection = Conexao3
	
	
	
	 
	 
	 
	 
	 
	 rs444VerificaConta022.Open strSQL444VerificaConta022, Conexao3
	


  dim vNumero_acessos
  
  vNumero_acessos = rs444VerificaConta022("acessos")
  
  vNumero_acessos = int(vNumero_acessos) + 1


if  not rs444VerificaConta022.eof then




	 Conexao3.execute"update compradores set data_ultimo_acesso='"&now()&"',acessos='"&vNumero_acessos&"'  where cod_compradores="&rs444VerificaConta022("cod_compradores")
	end if 
      
	  
	  session("acessos") = "acessado"
	  
	  

end if


'-----------------------------------fim no cadastro de número de acessos-------------
	





%>		





<html>
<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

<script>
function isValidDigitNumber (doublecombo)
{




	var strValidNumber1_6="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_valor.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.txt_valor.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números e a vírgula!");
doublecombo.txt_valor.focus();
doublecombo.txt_valor.select();
return false;
}
}


	
	if (doublecombo.txt_valor.value == "") {
        alert("O formulário Valor está vazio!");
        doublecombo.txt_valor.focus();
		doublecombo.txt_valor.select();
        return false;
    }
	
	if (doublecombo.combo2.value == "") {
        alert("O formulário Bairro do Imóvel está vazio!");
        doublecombo.combo1.focus();
		
        return false;
    }
	
	if (doublecombo.combo1.value == "") {
        alert("O formulário Cidade do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }





var elem=doublecombo.elements;





for (nCount=0; nCount < elem.length; nCount++)
  
    
  
	
	if(elem[nCount].type.indexOf("text")==0)	{
	var strValidNumber12_1="'";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)!=-1) {



alert("Este campo  não pode conter aspas");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}
//-----------------------------------------------

}










</script>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=700,height=500,resizable=yes')
   openWindow22.focus( )
   }

</SCRIPT>




</head>
<!--#include file="style_imoveis.asp"-->





<title>Sua conta de imóvel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<!--#include file="style_imoveis.asp"-->
<body bgcolor="<%=escuro%>" bottommargin="0" topmargin="20" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">

<br>
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><a href="acessoLink03.asp?varTelefone=<%=session("varTelefone")%>"><img  src="bt_voltar002.jpg" border="0" ></img></a></td>
    </tr>
  </table>
</center>

<br>
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="80"><div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Aqui 
          você poderá atualizar os seus dados de vendedor de imóveis, modificar 
          seus interesses ou simplesmente verificar se novos compradores foram 
          indicados pelo nosso sistema de acordo com seu interesse.</strong></font></div></td>
    </tr>
  </table>
</center>



<center>
<br>
 
  
  
  <br>
  
  <%if not rsVerifica2.eof then %>
                  
				 
  
				  <div align="center"><a href="javascript:newWindow22('listar_proposta_conta.asp?varCod_proposta_imovel=<%=rsVerifica2("cod_imovel_proposta")%>')" style="color:#FFFFFF"> <font size="3" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" ><strong> Clique 
                    aqui e veja as propostas pelo seu imóvel.</strong></font></a> 
                    </strong></font>
					
					<% else %>
                    <%end if%>
                  </div>
  
  <br>
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_imoveis03.asp?varCod_imovel=<%=varCod_imovel%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
 
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td style="border:1px solid #FFFFFF;"><table width="580" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="290" height="170"  bgcolor="<%=medio%>">&nbsp;</td>
                <td width="290" height="170" bgcolor="<%=escuro%>" ><% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                    <div align="center"><img src="<%=rs("foto_grande")%>" name="photoslider" width="290" height="170"></img></div>
                      <% else %>
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div>
                    <% end if %></td>
				
				
				
				
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  
  <% if rs("foto_grande2")<>"imovel00000.jpg" then%>
  <tr>
  <td height="18"><table width="590" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td width="580"><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
				<script language="JavaScript">
                         var photos=new Array()
                         var which=0
                         
photos[0]="<%=rs("foto_grande1")%>"
photos[1]="<%=rs("foto_grande2")%>"
photos[2]="<%=rs("foto_grande3")%>"
photos[3]="<%=rs("foto_grande4")%>"
photos[4]="<%=rs("foto_grande5")%>"
photos[5]="<%=rs("foto_grande6")%>"
photos[6]="<%=rs("foto_grande7")%>"
photos[7]="<%=rs("foto_grande8")%>"
photos[8]="<%=rs("foto_grande9")%>"
photos[9]="<%=rs("foto_grande10")%>"


 var tam = 0;
<% if rs("foto_grande1")<>"imovel00000.jpg"  then%>
                         var tam = 0;
						<%end if%>

<% if rs("foto_grande2")<>"imovel00000.jpg"  then %>
                         var tam = 1;
						<%end if%>
						
<% if rs("foto_grande3")<>"imovel00000.jpg"  then%>
                         var tam = 2;
						<%end if%>					 
                       
 <% if rs("foto_grande4")<>"imovel00000.jpg"  then %>
                         var tam = 3;
						<%end if%>
						
<% if rs("foto_grande5")<>"imovel00000.jpg"  then %>
                      var tam = 4;
						<%end if%>
						
<% if rs("foto_grande6")<>"imovel00000.jpg"  then %>
                      var tam = 5;
						<%end if%>
						
<% if rs("foto_grande7")<>"imovel00000.jpg"  then %>
                      var tam = 6;
						<%end if%>
												
<% if rs("foto_grande8")<>"imovel00000.jpg"  then %>
                      var tam = 7;
						<%end if%>
						
						
<% if rs("foto_grande9")<>"imovel00000.jpg"  then %>
                      var tam = 8;
						<%end if%>
																							
<% if rs("foto_grande10")<>"imovel00000.jpg"  then %>
                      var tam = 9;
						<%end if%>					   
					     function anterior(){
                           if (which>0){
                             which--
                           }else{
                             which=tam;
                           }
                           document.images.photoslider.src=photos[which]
                         }
                         function proxima(){
                           if (which<tam){
                             which++
                           }else{
                             which=0;
                           }
                           document.images.photoslider.src=photos[which]
                         }
                      </script>
                  <td width="290">&nbsp;</td>
                  <td width="290"> <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:anterior()" class="link" onmouseover="window.status='Anterior'; return true" onmouseout="window.status=''"><img src="bt_anterior002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:proxima()" class="link" onmouseover="window.status='Próxima'; return true" onmouseout="window.status=''"><img src="bt_proxima002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                      </tr>
                    </table> </td>
                </tr>
              </table></td>
            <td width="5">&nbsp;</td>
          </tr>
        </table></td>
  </tr>
  <%else%><%end if%>
  <tr>
        <td width="590" height="100"> 
          <div align="right"><a href="javascript:newWindow3('form_adicionar_foto.asp?varCodimovel=<%=varCodimovel%>')" style="font color:#FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Click 
            aqui e inclua mais uma foto na conta do seu imóvel.</strong></font></a>
			
			<br>
  <br>
            <strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Tecle 
            F5 do seu teclado para carregar a foto depois de incluída.</font></strong> 
          </div></td>
  
  </tr>
  <tr>
  
 
    <td height="18"><div align="right"><table width="590" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="570"><div align="center">
			  <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
        <%else%>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
        foi atualizado com sucesso.</font> 
        <% end if %>

			  
			  
			  
			  </div></td>
              
            <td>&nbsp; </td>
            </tr>
          </table></td>
  </tr>
  
  
  
  
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de inclus&atilde;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs("data")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      da &uacute;ltima atualiza&ccedil;&atilde;o</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs("data_atualizacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			   
			   
			   
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      do im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                    <div align="left"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Cod_imovel")%></font> 
                    </div></td>
              </tr>
			 
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Propriet&aacute;rio 
                    do im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                        <input name="txt_proprietario" type="text" id="txt_proprietario" value="<%=rs("proprietario")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;">
                  </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                    do propriet&aacute;rio do im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                        <input name="txt_telefone" type="text" id="txt_telefone" value="<%=rs("telefone")%>" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>; ">
                  </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                    do propriet&aacute;rio do im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                    <input name="txt_email" type="text" id="txt_email" value="<%=rs("email")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=medio%>;">
                  </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                    do Im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                        <input name="txt_endereco" type="text" id="txt_endereco" value="<%=rs("endereco")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>;">
                  </div></td>
              </tr>
              
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="<% if rs("cidade") = "não informado" or rs555.eof then response.write "cqualquer" else response.write rs555("id_combo1") end if  %>" select><%=rs("cidade")%></option>
                 <% if not rs3.eof then %>
				    <% While NOT Rs3.EoF %>
                    <option value="<% = Rs3("id_combo1") %>">
                    <% = Rs3("nome_combo1") %>
                    </option>
                    <% Rs3.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
				 
				  </select>
                    </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF"> 
                      <select name="combo2" class="inputBox" onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                       
                        <option value="<%if rs("bairro") = "não informado" or rs4.eof then response.write "bqualquer" else response.write rs4("id_combo2") end if%>" ><%=rs("bairro")%></option>
                       
                        <option value=""></option>
                      
                      </select>
                      </font></div></td>
              </tr>
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF"> 
                      <select name="combo5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="<%if rs("vila") <> "não informado" and not rs444.eof then response.write rs444("id_combo3") else response.write "vlqualquer" end if%>" selected><%=rs("vila")%></option>
                      </select>
                      </font></div></td>
              </tr>
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF">
                    <select name="txt_tipo" size="1" id="txt_tipo" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>;">
                      <option value="<%if rs("tipo") <> "tqualquer" then response.write rs("tipo") else response.write "não informado" end if%>" selected><%if rs("tipo") <> "tqualquer" then response.write rs("tipo") else response.write "não informado" end if%></option>
                      	<% if not rs444Tipo22.eof then%>
					<% While NOT rs444Tipo22.EoF %>
                    <option value="<% = rs444Tipo22("tipo") %>">
                    <% =rs444Tipo22("tipo") %>
                    </option>
                    <% rs444Tipo22.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
                    </select>
                      </font></div></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Total</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF"> 
                      <input name="txt_a_total" value="<% if rs("area_total") = "não informado" then response.write "00" else response.write rs("area_total") end if %>" type="text" id="txt_a_total" size="12" maxlength="12" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">m&sup2;</font> 
                      </font></div></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Constru&iacute;da </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <input name="txt_a_constr" value="<% if rs("area_construida") = "não informado" then response.write "00" else response.write rs("area_construida") end if %>" type="text" id="txt_a_constr" size="12" maxlength="12" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>;">
                      <font color="#FFFFFF"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">m&sup2;</font> 
                      </font></div></td>
              </tr>
                <tr > 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_quartos" id="select4" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="<%=rs("quartos")%>" selected><%=rs("quartos")%></option>
                      <option value="não informado">não informado</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                    </select>
                    </div></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_banheiros" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="<%=rs("banheiros")%>" selected><%=rs("banheiros")%></option>
                      <option value="não informado">não informado</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                    </select>
                    </div></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na Garagem</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_vagas" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="<%=rs("vagas")%>" selected><%=rs("vagas")%></option>
                      <option value="não informado">não informado</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                    </select>
                    </div></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_negociacao" id="txt_negociacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <% if rs("negociacao") = "aluguel" then %>
                      <option value="aluguel" selected>Aluguel</option>
                      <option value="venda">Venda</option>
                      <%end if%>
                      <% if rs("negociacao") = "venda" then %>
                      <option value="aluguel">Aluguel</option>
                      <option value="venda" selected>Venda</option>
                      <% end if %>
                    </select>
                    </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    
                    <input name="txt_valor" value="<%=FormatNumber(rs("valor"),2)%>" type="text" id="txt_valor2" size="12" maxlength="12" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>"></td>
              </tr>
             
			  
			  
			   
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                     <option value="<%=rs("ocupacao")%>" selected><%=rs("ocupacao")%></option>
					<option value="não informado" >não informado</option>
                    <option value="vago">vago</option>
                    <option value="ocupado">ocupado</option>
                    
                  </select>
                  </td>
              </tr>
			  
			  
			  <input name="obs_proprietario" value="<%=rs("obs_proprietario")%>" type="hidden" id="obs_proprietario" size="38" maxlength="200" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
			   
              <tr> 
                <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center">
                    <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                            <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&atilde;o 
                                sobre o imóvel</font></div></td>
                      </tr>
                      <tr>
                          <td width="290" height="82" bgcolor="<%=medio%>">&nbsp;</td>
                      </tr>
                    </table>
                  </div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="left">
                        <textarea name="obs_imovel" class="inputBox" id="obs_imovel" style="HEIGHT: 100px; WIDTH: 290px; background: <%=medio%> " onKeyPress="return limitfield(this, 800)"><%=rs("obs_imovel")%></textarea>
                  </div></td>
              </tr>
              <tr> 
                <td><div align="center"></div></td>
                <td><div align="left">
                    <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                          <td><input name="image" type="image"  src="bt_atualizar002.jpg" width="145" height="18" border="0"></td>
                          <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_restaurar001.jpg" width="145" height="18" border="0"></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
</center>
<br>
<%

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
	 varIndicacaoTipo = rs("tipo")
	 varIndicacaoQuartos = rs("quartos")
	 varIndicacaoVagas = rs("vagas")
	 varIndicacaoValor = rs("Valor")





 





  %>
<center>
  <font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
  abaixo indicações de compradores para você.</strong></font> 
</center>
<br>
<table width="740" height="1000" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td><iframe src="indicacao_imoveis02.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>" name="meio" width="740px" height="2030px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>
<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui é criada uma variável "groups" que receberá os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a variável group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero até o número de elementos do array "groups" */

group2[i2]=new Array()
/* aqui é criado o array "group" que receberá valores conforme o número de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receberá valores de opções. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receberá valores de opções. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("200,00 até 500,00","0000000200 0000000500")
group2[2][4]=new Option("500,00 até 1000,00","0000000500 0000001000")
group2[2][5]=new Option("1000,00 até 2000,00","0000001000 0000002000")
group2[2][6]=new Option("Mais de 2000,00","0000002000 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Menos de 20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.000,00 até 50.000,00","0000020000 0000050000")
group2[3][4]=new Option("50.000,00 até 100.000,00","0000050000 0000100000")
group2[3][5]=new Option("100.000,00 até 200.000,00","0000100000 0000200000")
group2[3][6]=new Option("Mais de 200.000,00","0000200000 1000000000")









/* aqui temos um array bidimensional "group" que receberá valores de opções. */


var temp2=document.doublecombo.stage22
/* aqui a variável "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui é criada a função "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que dá um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */


for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que é escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" será o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a variável "location" recebe os valores de "stage2" que corresponde ao endereço de
link para o carregamento de página. */


//-->
</script>


<%else%>

<%  Response.Redirect "acesso03.asp?SecondTry=True&WrongPW=True" %> 

<% end if %>




<%
           rs.Close
           'fecha a conexão
          
           Set rs = Nothing
		   
		   
		   
		   '------------------------------------------------------
		   
		   
		   rs9.Close
           'fecha a conexão
          
           Set rs9 = Nothing
		   
		   
		   
		   '------------------------------------------------------
		   
		   
		   rs444VerificaConta.Close
           'fecha a conexão
          
           Set rs444VerificaConta = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		    '------------------------------------------------------
		   
		   
		   rs2.Close
           'fecha a conexão
          
           Set rs2 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		      '------------------------------------------------------
		   
		   
		   rs4.Close
           'fecha a conexão
          
           Set rs4 = Nothing
		   
		   '----------------------------------------
		   
		   
		      '------------------------------------------------------
		   
		   
		   rs3.Close
           'fecha a conexão
          
           Set rs3 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		      '------------------------------------------------------
		   
		   
		   rs333.Close
           'fecha a conexão
          
           Set rs333 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		      '------------------------------------------------------
		   
		   
		  
          
           Set objfso = Nothing
		   
		   '----------------------------------------
		   
		   
		      '------------------------------------------------------
		   
		   
		   rs444.Close
           'fecha a conexão
          
           Set rs444 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		   
		      '------------------------------------------------------
		   
		   
		   rs555.Close
           'fecha a conexão
          
           Set rs555 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		      '------------------------------------------------------
		   
		   
		   rs444Tipo22.Close
           'fecha a conexão
          
           Set rs444Tipo22 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		      '------------------------------------------------------
		   
		   
		   rsVerifica2.Close
           'fecha a conexão
          
           Set rsVerifica2 = Nothing
		   
		   '----------------------------------------
		   
		   
		   
		   
		   
		   
		   
		   
		   
		   
		   
		   
           %>
 

<% response.flush%>
  <%response.clear%>
  <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>
  
  <%
  
  conexao3.close
  
  set conexao3 = nothing
  
  %>
</body>
</html>

