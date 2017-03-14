<!--#include file="dsn.asp"-->
<!--#include file="style_imoveis.asp"-->
<!--#include file="loggedin.asp"-->
<%
Function EscreveFuncaoJavaScript ( Conexao3 )
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

rsMarcas3.ActiveConnection = Conexao3
	
	
	rsMarcas3.Open SqlMarcas3, Conexao3







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

rsCarros3.ActiveConnection = Conexao3
	
	
	rsCarros3.Open SqlCarros3, Conexao3







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

Sql33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs33.ActiveConnection = Conexao3
	
	
	rs33.Open Sql33, Conexao3




	strSQL44 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 




Set rs44 = Server.CreateObject("ADODB.RecordSet")

	rs44.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs44.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs44.ActiveConnection = Conexao3
	
	
	rs44.Open strSQL44, Conexao3


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
	
	
	rsMarcas333.Open SqlMarcas333, Conexao3








While NOT (rsMarcas333.EOF)

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
While NOT (rsCarros333.EoF)

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




rsMarcas333.Close           
		   
           Set rsMarcas333 = Nothing
             
			rsCarros333.Close           
		   
           Set rsCarros333 = Nothing 




End Function
%> 


<%

'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open dsn

'

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

<%
Function EscreveFuncaoJavaScript2 ( Conexao3 )
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

rsMarcas33.ActiveConnection = Conexao3
	
	
	rsMarcas33.Open SqlMarcas33, Conexao3





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

rsCarros33.ActiveConnection = Conexao3
	
	
	rsCarros33.Open SqlCarros33, Conexao3





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


<%
Function EscreveFuncaoJavaScript999 ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros999 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo4.options[doublecombo.combo4.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas999 = Conexao333.Execute ( SqlMarcas999 )

While NOT (rsMarcas999.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""

Set rsCarros999 = Conexao3.Execute ( SqlCarros999 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo7.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1 
While NOT (rsCarros999.EoF)

Response.Write "doublecombo.combo7.options[" & i & "] = new Option('" & rsCarros999("nome_combo3") & "','" & rsCarros999("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros999.MoveNext
Wend


Response.Write "doublecombo.combo7.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas999.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 





rsMarcas999.Close           
		   
           Set rsMarcas999 = Nothing
             
			rsCarros999.Close           
		   
           Set rsCarros999 = Nothing 




End Function
%> 





<!--#include file="cores03.asp"-->

<% response.buffer=True%>


<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
   dim rs4,strSQL4
   
    
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2 ASC" 
	
	
Set rs4 = Server.CreateObject("ADODB.RecordSet")

	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao3
	
	
	rs4.Open strSQL4, Conexao3
	
	
	
	
	
	
	
	varCod_imovel = request.querystring("varCod_imovel")




if varCod_imovel = "" then
varCod_imovel = "0"
end if
	dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel  FROM imoveis where cod_imovel="&varCod_imovel

   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
	
		
		if rs.eof then
		response.redirect "imovel_negativo01.html"
		end if
		
		
	dim vTarja02
	dim vData01_Tarja02
	dim vData02_Tarja02
	
	
	 vTarja02 = rs("tarja02")
	 vData01_Tarja02 = rs("data01_tarja02")
	 vData02_Tarja02 = rs("data02_tarja02")
	 
	 
	
	
	
dim rs444Placa,strSQL444Placa
   
    
	strSQL444Placa = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	 
	 
	 
Set rs444Placa = Server.CreateObject("ADODB.RecordSet")

	rs444Placa.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Placa.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Placa.ActiveConnection = Conexao
	
	
	rs444Placa.Open strSQL444Placa, Conexao
	 
	 
	 
	 
	 
			
	 
		
		dim rs444Captacao,strSQL444Captacao
   
	strSQL444Captacao = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
			
			 
Set rs444Captacao = Server.CreateObject("ADODB.RecordSet")

	rs444Captacao.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Captacao.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Captacao.ActiveConnection = Conexao
	
	
	rs444Captacao.Open strSQL444Captacao, Conexao
	 	
			
			
			
			
			
			
	
	
		dim rs444Captacao22,strSQL444Captacao22
   
	strSQL444Captacao22 = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
			
			 
Set rs444Captacao22 = Server.CreateObject("ADODB.RecordSet")

	rs444Captacao22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Captacao22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Captacao22.ActiveConnection = Conexao
	
	
	rs444Captacao22.Open strSQL444Captacao22, Conexao		
	
	
	
	
	dim rs444Origem,strSQL444Origem
   
    Set rs444Origem = Server.CreateObject("ADODB.RecordSet")
	strSQL444Origem = "SELECT origem.id_origem,origem.origem FROM origem  ORDER BY id_origem Desc" 
	 
	 
	 
	 rs444Origem.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Origem.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Origem.ActiveConnection = Conexao
	 
	 
	 
	 
	 
	 rs444Origem.Open strSQL444Origem, Conexao
	
	
	
	
	
     dim rs444Responsavel,strSQL444Responsavel
   
    Set rs444Responsavel = Server.CreateObject("ADODB.RecordSet")
	strSQL444Responsavel = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	
	rs444Responsavel.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Responsavel.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Responsavel.ActiveConnection = Conexao
	
	
	
	
	
	 rs444Responsavel.Open strSQL444Responsavel, Conexao

	
	
	
	
	 dim rs444Responsavel22,strSQL444Responsavel22
   
    Set rs444Responsavel22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Responsavel22 = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	rs444Responsavel22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Responsavel22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Responsavel22.ActiveConnection = Conexao
	
	
	
	
	 rs444Responsavel22.Open strSQL444Responsavel22, Conexao
	
	
	
	'------------------------Quem tirou a foto-----------------------
	
	
	
	
	 dim rs445Responsavel22,strSQL445Responsavel22
   
    Set rs445Responsavel22 = Server.CreateObject("ADODB.RecordSet")
	strSQL445Responsavel22 = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	rs445Responsavel22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs445Responsavel22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs445Responsavel22.ActiveConnection = Conexao
	
	
	
	
	 rs445Responsavel22.Open strSQL445Responsavel22, Conexao
	
	
	
	
	
	
	
	
	
	
	
	
	
	'------------------------Quem conseguiu proposta----------------------------------------------
	
	dim rs444ConseguiuProposta22,strSQL444ConseguiuProposta22
   
    Set rs444ConseguiuProposta22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444ConseguiuProposta22 = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	rs444ConseguiuProposta22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444ConseguiuProposta22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444ConseguiuProposta22.ActiveConnection = Conexao
	
	
	
	
	 rs444ConseguiuProposta22.Open strSQL444ConseguiuProposta22, Conexao
	
	
	
	
	
	
	
	
	
	
	'----------------------------------------------------------------------------



'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	
	
	
	rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao
	
	
	
	
	
	 rs444Tipo22.Open strSQL444Tipo22, Conexao







 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC"
	
	
	
	rs444Tipo23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo23.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo23.ActiveConnection = Conexao
	
	
	
	
	 
	 rs444Tipo23.Open strSQL444Tipo23, Conexao




'-------------------------------------------------------------------------------------------------





dim Sql4Bairro,rs4Bairro
	  Set rs4Bairro = Server.CreateObject("ADODB.RecordSet")
Sql4Bairro = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 like '"& rs("bairro") &"' and cidade_combo2 like '"& rs("cidade") &"' ORDER BY nome_combo2" 

        rs4Bairro.CursorLocation = 3
         rs4Bairro.CursorType = 3
           rs4Bairro.ActiveConnection = Conexao

           rs4Bairro.Open Sql4Bairro, Conexao

	
	 
	 
	dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 where nome_combo1 ='"&rs("cidade")&"'  ORDER BY nome_combo1" 
	 
	 rs555.CursorLocation = 3
     rs555.CursorType = 3
     rs555.ActiveConnection = Conexao
	 
	 
	 
	 rs555.Open strSQL555, Conexao  
	 
	
	
	
	dim rs444Vila,strSQL444Vila
   
    Set rs444Vila = Server.CreateObject("ADODB.RecordSet")
	strSQL444Vila = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where nome_combo3 ='"&rs("vila")&"' and bairro_combo3 ='"&rs("bairro")&"' and cidade_combo3 ='"&rs("cidade")&"'   ORDER BY nome_combo3" 
	 
	 
	 rs444Vila.CursorLocation = 3
         rs444Vila.CursorType = 3
           rs444Vila.ActiveConnection = Conexao
	 
	 
	 rs444Vila.Open strSQL444Vila, Conexao
	
	
	
	dim vPerguntaCompradores
	
	vPerguntaCompradores = "não"
	
	dim vPerguntaPermuta
	
	vPerguntaPermuta = "não"
	
	
	
	 
%>	





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE="Javascript">
<!--

//showSubTopNav();
//showSubLeftNav(0, 1);

var popupVisible = false;

function show_info_popup(thisObj,menu_id,vertical_offset) {
	if (popupVisible == false) {
		menuObj = document.getElementById(menu_id);
		position = getAnchorPosition(thisObj.id);
		moveObject(menu_id,position.x+35,position.y - vertical_offset);
		changeObjectVisibility(menu_id,'visible');
		popupVisible = true;
	}
}

function hide_info_popup(thisObj,menu_id) {
	menuObj = document.getElementById(menu_id);
	// moveObject(menu_id,1,1);
	changeObjectVisibility(menu_id,'hidden');
	popupVisible = false;
}

function changeObjectVisibility(objectId, newVisibility) {
    // get a reference to the cross-browser style object and make sure the object exists
    var styleObject = getStyleObject(objectId);
    if(styleObject) {
	styleObject.visibility = newVisibility;
	return true;
    } else {
    	return false;
    }
}

function getStyleObject(objectId) {
     if(document.getElementById(objectId)){
	   return (document.getElementById(objectId).style);
     } else {
	   return false;
     }
}

function moveObject(objectId, newXCoordinate, newYCoordinate) {
    var styleObject = getStyleObject(objectId);
    if(styleObject) {
	styleObject.left = newXCoordinate;
	styleObject.top = newYCoordinate;
    }
}

function getAnchorPosition(anchor_id) {// This function will return an Object with x and y properties
	var position=new Object();
	// Logic to find position
	position.x=AnchorPosition_getPageOffsetLeft(document.getElementById(anchor_id));
	position.y=AnchorPosition_getPageOffsetTop(document.getElementById(anchor_id));
	return position;
}

function AnchorPosition_getPageOffsetLeft (el) {
	var ol=el.offsetLeft;
	while((el=el.offsetParent) != null) {
	  ol += el.offsetLeft;
	}
	return ol;
}

function AnchorPosition_getPageOffsetTop (el) {
	var ot=el.offsetTop;
	while( (el=el.offsetParent) != null) {
	  ot += el.offsetTop;
	}
	return ot;
}
//-->
</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow4354(abrejanela4354) {
   openWindow4354 = window.open(abrejanela4354,'openWin4354','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow4354.focus( )
   }

</SCRIPT>





<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow2 = window.open(abrejanela2,'openWin2','width=800,height=600,resizable=yes,scrollbars=yes')
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

function newWindow3223(abrejanela3223) {
   openWindow3223 = window.open(abrejanela3223,'openWin3223','width=610,height=380,resizable=yes,scrollbars=yes')
   openWindow3223.focus( )
   }

</SCRIPT>




<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow55(abrejanela55) {
   openWindow55 = window.open(abrejanela55,'openWin55','width=603,height=500,resizable=yes,left=200,scrollbars=yes')
   openWindow55.focus( )
   }

</SCRIPT>




<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow77(abrejanela77) {
   openWindow77 = window.open(abrejanela77,'openWin77','width=800,height=550,resizable=yes,left=200,scrollbars=yes')
   openWindow77.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow88(abrejanela88) {
   openWindow88 = window.open(abrejanela88,'openWin88','width=630,height=500,resizable=yes,left=200,scrollbars=yes')
   openWindow88.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow808(abrejanela808) {
   openWindow808 = window.open(abrejanela808,'openWin808','width=610,height=350,resizable=yes,left=200,scrollbars=yes')
   openWindow808.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow44(abrejanela44) {
   openWindow44 = window.open(abrejanela44,'openWin44','width=603,height=500,resizable=yes,left=100,scrollbars=yes')
   openWindow44.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow11(abrejanela11) {
   openWindow11 = window.open(abrejanela11,'openWin11','width=330,height=473,resizable=yes,left=100,scrollbars=yes')
   openWindow11.focus( )
   }

</SCRIPT>





<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3121(abrejanela3121) {
   openWindow3121 = window.open(abrejanela3121,'openWin3121','width=800,height=600,resizable=yes,left=100,scrollbars=yes')
   openWindow3121.focus( )
   }

</SCRIPT>


<script language="JavaScript">
var today=new Date();
var todaysec=today.getSeconds();

function xpop(){
if (confirm("Se você não tem nada a modificar nessa ficha, clique OK, caso tenha, clique em cancelar, modifique e clique em atualizar.")){
window.close();
}
else {
window.open('visualizar_imovel33.asp?varCod_imovel=<%=rs("cod_imovel")%>', todaysec+'floyd','width=800,height=600,resizable=yes,scrollbars=yes,Left=0,Top=0')
}   
}
</script>

<%
if session("permissao") <> "3" and session("permissao") <> "6" then
%>

<script>
function isValidDigitNumber (doublecombo)
{



if ((doublecombo.txt_quartos_vend.value == "00" || doublecombo.txt_quartos_vend.value == "")&&(doublecombo.txt_tipo_vend.value != "TERRENO/ÁREA")) {
        alert("Você precisa indicar o número de quartos que tem o imóvel!");
        doublecombo.txt_quartos_vend.focus();
		
        return false;
    }


if (doublecombo.txt_imovel_em_negociacao_vend.value == "não informado") {
        alert("Você precisa informar se o imóvel está oK!");
        doublecombo.txt_imovel_em_negociacao_vend.focus();
		
        return false;
    }


var strText2_401 = doublecombo.txt_data_futuro_contato_vend.value;
var s_strText2_401 = strText2_401.length
if ((strText2_401.substring((s_strText2_401 - (s_strText2_401 - 3)), (s_strText2_401 - (s_strText2_401 - 2))) != "/") && (strText2_401.substring((s_strText2_401 - (s_strText2_401 - 2)), (s_strText2_401 - (s_strText2_401 - 1))) != "/")){

       alert("A data parece não estar correta!");
       doublecombo.txt_data_futuro_contato_vend.focus();
		
		doublecombo.txt_data_futuro_contato_vend.select();
		
return false;

}





if (doublecombo.txt_imovel_em_negociacao_vend.value == "" || doublecombo.txt_imovel_em_negociacao_vend.value == "não informado") {
        alert(" Você precisa informar se o imóvel está OK ou não!");
        doublecombo.txt_imovel_em_negociacao_vend.focus();
		
        return false;
    }


if (doublecombo.txt_ocupacao_vend.value == "não informado") {
        alert("Você precisa indicar o se o está vago ou não!");
        doublecombo.txt_ocupacao_vend.focus();
		
        return false;
    }


if (doublecombo.txt_proprietario_vend.value == "") {
        alert("Você precisa indicar o nome do proprietário!");
        doublecombo.txt_proprietario_vend.focus();
		doublecombo.txt_proprietario_vend.select();
        return false;
    }
	
	if (doublecombo.txt_pergunta.value == "não informado") {
        alert("Você precisa informar se esse proprietário é também comprador!");
        doublecombo.txt_pergunta.focus();
		
        return false;
    }
	
if (doublecombo.txt_email_vend.value == "" || doublecombo.txt_email_vend.value == "não informado") {
        alert("Você precisa indicar o email do proprietário!");
        doublecombo.txt_email_vend.focus();
		doublecombo.txt_email_vend.select();
        return false;
    }	
	
	
	
if (doublecombo.txt_endereco_vend.value == "" || doublecombo.txt_endereco_vend.value == "não informado" ) {
        alert("Você precisa indicar o endereço do imóvel!");
        doublecombo.txt_endereco_vend.focus();
		doublecombo.txt_endereco_vend.select();
        return false;
    }
	
	
	
	
	
	if (doublecombo.txt_obs_imovel_vend.value == "" || doublecombo.txt_obs_imovel_vend.value == "não informado" ) {
        alert("Você precisa colocar as observações sobre o imóvel e forma de pagamento!");
        doublecombo.txt_obs_imovel_vend.focus();
		doublecombo.txt_obs_imovel_vend.select();
        return false;
    }	
	
	
	
	
	
		
if (doublecombo.txt_data_futuro_contato_vend.value == "0/0/2007 00:00:00") {
        alert("Você precisa indicar a data para o futuro contato !");
        doublecombo.txt_data_futuro_contato_vend.focus();
		doublecombo.txt_data_futuro_contato_vend.select();
        return false;
    }	
	
	
	
	
	
	
	
	
	
	
	
	
	
	//----------txt_obs_imovel_vend
	
	
	
	if (doublecombo.txt_valor_vend.value == "0,00") {
        alert("Você precisa indicar um valor para o imóvel!");
        doublecombo.txt_valor_vend.focus();
		doublecombo.txt_valor_vend.select();
        return false;
    }
	
	
	
	
	if (doublecombo.txt_valor_vend.value == "") {
        alert("Você precisa indicar um valor para o imóvel!");
        doublecombo.txt_valor_vend.focus();
		doublecombo.txt_valor_vend.select();
        return false;
    }
	
	
	
	
	if (doublecombo.txt_tipo_vend.value == "APARTAMENTO" && doublecombo.txt_condominio_vend.value == "0,00" ) {
        alert("Você precisa indicar o valor do condomínio!");
        doublecombo.txt_condominio_vend.focus();
		doublecombo.txt_condominio_vend.select();
        return false;
    }
	
	
	
	
	
	
	if (doublecombo.txt_tipo_vend.value == "COBERTURA" && doublecombo.txt_condominio_vend.value == "0,00" ) {
        alert("Você precisa indicar o valor do condomínio!");
        doublecombo.txt_condominio_vend.focus();
		doublecombo.txt_condominio_vend.select();
        return false;
    }
	
	
	if (doublecombo.txt_tipo_vend.value == "CASAS EM CONDOMINIO" && doublecombo.txt_condominio_vend.value == "0,00" ) {
        alert("Você precisa indicar o valor do condomínio!");
        doublecombo.txt_condominio_vend.focus();
		doublecombo.txt_condominio_vend.select();
        return false;
    }
	
	


if (doublecombo.txt_tipo_vend.value == "APARTAMENTO" && doublecombo.txt_nome_edificio_vend.value == "" ) {
        alert("Você precisa o nome do edifício!");
        doublecombo.txt_nome_edificio_vend.focus();
		doublecombo.txt_nome_edificio_vend.select();
        return false;
    }


	if (doublecombo.txt_tipo_vend.value == "COBERTURA" && doublecombo.txt_nome_edificio_vend.value == "" ) {
        alert("Você precisa o nome do edifício!");
        doublecombo.txt_nome_edificio_vend.focus();
		doublecombo.txt_nome_edificio_vend.select();
        return false;
    }
	
	
	if (doublecombo.txt_tipo_vend.value == "CASAS EM CONDOMINIO" && doublecombo.txt_nome_edificio_vend.value == "" ) {
        alert("Você precisa o nome do edifício!");
        doublecombo.txt_nome_edificio_vend.focus();
		doublecombo.txt_nome_edificio_vend.select();
        return false;
    }
	
	
	
var strValidNumber1_7="1234567890";
for (nCount=0; nCount < doublecombo.txt_telefone_vend.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_telefone_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O telefone do proprietário só pode conter números!");
doublecombo.txt_telefone_vend.focus();
doublecombo.txt_telefone_vend.select();
return false;
}
}



	
var strValidNumber1_total="1234567890";
for (nCount=0; nCount < doublecombo.txt_a_total_vend.value.length; nCount++) 
		{
strTempChar1_total=doublecombo.txt_a_total_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_total.indexOf(strTempChar1_total,0)==-1) 
{
alert("A área total só pode conter números!");
doublecombo.txt_a_total_vend.focus();
doublecombo.txt_a_total_vend.select();
return false;
}
}








	var strValidNumber1_6="1234567890,.";
for (nCount=0; nCount < doublecombo.stage22.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.stage22.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.stage22.focus();
doublecombo.stage22.select();
return false;
}
}


var strValidNumber1_7="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_valor_vend.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_valor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor_vend.focus();
doublecombo.txt_valor_vend.select();
return false;
}
}








//------------------------------saldo devedor--------------------------







if (doublecombo.txt_ja_pago_devedor_vend.value == "") {
        alert("O formulário valor já pago no saldo devedor está vazio!");
        doublecombo.txt_ja_pago_devedor_vend.focus();
		doublecombo.txt_ja_pago_devedor_vend.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_ja="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_ja_pago_devedor_vend.value.length; nCount++) 
		{
strTempChar1_ja=doublecombo.txt_ja_pago_devedor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_ja.indexOf(strTempChar1_ja,0)==-1) 
{
alert("O formulário valor já pago no saldo devedor só pode conter números!");
doublecombo.txt_ja_pago_devedor_vend.focus();
doublecombo.txt_ja_pago_devedor_vend.select();
return false;
}
}






if (doublecombo.txt_devendo_devedor_vend.value == "") {
        alert("O formulário valor devido no saldo devedor está vazio!");
        doublecombo.txt_devendo_devedor_vend.focus();
		doublecombo.txt_devendo_devedor_vend.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_devendo="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_devendo_devedor_vend.value.length; nCount++) 
		{
strTempChar1_devendo=doublecombo.txt_devendo_devedor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_devendo.indexOf(strTempChar1_devendo,0)==-1) 
{
alert("O formulário valor devido no saldo devedor só pode conter números!");
doublecombo.txt_devendo_devedor_vend.focus();
doublecombo.txt_devendo_devedor_vend.select();
return false;
}
}














//------




//------------------------------Área Total--------------------------







if (doublecombo.txt_a_total_vend.value == "") {
        alert("O formulário área total  está vazio!");
        doublecombo.txt_a_total_vend.focus();
		doublecombo.txt_a_total_vend.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_atotal="1234567890";
for (nCount=0; nCount < doublecombo.txt_a_total_vend.value.length; nCount++) 
		{
strTempChar1_atotal=doublecombo.txt_a_total_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_atotal.indexOf(strTempChar1_atotal,0)==-1) 
{
alert("O formulário área total só pode conter números!");
doublecombo.txt_a_total_vend.focus();
doublecombo.txt_a_total_vend.select();
return false;
}
}






if (doublecombo.txt_a_total_vend.value == "00") {
        alert("O formulário área total precisa de um valor!");
        doublecombo.txt_a_total_vend.focus();
		doublecombo.txt_a_total_vend.select();
        return false;
    }



if (doublecombo.txt_titulo_anuncio_vend.value == "") {
        alert("O você precisa colocar um título no anúncio!");
        doublecombo.txt_titulo_anuncio_vend.focus();
		doublecombo.txt_titulo_anuncio_vend.select();
        return false;
    }
	
	
	if (doublecombo.txt_texto_anuncio_vend.value == "") {
        alert("O você precisa colocar o texto do anúncio!");
        doublecombo.txt_texto_anuncio_vend.focus();
		doublecombo.txt_texto_anuncio_vend.select();
        return false;
    }
	
	
	
	








	

	
	
	



//------


















var strValidNumber1_8="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_condominio_vend.value.length; nCount++) 
		{
strTempChar1_8=doublecombo.txt_condominio_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_8.indexOf(strTempChar1_8,0)==-1) 
{
alert("O formulário Condomínio só pode conter números!");
doublecombo.txt_condominio_vend.focus();
doublecombo.txt_condominio_vend.select();
return false;
}
}





	
	if (doublecombo.stage22.value == "" && doublecombo.txt_pergunta.value == "sim" ) {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }
	
	
	
	if (doublecombo.stage22.value == "0,00" && doublecombo.txt_pergunta.value == "sim" ) {
        alert("Você precisa indicar uma valor para o imóvel pretendido");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }
	
	
	
	
	if (doublecombo.txt_tipo_vend.value == "Apartamento" && doublecombo.txt_condominio_vend.value == "0,00") {
        alert("Você precisa indicar o valor do condomínio!");
        doublecombo.txt_condominio_vend.focus();
		doublecombo.txt_condominio_vend.select();
        return false;
    }
	
	if (doublecombo.txt_condominio_vend.value == "" ) {
        alert("O item condomínio do imóvel está vazio!");
        doublecombo.txt_condominio_vend.focus();
		doublecombo.txt_condominio_vend.select();
        return false;
    }
	
	
	
	
	//------------Compradores condominio,área total e área construida----------------------

var strValidNumber8_ja="1234567890";
for (nCount=0; nCount < doublecombo.txt_area_total.value.length; nCount++) 
		{
strTempChar8_ja=doublecombo.txt_area_total.value.substring(nCount,nCount+1);
if (strValidNumber8_ja.indexOf(strTempChar8_ja,0)==-1 && doublecombo.txt_pergunta.value == "sim") 
{
alert("O item área total só pode ter números!");
doublecombo.txt_area_total.focus();
doublecombo.txt_area_total.select();
return false;
}
}





var strValidNumber9_ja="1234567890";
for (nCount=0; nCount < doublecombo.txt_area_construida.value.length; nCount++) 
		{
strTempChar9_ja=doublecombo.txt_area_construida.value.substring(nCount,nCount+1);
if (strValidNumber9_ja.indexOf(strTempChar9_ja,0)==-1 && doublecombo.txt_pergunta.value == "sim") 
{
alert("O item área construida só pode ter números!");
doublecombo.txt_area_construida.focus();
doublecombo.txt_area_construida.select();
return false;
}
}


var strValidNumber99_ja="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_condominio.value.length; nCount++) 
		{
strTempChar99_ja=doublecombo.txt_condominio.value.substring(nCount,nCount+1);
if (strValidNumber99_ja.indexOf(strTempChar99_ja,0)==-1 && doublecombo.txt_pergunta.value == "sim") 
{
alert("O item condomínio só pode conter números e uma vírgula!");
doublecombo.txt_condominio.focus();
doublecombo.txt_condominio.select();
return false;
}
}
	
	
	
	
	
	


var strText2_4 = doublecombo.stage22.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != "," && doublecombo.txt_pergunta.value == "sim"){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.stage22.focus();
		
		doublecombo.stage22.select();
		
return false;

}





var strText2_401 = doublecombo.txt_data_futuro_contato_comprador.value;
var s_strText2_401 = strText2_401.length
if ((strText2_401.substring((s_strText2_401 - (s_strText2_401 - 3)), (s_strText2_401 - (s_strText2_401 - 2))) != "/") && (strText2_401.substring((s_strText2_401 - (s_strText2_401 - 2)), (s_strText2_401 - (s_strText2_401 - 1))) != "/") && doublecombo.txt_pergunta.value == "sim"){

       alert("A data parece não estar correta!");
       doublecombo.txt_data_futuro_contato_comprador.focus();
		
		doublecombo.txt_data_futuro_contato_comprador.select();
		
return false;

}


if (doublecombo.txt_origem.value == "não informado" && doublecombo.txt_pergunta.value == "sim") {
        alert("Você precisa informar qual a origem desse comprador!");
        doublecombo.txt_origem.focus();
		
        return false;
    }







if (doublecombo.combo1.value == "cqualquer" && doublecombo.txt_pergunta.value == "sim") {
        alert("Você precisa indicar em qual cidade o cliente está interessado!");
        doublecombo.combo1.focus();
		
        return false;
    }
	
	
	
	if ((doublecombo.combo2.value == "bqualquer" || doublecombo.combo2.value == "")&&(doublecombo.txt_pergunta.value == "sim")) {
        alert("Você precisa indicar em qual bairro o cliente está interessado!");
        doublecombo.combo2.focus();
		
        return false;
    }
	
	
	if ((doublecombo.txt_tipo.value == "tqualquer" || doublecombo.txt_tipo.value == "")&&(doublecombo.txt_pergunta.value == "sim")) {
        alert("Você precisa indicar o tipo de imóvel que o cliente quer!");
        doublecombo.txt_tipo.focus();
		
        return false;
    }



if ((doublecombo.txt_quartos.value == "00" || doublecombo.txt_quartos.value == "")&&(doublecombo.txt_pergunta.value == "sim")) {
        alert("Você precisa indicar o número de quartos que o cliente quer!");
        doublecombo.txt_quartos.focus();
		
        return false;
    }


if ((doublecombo.txt_vagas.value == "00" || doublecombo.txt_vagas.value == "")&&(doublecombo.txt_pergunta.value == "sim")) {
        alert("Você precisa indicar o número de vagas que o cliente quer!");
        doublecombo.txt_vagas.focus();
		
        return false;
    }
	
	
	

if (doublecombo.example2.value == "" && doublecombo.txt_pergunta.value == "sim") {
        alert("Você precisa indicar qual negociação o cliente quer!");
        doublecombo.example2.focus();
		
        return false;
    }
	
	
	
	if ((doublecombo.txt_standby.value == "não informado" || doublecombo.txt_standby.value == "")&&(doublecombo.txt_pergunta.value == "sim")) {
        alert("Você precisa indicar a situação do cliente!");
        doublecombo.txt_standby.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_descricao.value == "" && doublecombo.txt_pergunta.value == "sim") {
        alert("Você precisa fazer observações sobre esse cliente !");
        doublecombo.txt_descricao.focus();
		
        return false;
    }
	
	
	
	if ((doublecombo.txt_condicoes_pagamento.value == "" || doublecombo.txt_condicoes_pagamento.value == "não informado")&&(doublecombo.txt_pergunta.value == "sim")) {
        alert("Você precisa informar as condições de pagamento para  esse cliente !");
        doublecombo.txt_condicoes_pagamento.focus();
		
        return false;
    }
	
	if (doublecombo.txt_condominio.value == "" && doublecombo.txt_pergunta.value == "sim" ) {
        alert("O item condomínio do imóvel pretendido está vazio!");
        doublecombo.txt_condominio.focus();
		doublecombo.txt_condominio.select();
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











}



</script>
<%
else


end if

%>


</head>

<body bgcolor="<%=escuro%>" >


<%



varCod_imovel = rs("cod_imovel")

%>
 <div align="center"><%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
        <%else%>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
        foi atualizado com sucesso.</font> 
        <% end if %>
		</div>
		<br>
		<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="400" height="20">
              <%
				   dim varRs444Permuta
	if rs("cod_permuta") <> "" then
	varRs444Permuta = rs("cod_permuta")
	else
	varRs444Permuta = "0"
	end if
				   
				   
				   
				   dim rs444Permuta,SQL444Permuta
 Set rs444Permuta = Server.CreateObject("ADODB.RecordSet")
 SQL444Permuta = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone like'"& rs("telefone")&"' order by cod_permuta DESC" 
	
	
	rs444Permuta.CursorLocation = 3
         rs444Permuta.CursorType = 3
           rs444Permuta.ActiveConnection = Conexao
	
	
	rs444Permuta.open SQL444Permuta,Conexao,2,1  
	
	dim varCod_Permuta006
	
			
	if  not rs444Permuta.eof then
	
	varCod_Permuta006 = rs444Permuta("cod_permuta") 
	
	vPerguntaPermuta = "sim"
				  
				 'while not rs444Permuta.eof 
				  %>
              <div align="left"><a href="javascript:newWindow55('visualizar_permuta33.asp?varCodPermuta=<%=rs444Permuta("cod_permuta")%>')"><img src="bt_foto22perm.jpg" width="290" height="18" border="0" align="middle" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rs444Permuta("cod_permuta")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs444Permuta("cod_permuta")%>')"></a>
               
			     
	  
	   <DIV STYLE="border: #000000 1px solid;  width: 270px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 1px; right: 1px;" CLASS="smalltext" ID="<%=rs444Permuta("cod_permuta")%>">
	   
	   <table width="580" border="0" cellspacing="0" cellpadding="0">
               <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Última atualização</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs444Permuta("data_atualizacao")%></font></td>
              </tr>
			   
			   
			   
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      da permuta</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs444Permuta("cod_permuta")%></font></td>
              </tr>
			 
			 
			 
			 
			    <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      nome &eacute;:</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_nome" value="<%=rs444Permuta("nome")%>" type="text" id="txt_nome" size="38" maxlength="200" align="left" class="inputBox" style="border-color:#FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>; "></td>
              </tr>
              
			 
              
			  
			  
             
			  
			  
			  
                
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atualmente 
                      tenho um im&oacute;vel na cidade de:</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("cidade_vend") = "cqualquer" then response.write "não informado" else response.write rs444Permuta("cidade_vend") end if %>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      bairro: </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("bairro_vend") = "bqualquer" then response.write "não informado" else response.write rs444Permuta("bairro_vend") end if %>
                    </font></td>
              </tr>
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Na 
                      vila: </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("vila_vend") = "vlqualquer" then response.write "não informado" else response.write rs444Permuta("vila_vend") end if %>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">do 
                      tipo: </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("tipo_vend") = "tqualquer" then response.write "não informado" else response.write rs444Permuta("tipo_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("quartos_vend") = "qqualquer" then response.write "não informado" else response.write rs444Permuta("quartos_vend") end if %>
                    </font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de vagas na garagem</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("vagas_vend") = "vgqualquer" then response.write "não informado" else response.write rs444Permuta("vagas_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      valor de</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("valor_vend") = "vqualquer" then response.write "não informado" else response.write FormatNumber(rs444Permuta("valor_vend"),2) end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      im&oacute;vel tem a seguinte descri&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> <textarea name="textarea" class="inputBox" id="textarea"  style="border-color:#FFFFFF;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Permuta("descricao_vend")%></textarea></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Pretendo 
                      morar na cidade de:</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("cidade_comp") = "cqualquer" then response.write "não informado" else  response.write rs444Permuta("cidade_comp") end if %>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      bairro: </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("bairro_comp") = "bqualquer" then response.write "não informado" else  response.write rs444Permuta("bairro_comp") end if %>
                    </font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Na 
                     vila: </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("vila_comp") = "vlqualquer" then response.write "não informado" else  response.write rs444Permuta("vila_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quero 
                      trocar por um im&oacute;vel do tipo:</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp; 
                    </font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("tipo_comp") = "tqualquer" then response.write "não informado" else  response.write rs444Permuta("tipo_comp") end if %>
                    </font></td>
              </tr>
                
				
				 
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("quartos_comp") = "qqualquer" then response.write "não informado" else  response.write rs444Permuta("quartos_comp") end if %>
                    </font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de vagas na garagem</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("vagas_comp") = "vgqualquer" then response.write "não informado" else  response.write rs444Permuta("vagas_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      valor de</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs444Permuta("valor_comp") = "vqualquer" then response.write "não informado" else  response.write FormatNumber(rs444Permuta("valor_comp"),2) end if %>
                    </font></td>
              </tr>
				
				
				
              <tr>
                  <td width="290" bgcolor="<%=medio%>" height="100" style="border:1px solid #FFFFFF;" >
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Que 
                      tenha a seguinte descri&ccedil;&atilde;o</font></div></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao2" class="inputBox" id="txt_descricao2"  style="border-color:#FFFFFF;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Permuta("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
              </tr>
            </table>
	   
	   
	   
	   
	   
	   
	   
	   
	   </DIV>
	   
	   
	   
			   
			   
			   
			    <%
				  ' rs444Permuta.movenext
				  ' wend
				   %>
                <%else%>
				
				<%
				varCod_Permuta006 = "0"
				%>
              </div>
              <%end if%></td>
    <td width="200" height="20"> 
              <div align="center"><font color="<%=letra%>" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>Im&oacute;vel</strong></font><br>
        <font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><br>
        </font></div></td>
    <td width="400" height="20">
              <%
				   dim varRs444Comprador
	if rs("cod_comprador") <> "" then
	varRs444Comprador = rs("cod_comprador")
	else
	varRs444Comprador = "0"
	end if
				   
				   
				   
				   dim rs444Comprador,SQL444Comprador
 Set rs444Comprador = Server.CreateObject("ADODB.RecordSet")
 SQL444Comprador = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta  FROM compradores where  telefone like'%"& rs("telefone")&"%' or telefone02 like'%"& rs("telefone")&"%' or telefone03 like'%"& rs("telefone")&"%'  order by cod_compradores DESC" 
	
	
	rs444Comprador.CursorLocation = 3
         rs444Comprador.CursorType = 3
           rs444Comprador.ActiveConnection = Conexao
	
	
	rs444Comprador.open SQL444Comprador,Conexao,2,1  
	
	dim CidadeCompradorFicha01
	dim BairroCompradorFicha01
	
	
	
	dim varCod_compradores006
	
	
	dim semComprador
	
	semComprador = "notNothing"
			
	if  not rs444Comprador.eof then
	
	varCod_compradores006 = rs444Comprador("cod_compradores")
	
				 vPerguntaCompradores = "sim"  
				
				CidadeCompradorFicha01 = rs444Comprador("cidade")
				BairroCompradorFicha01 = rs444Comprador("bairro")
				  
				 ' while not rs444Comprador.eof
				  
				  %>
              <div align="right"><a href="javascript:newWindow44('visualizar_compradores33.asp?varCodCompradores=<%=rs444Comprador("cod_compradores")%>')"><img src="bt_foto22Compr.jpg" width="290" height="18" border="0" align="middle" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rs444Comprador("cod_compradores")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs444Comprador("cod_compradores")%>')"></a>
               
			   
			    <DIV STYLE="border: #000000 1px solid;  width: 570px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 200px; right: 200px;" CLASS="smalltext" ID="<%=rs444Comprador("cod_compradores")%>">
		
		
		<table width="580" border="0" cellspacing="0" cellpadding="0">
                
             <tr>
                        <td height="30" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente</strong></font></div></td>
                        <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444Comprador("atendimento")%></strong></font></td>
              </tr>
			 
			 
			 
			 <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                            do comprador</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("cod_compradores")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Uacute;tima 
                            atualiza&ccedil;&atilde;o </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="border-color #FFFFFF;color:#FFFFFF;;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("data_atualizacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			 
			 
			 
			 
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      nome </font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="border-color : #FFFFFF;color:#FFFFFF;;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("nome")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("telefone")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                
			  
			 
              
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("email")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
              
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("endereco")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      que estou interessado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_cidade" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("cidade")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      que estou interessadol</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_bairro" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("bairro")%>" size="38" maxlength="33" align="left">
                  </td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
                    <input name="txt_tipo" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%if rs444Comprador("Tipo") <> "tqualquer" then response.write rs444Comprador("Tipo") else response.write "qualquer um" end if  %>" size="38" maxlength="33" align="left">
                    </td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de quartos do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("quartos")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                  de vagas do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Comprador("vagas")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      que eu quero</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Comprador("negociacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Faixa 
                      de pre&ccedil;o que eu quero</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value=" <%if vValor <> "vqualquer" then%><%=FormatNumber(rs444Comprador("Valor"),2)%> <%else%>não informado<%end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
			 
			                           
			 
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Aqui 
                            tem a descri&ccedil;&atilde;o do im&oacute;vel que 
                            eu quero</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=claro%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="border-color : #FFFFFF;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Comprador("descricao")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="145">&nbsp; </td>
                        <td width="145">&nbsp;</td>
                      </tr>
                    </table></td>
              </tr>
            </table>
		
</DIV>
			   
			   
			   
			   
			   
			   
			   
			   
			    <%
					
					'rs444Comprador.movenext
					'wend
					
					
					%>
                <%else%>
				
				
				<%
				
				
				varCod_compradores006 = "0"
				
				
				
				
				semComprador = "nothing"
				
				
				CidadeCompradorFicha01 = "não informado"
				BairroCompradorFicha01 = "não informado"
				%>
              </div>
              <%end if%></td>
  </tr>
  <tr>
  <td></td>
    <td> 
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow4354('indicacao_imoveis22.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>')" style="color:<%=letra%>"> 
        </a></font></div></td>
  <td></td>
  </tr>
</table>

		
		
		<br>
		
		
<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="166"> 
      <div align="center"><font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow11('simulador01.asp')" style="color:<%=letra%>">simulador 
        de financiamento</a></font></div></td>
    <td width="166"> 
      <div align="center"><font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("imovel_em_negociacao") = "Vendido pela Veja" then %> <a href="javascript:newWindow77('form_venda_realizada01.asp?varCod_imovel=<%=varCod_imovel%>')" style="color:<%=letra%>">Dados da venda realizada</a><%else%><%end if%>
      </font></div></td>
    <td width="166"> 
      <div align="center"><font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow88('imprimir_imovel22.asp?varCod_imovel=<%=varCod_imovel%>')" style="color:<%=letra%>">Imprimir</a></font></div></td>
    <td width="166"> 
      <div align="center"><font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><% if session("permissao") = "6"  then %> <a href="javascript:newWindow808('visualizar_historico33.asp?varCod_imovel=<%=varCod_imovel%>')" style="color:<%=letra%>">Histórico de atualizações</a><%else%><%end if%></font></div></td>
    <td width="166"> 
      <div align="center"><font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow3121('form_duplicar_imovel01.asp?varCod_imovel=<%=varCod_imovel%>')" style="color:<%=letra%>">Duplicar</a></font></div></td>
     <td width="166"> 
      <div align="center"><font size="2" color="<%=letra%>" face="Verdana, Arial, Helvetica, sans-serif"><%
     dim varIndicacaoCidade 
	 dim varIndicacaoBairro 
	 dim varIndicacaoNegociacao  
	 dim varIndicacaoTipo  
	 dim varIndicacaoQuartos  
	 dim varIndicacaoVagas 
	 dim varIndicacaoValor 
	 
	 
	 
	  varIndicacaoCidade = rs("cidade")
	 varIndicacaoBairro = rs("bairro")
	 varIndicacaoNegociacao = rs("negociacao")
	 varIndicacaoTipo = rs("tipo")
	 varIndicacaoQuartos = rs("quartos")
	 varIndicacaoVagas = rs("vagas")
	 varIndicacaoValor = rs("Valor") 
  
  
  %>
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow4354('indicacao_imoveis22.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>')" style="color:<%=letra%>"> 
       Indicações</a></font> </div></font></div></td>
    
  </tr>
</table>

		
		
		
		<br>
		
		
<form name="doublecombo" target="_self"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_imovel33.asp?varCod_imovel=<%=varCod_imovel%>&vPerguntaPermuta=<%=vPerguntaPermuta%>&vPerguntaCompradores=<%=vPerguntaCompradores%>&varCod_permuta006=<%=varCod_permuta006%>&varCod_compradores006=<%=varCod_compradores006%>">
<table width="1000" border="0" cellspacing="0" cellpadding="0">
 
 
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        
		
		
		<tr> 
          <td width="192"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
            futuro contato</font></td>
          <td width="10">&nbsp;</td>
          <td width="762"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto 
            futuro contato</font></td>
          <td width="10">&nbsp;</td>
          <td width="26"> <a href="javascript:newWindow2121('visualizar_fotos.asp?varCodimovel=<%=varCod_imovel%>')"><img src="bt_mais03.jpg" width="18" height="18" border="0"></a></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
		  <td width="26">
		<% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
        
		
		
		 
		 
		 
		<div align="center"><a href="javascript:newWindow3223('mostrar_imovel03.asp?varCod_imovel=<%=varCod_imovel%>')"><IMG SRC="icon_foto.gif" border="0" align="middle" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rs("cod_imovel")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs("cod_imovel")%>')"></img></a></div>
       
	   
	    <DIV STYLE="border: #000000 1px solid;  width: 270px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 1px; right: 1px;" CLASS="smalltext" ID="<%=rs("cod_imovel")%>"><img src="<%=rs("foto_grande")%>" width="270" height="157"></DIV>
	   
	   
	 
	   
	   
	
	   <%else%>
	   
	   <IMG SRC="icon_foto.gif" border="0" align="middle" ></img>
	   <%end if%>
		  
		  
		  <td width="10">&nbsp;</td>
          <td width="154" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
		  
		  
		  <%
		  if (session("nome_id") = rs("captacao") or session("permissao") = "6" )then 
		  %>
		  <input name="txt_data_futuro_contato_vend" type="text" class="inputBox" id="txt_data_futuro_contato_vend" style="HEIGHT: 18px; WIDTH: 154px; background:<%=medio%>" value="<%if (rs("data_futuro_contato") <> "não informado" and  rs("data_futuro_contato") <> "" )   then response.write rs("data_futuro_contato") else response.write rs("data") end if%>" size="38" maxlength="50" align="left"></td>
         
		  <% else%>
		  <font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if (rs("data_futuro_contato") <> "não informado" and  rs("data_futuro_contato") <> "" )   then response.write rs("data_futuro_contato") else response.write rs("data") end if%></font>
		   <input name="txt_data_futuro_contato_vend" type="hidden" class="inputBox" id="txt_data_futuro_contato_vend" style="HEIGHT: 18px; WIDTH: 154px; background:<%=medio%>" value="<%if (rs("data_futuro_contato") <> "não informado" and  rs("data_futuro_contato") <> "" )   then response.write rs("data_futuro_contato") else response.write rs("data") end if%>" size="38" maxlength="50" align="left"></td>
         
		  <%end if%>

<td width="10">&nbsp;</td>
          <td width="798" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_assunto_futuro_contato_vend" type="text" class="inputBox" id="txt_assunto_futuro_contato_vend" style="HEIGHT: 18px; WIDTH: 796px; background:<%=medio%>" value="<%if rs("assunto_futuro_contato") <> "" then response.write rs("assunto_futuro_contato") else response.write "não informado" end if%>" size="38" maxlength="50" align="left"></td>
          
         
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td> <table width="1000" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Código 
            do imóvel</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
            capta&ccedil;&atilde;o </font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Origem 
            Capta&ccedil;&atilde;o </font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="Input2" type="text" class="inputBox" id="Input" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("cod_imovel")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
              <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                 <option value="<%=rs("tipo")%>" selected><%if   rs("tipo") = "tqualquer" then response.write "não informado" else response.write rs("tipo") end if  %></option>					 
					 
				
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
              </font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_negociacao_vend" id="txt_negociacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                 
                     <option value="<%=rs("negociacao")%>" selected><%if rs("negociacao") = "nqualquer" then response.write "qualquer uma" else response.write rs("negociacao") end if%></option>
					 
				
				<option value="aluguel">aluguel</option>
                <option value="venda" >venda</option>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_captacao_vend" type="text" class="inputBox" id="txt_data_captacao_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("data_captacao")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_origem_captacao_vend" value="<%if rs("origem_captacao") <> "" then response.write rs("origem_captacao") else response.write "não informado" end if%>" type="text" id="txt_origem_captacao_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>"></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Capta&ccedil;&atilde;o 
              </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">imovel 
              atualizado por</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
            inclus&atilde;o </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><table width="192" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="20"><img src="pisca01.gif" width="20" height="20"></td>
                  <td width="172"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                    do &uacute;ltimo contato</font></td>
                </tr>
              </table></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
            cadastramento </font></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
				
				<select name="txt_captacao_vend" id="txt_captacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
               
			   
			   <% if UCase(rs("captacao")) = UCase("internet") or UCase(rs("captacao")) = UCase("Avaliação") or UCase(rs("captacao")) = UCase("Busca de compradores") or  session("permissao") = "6" then %>
			   
			    <option value="<%=rs("captacao")%>" selected ><%=rs("captacao")%></option>
					 <option value="internet">internet</option>
					 
                      <% if not rs444Captacao.eof then %>
                      <% While NOT rs444Captacao.EoF %>
                      <option value="<% = rs444Captacao("list_name") %>"> 
                      <% = rs444Captacao("list_name") %>
                      </option>
                      <% rs444Captacao.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Internet">Internet</option>
                      <%end if%>
					  
					  <%else%>
					  
					  <option value="<%=rs("captacao")%>"><%=rs("captacao")%></option>
					  
					  <%end if%>
                    </select>
					
					
					
					</td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="Input" type="text" value="<%if rs("quem_atualizou") <> "" then response.write rs("quem_atualizou") else response.write "não informado" end if%>"  size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input type="text" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("data")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input type="text" class="inputBox" id="txt_data_captacao3" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("data_atualizacao")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_responsavel_cadastramento" id="txt_responsavel_cadastramento" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                
				<option value="<%if rs("responsavel_cadastramento") <> "" then response.write rs("responsavel_cadastramento") else response.write "não informado" end if%>"><%if rs("responsavel_cadastramento") <> "" then response.write rs("responsavel_cadastramento") else response.write "não informado" end if%></option>
				<option value="Internet">Internet</option>
                <% if not rs444Responsavel22.eof then %>
                <% While NOT rs444Responsavel22.EoF %>
                <option value="<% = rs444Responsavel22("list_name") %>"> 
                <% = rs444Responsavel22("list_name") %>
                </option>
                <% rs444Responsavel22.MoveNext %>
                <% Wend %>
                <%else%>
                <option value="Internet">Internet</option>
                <%end if%>
              </select></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Propriet&aacute;rio</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
              01</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
              02</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
              03</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario_vend" type="text" class="inputBox" id="txt_proprietario_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("proprietario")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone_vend" type="text" class="inputBox" id="txt_telefone_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%=rs("telefone")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone03_vend" type="text" class="inputBox" id="txt_telefone03_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs("telefone03") <> "" then response.write rs("telefone03") else response.write "não informado" end if%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone02_vend" type="text" class="inputBox" id="txt_telefone02_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs("telefone02") <> "" then response.write rs("telefone02") else response.write "não informado" end if%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email_vend" type="text" class="inputBox" id="txt_email_vend" style="HEIGHT: 18px; WIDTH: 190px ; background:<%=medio%>" value="<%=rs("email")%>" size="38" maxlength="50" align="left"></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Local 
              das chaves e n&uacute;mero</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Qualidade</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Hor&aacute;rio 
            de visita</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Placa</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_chaves_do_imovel_vend" type="text" class="inputBox" id="txt_chaves_do_imovel_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%if rs("chaves_do_imovel") <> "" then response.write rs("chaves_do_imovel") else response.write "não informado" end if%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_qualidade_vend" id="txt_qualidade_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
               
			   <option value="<%if rs("qualidade") <> "" then response.write rs("qualidade") else  response.write "não informado" end if%>" selected><%if rs("qualidade") <> "" then response.write rs("qualidade") else  response.write "não informado" end if%></option>
			    <option value="bom negócio" >Bom Negócio</option>
                <option value="negócio comum">Negócio Comum</option>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_melhor_horario_visita_vend" size="1" class="inputBox" id="txt_melhor_horario_visita_vend"  style="HEIGHT: 18px; WIDTH: 190px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
                <option value="<%if rs("melhor_horario_visita") <> "" then response.write rs("melhor_horario_visita") else response.write "não informado" end if%>" selected><%if rs("melhor_horario_visita") <> "" then response.write rs("melhor_horario_visita") else response.write "não informado" end if%></option>
                 <option value="Ligar antes">Ligar antes</option>    
			    <option  value="Manhâ">Manhã </option>
                <option value="Tarde" >Tarde </option>
                <option  value="Noite">Noite </option>
                <option value="Manhã ou tarde" >Manhã ou tarde </option>
                <option  value="Manhã ou noite">Manhã ou noite</option>
                <option value="Tarde ou noite" >Tarde ou noite </option>
                <option value="Qualquer horário">Qualquer horário</option>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_placa_vend" id="txt_placa_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                
				<option value="<%if rs("placa") <> "" then %><%=rs("placa")%><%else%><%="Sem Placa"%><%end if%>" select><%if rs("placa") <> "" then %><%=rs("placa")%><%else%><%="Sem Placa"%><%end if%></option>
					   
				<option value="Sem Placa" >Sem Placa</option>
                <% if not rs444Placa.eof then %>
                <% While NOT rs444Placa.EoF %>
                <option value="<% = rs444Placa("list_name") %>"> 
                <% = rs444Placa("list_name") %>
                </option>
                <% rs444Placa.MoveNext %>
                <% Wend %>
                <%else%>
                <option value="Sem Placa">Sem Placa</option>
                <%end if%>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_ocupacao_vend" id="txt_ocupacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
               
			    <option value="<%=rs("ocupacao")%>" selected><%=rs("ocupacao")%></option>
					
			    <option value="não informado">não informado</option>
                <option value="vago">vago</option>
                <option value="alugado">Alugado</option>
                <option value="ocupado por terceiros">Ocupado por terceiros</option>
                <option value="ocupado pelo proprietário">Ocupado pelo proprietário</option>
              </select></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
      <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="192"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></td>
            <td width="10">&nbsp;</td>
            <td width="192"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></td>
            <td width="10">&nbsp;</td>
            <td width="192"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></td>
            <td width="10">&nbsp;</td>
            <td width="394"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></td>
          </tr>
        </table></td>
  </tr>
  <tr> 
      <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" onChange="javascript:atualizacarros2(this.form);">
                      <option value="<% if rs("cidade") = "não informado" or rs555.eof then response.write "cqualquer" else response.write rs555("id_combo1") end if  %>" select><%if   rs("cidade") = "cqualquer" then response.write "não informado" else response.write rs("cidade") end if  %></option>
					
					 <% if not rs33.eof then %>
				    <% While NOT Rs33.EoF %>
                    <option value="<% = Rs33("id_combo1") %>">
                    <% = Rs33("nome_combo1") %>
                    </option>
                    <% Rs33.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
                  </select></td>
            <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo4" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" onChange="javascript:atualizacarros999(this.form);">
                      
                        <option value="<%if rs("bairro") = "não informado" or rs4Bairro.eof then response.write "bqualquer" else response.write rs4Bairro("id_combo2") end if%>" ><%if   rs("bairro") = "bqualquer" then response.write "não informado" else response.write rs("bairro") end if  %></option>
                       
                        <option value=""></option>
                  </select></td>
            <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                     <option value="<%if rs("vila") <> "não informado" and  rs("vila") <>"" and  not rs444Vila.eof then response.write rs444Vila("id_combo3") else response.write "vlqualquer" end if%>" selected><%if rs("vila") <> "não informado" and  rs("vila") <>"" then response.write rs("vila") else response.write "não informado" end if%></option>
				  <option value="vlqualquer">qualquer um</option>
                    </select></td>
            <td width="10">&nbsp;</td>
            <td width="394" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco_vend" type="text" class="inputBox" id="txt_endereco_vend" style="HEIGHT: 18px; WIDTH: 390px; background:<%=medio%>" value="<%=rs("endereco")%>" size="38" maxlength="50" align="left"></td>
          </tr>
        </table></td>
  </tr>
  <tr> 
    <td> <table width="1000" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
            do im&oacute;vel</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
            do condom&iacute;nio</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
            do IPTU</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Outros</font></td>
          <td width="10">&nbsp;</td>
            <td width="192"  bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Rateio</font></td>
        </tr>
        <tr> 
            <td width="192" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=FormatNumber(rs("valor"),2)%>" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_condominio_vend" type="text" class="inputBox" id="txt_condominio_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%if rs("condominio") <> "" then response.write FormatNumber(rs("condominio"),2) else response.write "0,00" end if%>" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_iptu_vend" type="text" class="inputBox" id="txt_valor_iptu_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("valor_iptu")%>" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_outros_vend" type="text" class="inputBox" id="txt_valor_outros_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("valor_outros")%>" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_rateio_vend" type="text" class="inputBox" id="txt_rateio_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("rateio")%>" size="12" maxlength="13"></td>
        </tr>
        <tr> 
            <td width="192" height="21" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
              devedor</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
              devedor j&aacute; pago</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
              devedor a pagar</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
              total apartamento ou terreno</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
              constru&iacute;da p/casas ou &uacute;til apart</font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_saldo_devedor_vend" id="txt_saldo_devedor_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="<%if rs("saldo_devedor") <> "" then response.write rs("saldo_devedor") else response.write "sem saldo devedor" end if%>" selected >
                <%if rs("saldo_devedor") <> "" then response.write rs("saldo_devedor") else response.write "sem saldo devedor" end if%>
                </option>
                <option value="sem saldo devedor">Sem saldo devedor</option>
                <option value="com saldo devedor" >Com saldo devedor</option>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_ja_pago_devedor_vend" type="text" class="inputBox" id="txt_ja_pago_devedor_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%if rs("ja_pago_devedor") <> "" then response.write formatNumber(rs("ja_pago_devedor"),2) else response.write "0,00" end if %>" size="12" maxlength="30"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
              <input name="txt_devendo_devedor_vend" type="text" class="inputBox" id="txt_devendo_devedor_vend2" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%if rs("devendo_devedor") <> "" then response.write formatNumber(rs("devendo_devedor"),2) else response.write "0,00" end if%>" size="12" maxlength="30">
              </td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
              <input name="txt_a_total_vend" type="text" class="inputBox" id="txt_a_total_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("area_total")%>" size="12" maxlength="20">
              </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_a_constr_vend" type="text" class="inputBox" id="txt_a_constr_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%=rs("area_construida")%>" size="12" maxlength="20"></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
              de frente</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
              de fundo</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
              lateral esquerda</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
              lateral direita</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              Nome do edif&iacute;cio</font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_de_frente_vend" type="text" class="inputBox" id="txt_metros_de_frente_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs("metros_de_frente") <> "" then response.write rs("metros_de_frente") else response.write "00" end if%>" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_de_fundo_vend" type="text" class="inputBox" id="txt_metros_de_fundo_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs("metros_de_fundo") <> "" then response.write rs("metros_de_fundo") else response.write "00" end if%>" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_lateral_esquerda_vend" type="text" class="inputBox" id="txt_metros_lateral_esquerda_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs("metros_lateral_esquerda") <> "" then response.write rs("metros_lateral_esquerda") else response.write "00" end if%>" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_lateral_direita_vend" type="text" class="inputBox" id="txt_metros_lateral_direita_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs("metros_lateral_direita") <> "" then response.write rs("metros_lateral_direita") else response.write "00" end if%>" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_nome_edificio_vend" type="text" class="inputBox" id="txt_nome_edificio_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%=rs("nome_edificio")%>" size="12" maxlength="20"></td>
        </tr>
      </table></td>
  </tr>
  
  <tr>
  <td>
  <table cellpadding="0" cellspacing="0" border="0">
  
  
  <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Primeira 
              p&aacute;gina </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
              do an&uacute;ncio</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto 
              do an&uacute;ncio</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quem 
              conseguiu uma proposta</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Situa&ccedil;&atilde;o 
              do im&oacute;vel</font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_presenca_primeira_vend" id="txt_presenca_primeira_vend" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                <option value="<%if rs("presenca_primeira") <> "" then response.write rs("presenca_primeira") else response.write "excluido" end if %>"selected>
                <%if rs("presenca_primeira") <> "" then response.write rs("presenca_primeira") else response.write "excluido" end if %>
                </option>
                <% if session("permissao") <> "6" and session("permissao") <> "5" then %>
                <option value="excluido">Excluído</option>
                <%else%>
                <option value="incluido">Incluído</option>
                <option value="excluido">Excluído</option>
                <%end if%>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_titulo_anuncio_vend" type="text" class="inputBox" id="txt_titulo_anuncio_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs("titulo_anuncio")%>" size="38" maxlength="40" align="left"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_texto_anuncio_vend" type="text" class="inputBox" id="txt_texto_anuncio_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%=rs("texto_anuncio")%>" size="38" maxlength="120" align="left"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>" style="border:1px solid #FFFFFF;"><select name="txt_conseguiu_proposta_vend" id="txt_conseguiu_proposta_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="<%=rs("conseguiu_proposta")%>" ><%=rs("conseguiu_proposta")%></option>
                <% if not rs444ConseguiuProposta22.eof then %>
                <% While NOT rs444ConseguiuProposta22.EoF %>
                <option value="<% = rs444ConseguiuProposta22("list_name") %>"> 
                <% = rs444ConseguiuProposta22("list_name") %>
                </option>
                <% rs444ConseguiuProposta22.MoveNext %>
                <% Wend %>
                <%else%>
                <option value="Internet">Internet</option>
                <%end if%>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>" style="border:1px solid #FFFFFF;"><select name="txt_imovel_em_negociacao_vend" id="txt_imovel_em_negociacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="<%if rs("imovel_em_negociacao") <> "" then response.write rs("imovel_em_negociacao") else response.write "não informado" end if%>" selected>
                <%if rs("imovel_em_negociacao") <> "" then response.write rs("imovel_em_negociacao") else response.write "não informado" end if%>
                </option>
                <option value="não informado">não informado</option>
                <option value="Vendido pela Veja">Vendido pela Veja</option>
                <option value="Vendido por outros">Vendido por outros</option>
                <option value="Suspenso">Suspenso</option>
                <option value="Com proposta">Com proposta</option>
                <option value="Imóvel inexistente">Imóvel inexistente</option>
                <option value="Imóvel para permuta">Imóvel para permuta</option>
               
                <option value="imóvel OK" >imóvel OK</option>
			 
			  
			  </select></td>
        </tr>
  
  
  </table>
  
  </td>
  
  </tr>
  
  <tr>
  <td>
  <table cellpadding="0" cellspacing="0" border="0">
  
  
  <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quem 
              tirou a foto</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
              de propostas recebidas</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cliques 
              no im&oacute;vel</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<select name="txt_quem_tirou_foto_vend" id="txt_quem_tirou_foto_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <% if rs("quem_tirou_foto") = "" or rs("quem_tirou_foto") = "não informado"  or  session("permissao") = "6" then %>
                <option value="<%if rs("quem_tirou_foto") <> "" then response.write rs("quem_tirou_foto") else response.write "não informado" end if%>" selected>
                <%if rs("quem_tirou_foto") <> "" then response.write rs("quem_tirou_foto") else response.write "não informado" end if%>
                </option>
                <% if not rs445Responsavel22.eof then %>
                <% While NOT rs445Responsavel22.EoF %>
                <option value="<% = rs445Responsavel22("list_name") %>"> 
                <% = rs445Responsavel22("list_name") %>
                </option>
                <% rs445Responsavel22.MoveNext %>
                <% Wend %>
                <%else%>
                <option value="Internet">Internet</option>
                <%end if%>
                <%else%>
                <option value="<%if rs("quem_tirou_foto") <> "" then response.write rs("quem_tirou_foto") else response.write "não informado" end if%>" selected>
                <%if rs("quem_tirou_foto") <> "" then response.write rs("quem_tirou_foto") else response.write "não informado" end if%>
                </option>
                <%end if%>
              </select>
            </td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
              <%
			dim SQLProposta
			dim RSProposta
			
			SQLProposta ="SELECT proposta.telefone_proposta,proposta.Nome_proposta,proposta.proposta_proposta,proposta.Cod_proposta,proposta.interesse_proposta,proposta.cod_imovel_proposta  FROM proposta where  cod_imovel_proposta ='"&varCod_imovel&"'" 
	
			Set RSProposta = Server.CreateObject("ADODB.RecordSet")

	RSProposta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RSProposta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RSProposta.ActiveConnection = Conexao3
	
	
	RSProposta.Open SQLProposta, Conexao3
			
			
			if not RSProposta.eof then
			
			
			%>
			<font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=RSProposta.recordCount%> </font> 
              <%else%>
              <font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nenhuma 
              proposta </font> 
              <%end if%>
              <%
			RSProposta.Close           
		   
           Set RSProposta = Nothing
			
			%>
            </td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input type="text" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%=rs("cliques_no_imovel")%>" size="38" maxlength="120" align="left"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>" >&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>" >&nbsp;</td>
        </tr>
  
  
  </table>
  
  </td>
  
  </tr>
  
  
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
            <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
            na garagem</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quartos_vend" id="txt_quartos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					  <option value="<%=rs("quartos")%>" selected><% if rs("quartos") = "0" then response.write "não informado" else response.write rs("quartos") end if%></option>
                    
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quartos_vend" type="text" class="inputBox" id="txt_obs_quartos_vend" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_quartos")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_vagas_vend" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					  <option value="<%=rs("vagas")%>" selected><% if rs("vagas") = "0" then response.write "não informado" else response.write rs("vagas") end if%></option>
                     
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_vagas_vend" type="text" class="inputBox" id="txt_obs_vagas_vend" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_vagas")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_banheiros_vend" id="txt_banheiros_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					  <option value="<%=rs("banheiros")%>" selected><%=rs("banheiros")%></option>
                     
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_banheiros_vend" type="text" class="inputBox" id="txt_obs_banheiros_vend" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_banheiros")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
          <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Su&iacute;tes</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ed&iacute;cula</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Entrada 
                  lateral </font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_suites_vend" id="txt_suites_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      <option value="<%if rs("suites") <> "" then response.write rs("suites") else response.write "0" end if%>" selected><%if rs("suites") <> "" then response.write rs("suites") else response.write "0" end if%></option>
                  
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_suites_vend" type="text" class="inputBox" id="txt_obs_suites_vend" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_suites")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_edicula_vend" id="txt_edicula_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					 
					  
					  <option value="<%if rs("edicula") <> "" then response.write rs("edicula") else response.write "não informado" end if%>" selected><%if rs("edicula") <> "" then response.write rs("edicula") else response.write "não informado" end if%></option>
                  
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_edicula_vend" type="text" class="inputBox" id="txt_obs_vagas4" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_edicula")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_entrada_lateral_vend" id="txt_entrada_lateral_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      
					  
					   <option value="<%if rs("entrada_lateral") <> "" then response.write rs("entrada_lateral") else response.write "não informado" end if%>" selected><%if rs("entrada_lateral") <> "" then response.write rs("entrada_lateral") else response.write "não informado" end if%></option>
                  
					  
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_entrada_lateral_vend" type="text" class="inputBox" id="txt_obs_vagas5" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_entrada_lateral")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sal&atilde;o 
                  de festas</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sal&atilde;o 
                  de Jogos</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Churrasqueira</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_salao_de_festas_vend" id="txt_salao_de_festas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      <option value="<%if rs("salao_de_festas") <> "" then response.write rs("salao_de_festas") else response.write "não informado" end if%>" selected><%if rs("salao_de_festas") <> "" then response.write rs("salao_de_festas") else response.write "não informado" end if%></option>
                  
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_festas_vend" type="text" class="inputBox" id="txt_obs_suites" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_salao_de_festas")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_salao_de_jogos_vend" id="txt_salao_de_jogos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                    
					 
					  <option value="<%if rs("salao_de_jogos") <> "" then response.write rs("salao_de_jogos") else response.write "não informado" end if%>" selected><%if rs("salao_de_jogos") <> "" then response.write rs("salao_de_jogos") else response.write "não informado" end if%></option>
                  
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_jogos_vend" type="text" class="inputBox" id="txt_obs_suites2" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_salao_de_jogos")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_churrasqueira_vend" id="txt_churrasqueira_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      
					  
					   <option value="<%if rs("churrasqueira") <> "" then response.write rs("churrasqueira") else response.write "não informado" end if%>" selected><%if rs("churrasqueira") <> "" then response.write rs("churrasqueira") else response.write "não informado" end if%></option>
                  
					  
					  
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_churrasqueira_vend" type="text" class="inputBox" id="txt_obs_suites3" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_churrasqueira")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="1019" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
          <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Piscina</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quintal</font></td>
          <td width="10">&nbsp;</td>
          <td width="346"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quadras</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_piscina_vend" id="txt_piscina_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					  <option value="<%if rs("piscina") <> "" then response.write rs("piscina") else response.write "não informado" end if%>" selected><%if rs("piscina") <> "" then response.write rs("piscina") else response.write "não informado" end if%></option>
                  
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_piscina_vend" type="text" class="inputBox" id="txt_obs_suites4" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_piscina")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quintal_vend" id="txt_quintal_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      
					  
					  
					  <option value="<%if rs("quintal") <> "" then response.write rs("quintal") else response.write "não informado" end if%>" selected><%if rs("quintal") <> "" then response.write rs("quintal") else response.write "não informado" end if%></option>
                  
					  
					  
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quintal_vend" type="text" class="inputBox" id="txt_obs_suites5" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_quintal")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quadras_vend" id="txt_quadras_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					 
					 
					  
					   <option value="<%if rs("quadras") <> "" then response.write rs("quadras") else response.write "não informado" end if%>" selected><%if rs("quadras") <> "" then response.write rs("quadras") else response.write "não informado" end if%></option>
                  
					 
					 
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quadras_vend" type="text" class="inputBox" id="txt_obs_suites6" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_quadras")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
          <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Andares 
            do edif&iacute;cio</font></td>
          <td width="10">&nbsp;</td>
                
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quantidade 
            de elevadores</font></td>
          <td width="10">&nbsp;</td>
                
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Portaria</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_andares_edificio_vend" id="txt_andares_edificio_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                    
					  <option value="<%if rs("andares_edificio") <> "" then response.write rs("andares_edificio") else response.write "não informado" end if%>" selected><%if rs("andares_edificio") <> "" then response.write rs("andares_edificio") else response.write "não informado" end if%></option>
                  
					
					
					 <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    <option value="11">11</option>
                    <option value="12">12</option>
					<option value="13">13</option>
					<option value="14">14</option>
                    <option value="15">15</option>
					<option value="16">16</option>
					<option value="17">17</option>
                    <option value="18">18</option>
					<option value="19">19</option>
					<option value="20">20</option>
					<option value="21">21</option>
                    <option value="22">22</option>
					<option value="23">23</option>
					<option value="24">24</option>
                    <option value="25">25</option>
					<option value="26">26</option>
					<option value="27">27</option>
                    <option value="28">28</option>
					<option value="29">29</option>
					<option value="30">30</option>
					<option value="31">31</option>
                    <option value="32">32</option>
					<option value="33">33</option>
					<option value="34">34</option>
                    <option value="35">35</option>
					<option value="36">36</option>
					<option value="37">37</option>
                    <option value="38">38</option>
					<option value="39">39</option>
					<option value="40">40</option>
					<option value="41">41</option>
                    <option value="42">42</option>
					<option value="43">43</option>
					<option value="44">44</option>
                    <option value="45">45</option>
					<option value="46">46</option>
					<option value="47">47</option>
                    <option value="48">48</option>
					<option value="49">49</option>
					<option value="50">50</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_andares_do_edificio_vend" type="text" class="inputBox" id="txt_obs_piscina" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_andares_edificio")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quantidade_elevadores_vend" id="txt_quantidade_elevadores_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					 
					   <option value="<%if rs("quantidade_elevadores") <> "" then response.write rs("quantidade_elevadores") else response.write "não informado" end if%>" selected><%if rs("quantidade_elevadores") <> "" then response.write rs("quantidade_elevadores") else response.write "não informado" end if%></option>
                  
					 
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quantidade_elevadores_vend" type="text" class="inputBox" id="txt_obs_piscina2" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_quantidade_elevadores")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_portaria_vend" id="txt_portaria_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					 
					  
					   <option value="<%if rs("portaria") <> "" then response.write rs("portaria") else response.write "não informado" end if%>" selected><%if rs("portaria") <> "" then response.write rs("portaria") else response.write "não informado" end if%></option>
                  
					 
					 
					 
					 
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_portaria_vend" type="text" class="inputBox" id="txt_obs_piscina3" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>" value="<%=rs("obs_portaria")%>" size="38" maxlength="40" align="left"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr>
    <td></td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
              sobre forma de pagamento</font></td>
          <td width="10">&nbsp;</td>
            <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
              sobre o im&oacute;vel e forma de pagamento que o propriet&aacute;rio 
              aceita</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td width="495" height="224" bgcolor="<%=medio%>"><table width="495" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="102" style="border:1px solid #FFFFFF;"><textarea name="txt_obs_proprietario_vend" class="inputBox" id="textarea4" style="HEIGHT: 100px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%=rs("obs_proprietario")%></textarea></td>
                </tr>
                <tr>
                  <td height="20" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
                    confidenciais sobre o propriet&aacute;rio</font></td>
                </tr>
                <tr>
                  <td height="102" style="border:1px solid #FFFFFF;"><textarea class="inputBox" id="textarea6" style="HEIGHT: 100px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
                </tr>
              </table></td>
          <td width="10">&nbsp;</td>
          <td width="495" height="80" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_obs_imovel_vend" class="inputBox" id="txt_obs_imovel_vend" style="HEIGHT: 222px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%=rs("obs_imovel")%></textarea></td>
        </tr>
      </table></td>
  </tr>
  
  
  
  
  
  
  <tr> 
    <td height="300"><div align="center">
	<%
	dim rsFichaCidade
	
	dim SqlFichaCidade
	
	SqlFichaCidade ="SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 where nome_combo1='"&CidadeCompradorFicha01&"' ORDER BY nome_combo1" 

	
	Set rsFichaCidade = Server.CreateObject("ADODB.RecordSet")

	rsFichaCidade.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsFichaCidade.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsFichaCidade.ActiveConnection = Conexao3
	
	
	rsFichaCidade.Open SqlFichaCidade, Conexao3

    dim cidadeFichaComprador
	dim codCidadeFichaComprador
	
	if not rsFichaCidade.eof then
	cidadeFichaComprador = rsFichaCidade("nome_combo1")
	codCidadeFichaComprador = rsFichaCidade("id_combo1")
	else
	cidadeFichaComprador = "não informado"
	codCidadeFichaComprador = "cqualquer"
	end if
	
	'------------------------pegar o bairro em comprador para por na ficha de baixo--------
	
	
	dim rsFichaBairro
	
	dim SqlFichaBairro
	
	SqlFichaBairro ="Select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE nome_combo2='"&BairroCompradorFicha01&"' ORDER BY nome_combo2" 

	
	Set rsFichaBairro = Server.CreateObject("ADODB.RecordSet")

	rsFichaBairro.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsFichaBairro.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsFichaBairro.ActiveConnection = Conexao3
	
	
	rsFichaBairro.Open SqlFichaBairro, Conexao3

    dim bairroFichaComprador
	dim codBairroFichaComprador
	
	if not rsFichaBairro.eof then
	bairroFichaComprador = rsFichaBairro("nome_combo2")
	codBairroFichaComprador = rsFichaBairro("id_combo2")
	else
	bairroFichaComprador = "não informado"
	codBairroFichaComprador = "bqualquer"
	end if
	
	
	
	
	
	
	
	
	%>
	
	<table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
              <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Caso 
                o propriet&aacute;rio desse im&oacute;vel seja tamb&eacute;m um 
                comprador de im&oacute;veis, selecione sim e preencha a ficha 
                abaixo. </strong></font></td>
  </tr>
</table>

	
	<br><br> 
	
          <select name="txt_pergunta" id="txt_pergunta" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
           <option value="não"  >Não</option>
			<option value="sim" <% if  ( semComprador <> "nothing" and semComprador <> "" ) then response.write "selected" end if %> >Sim</option>
            
          </select>
        </div></td>
  </tr>
  
  <%
   dim PerguntaDe
  PerguntaDe = rs("pergunta")
  
  if PerguntaDe = "" then
  PerguntaDe = "sim"
  end if
  
  
 
  
  
  if ( semComprador <> "nothing" and semComprador <> "" )  then
  
  
  
  %>
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="192"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
            futuro contato</font></td>
          <td width="10">&nbsp;</td>
          <td width="762"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto 
            futuro contato</font></td>
          <td width="10">&nbsp;</td>
          <td width="26"></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_futuro_contato_comprador" type="text" class="inputBox" id="txt_data_futuro_contato_comprador" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=rs444Comprador("data_futuro_contato")%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="798" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_assunto_futuro_contato_comprador" type="text" class="inputBox" id="txt_assunto_futuro_contato_comprador" style="HEIGHT: 18px; WIDTH: 796px; background:<%=medio%>" value="<%=rs444Comprador("assunto_futuro_contato")%>" size="38" maxlength="50" align="left"></td>
          
        </tr>
      </table></td>
  </tr>
  
  
  <tr> 
    <td>
	<table width="1000" cellpadding="0" cellspacing="0">
	
        <tr> 
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Origem 
            do comprador</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
            pelo cadastramento</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
            pelo atendimento</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Melhor 
            hor&aacute;rio para visita</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_origem" id="txt_origem" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=claro%>">
              
			  
					 <option value="<% if rs444Comprador("origem") <> "não informado" and rs444Comprador("origem") <> "" then response.write rs444Comprador("origem") else response.write "não informado" end if%>" selected><% if rs444Comprador("origem") <> "não informado" and rs444Comprador("origem") <> "" then response.write rs444Comprador("origem") else response.write "não informado" end if%></option>
			    
			  
			    <% if not rs444Origem.eof then %>
                <% While NOT rs444Origem.EoF %>
                <option value="<% = rs444Origem("origem") %>"> 
                <% = rs444Origem("origem") %>
                </option>
                <% rs444Origem.MoveNext %>
                <% Wend %>
                <%else%>
                <option value="não informado">não informado</option>
                <%end if%>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_responsavel_cadastramento_comprador" id="txt_responsavel_cadastramento_comprador" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
              
			   <option value="<%=rs("captacao")%>"><%=rs("captacao")%></option>
					
			   
					 <option value="<% if rs444Comprador("responsavel_cadastramento") <> "não informado" and rs444Comprador("melhor_horario_visita") <> "" then response.write rs444Comprador("responsavel_cadastramento") else response.write rs("captacao") end if%>" ><% if rs444Comprador("responsavel_cadastramento") <> "não informado" and rs444Comprador("responsavel_cadastramento") <> "" then response.write rs444Comprador("responsavel_cadastramento") else response.write rs("captacao") end if%></option>
			    
			    <option value="Internet"  >Internet</option>
                <% if not rs444Responsavel.eof then %>
                <% While NOT rs444Responsavel.EoF %>
                <option value="<% = rs444Responsavel("list_name") %>"> 
                <% = rs444Responsavel("list_name") %>
                </option>
                <% rs444Responsavel.MoveNext %>
                <% Wend %>
                <%else%>
                <option value="Internet">Internet</option>
                <%end if%>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_atendimento" id="txt_atendimento" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                      
					    <option value="<%=rs("captacao")%>"><%=rs("captacao")%></option>
					
					 <option value="<% if rs444Comprador("atendimento") <> "não informado" and rs444Comprador("atendimento") <> "" then response.write rs444Comprador("atendimento") else response.write "não informado" end if%>" ><% if rs444Comprador("atendimento") <> "não informado" and rs444Comprador("atendimento") <> "" then response.write rs444Comprador("atendimento") else response.write "não informado" end if%></option>
			    
					  <option value="Internet" >Internet</option>
                      <% if not rs444Captacao22.eof then %>
                      <% While NOT rs444Captacao22.EoF %>
                      <option value="<% = rs444Captacao22("list_name") %>"> 
                      <% = rs444Captacao22("list_name") %>
                      </option>
                      <% rs444Captacao22.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Internet">Internet</option>
                      <%end if%>
                    </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_melhor_horario_visita_comprador" size="1" class="inputBox" id="txt_melhor_horario_visita_comprador"  style="HEIGHT: 18px; WIDTH: 190px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
                     
					 
					 <option value="<% if rs444Comprador("melhor_horario_visita") <> "não informado" and rs444Comprador("melhor_horario_visita") <> "" then response.write rs444Comprador("melhor_horario_visita") else response.write "não informado" end if%>" selected><% if rs444Comprador("melhor_horario_visita") <> "não informado" and rs444Comprador("melhor_horario_visita") <> "" then response.write rs444Comprador("melhor_horario_visita") else response.write "não informado" end if%></option>
			    
					  <option  value="Manhâ">Manhã </option>
                      <option value="Tarde" >Tarde </option>
					   <option  value="Noite">Noite </option>
                      <option value="Manhã ou tarde" >Manhã ou tarde </option>
                     <option  value="Manhã ou noite">Manhã ou noite</option>
                      <option value="Tarde ou noite" >Tarde ou noite </option>
					  <option value="Qualquer horário">Qualquer horário</option>
					
					</select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
               
			   <option value="<% if rs444Comprador("ocupacao") <> "não informado" and rs444Comprador("ocupacao") <> "" then response.write rs444Comprador("ocupacao") else response.write "não informado" end if%>" selected><% if rs444Comprador("ocupacao") <> "não informado" and rs444Comprador("ocupacao") <> "" then response.write rs444Comprador("ocupacao") else response.write "não informado" end if%></option>
			   
			    <option value="não informado" selected>não informado</option>
                <option value="vago">Vago</option>
                <option value="ocupado por terceiros">Ocupado por terceiros</option>
                <option value="ocupado pelo inquilino">Ocupado pelo inquilino</option>
                <option value="ocupado pelo proprietário">Ocupado por terceiros</option>
              </select></td>
        </tr>
        
        <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" valign="top" style="border:1px solid #FFFFFF;" ><select name="combo1" id="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                <option value="<%=codCidadeFichaComprador%>" selected><%=cidadeFichaComprador%></option>
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
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" height="120" valign="top" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><select name="combo2" id="combo2" class="inputBox" style="HEIGHT: 130px; WIDTH: 190px; background:<%=medio%>" multiple size="8">
               
			     <%				 
	  '-----------------------pegar vários bairros-----------
  
  
  
dim VariavelFicha
dim RetornoFicha
dim iFicha
VariavelFicha = BairroCompradorFicha01
RetornoFicha = Split(VariavelFicha,", ")

iFicha=0

Set rs04Ficha = Server.CreateObject("ADODB.RecordSet")


for iFicha=0 to UBound(RetornoFicha)



strSQL04Ficha = "select * from combo2 where nome_combo2 like '"& RetornoFicha(iFicha) &"' and cidade_combo2 ='"&cidadeFichaComprador&"' "

 
 

rs04Ficha.open strSQL04Ficha,Conexao,2,1




while not (rs04Ficha.eof)

%>
                <option value="<%=rs04Ficha("id_combo2")%>" selected><%=rs04Ficha("nome_combo2")%></option>
                <%
rs04Ficha.MoveNext
Wend



rs04Ficha.close




%>
                <%
next



%>
			   
			   
			   
			   
			   
			    <% if not rs4.eof then%>
                <% While NOT Rs4.EoF %>
                <option value="<% = Rs4("id_combo2") %>"> 
                <% = Rs4("nome_combo2") %>
                </option>
                <% Rs4.MoveNext %>
                <% Wend %>
                <% else %>
                <option value=""></option>
                <% end if %>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" valign="top" style="border:1px solid #FFFFFF;" ><select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                <option value="vlqualquer" selected>Vila</option>
                <option value="vlqualquer">qualquer um</option>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" valign="top" style="border:1px solid #FFFFFF;"><select name="txt_tipo" multiple size="8"  id="txt_tipo" class="inputBox" style="HEIGHT: 130px; WIDTH: 190px; background: <%=medio%>">
                   
					
					 <%				 
	  '-----------------------pegar vários bairros-----------
  
  
  
dim VariavelTipoFicha
dim RetornoTipoFicha
dim iTipoFicha
VariavelTipoFicha = rs444Comprador("tipo")
RetornoTipoFicha = Split(VariavelTipoFicha,", ")

iTipoFicha=0

Set rs04TipoFicha = Server.CreateObject("ADODB.RecordSet")


for iTipoFicha=0 to UBound(RetornoTipoFicha)



strSQL04TipoFicha = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo where tipo like '"& RetornoTipoFicha(iTipoFicha) &"'  ORDER BY tipo ASC"

 
 

rs04TipoFicha.open strSQL04TipoFicha,Conexao,2,1

while not (rs04TipoFicha.eof)

%>
                <option value="<%=rs04TipoFicha("tipo")%>" selected><%=rs04TipoFicha("tipo")%></option>
                <%
rs04TipoFicha.MoveNext
Wend

rs04TipoFicha.close




%>
                <%
next



%>
					
					
					
					
					
					
					
					
					
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
                  </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" valign="top" style="border:1px solid #FFFFFF;"><select name="example2" size="1" class="inputBox" id="select3"  style="HEIGHT: 18px; WIDTH: 190px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
                
				<option value="<% if rs444Comprador("negociacao") <> "nqualquer" and rs444Comprador("negociacao") <> "" then response.write rs444Comprador("negociacao") else response.write "não informado" end if%>" selected><% if rs444Comprador("negociacao") <> "nqualquer" and rs444Comprador("negociacao") <> "" then response.write rs444Comprador("negociacao") else response.write "não informado" end if%></option>
				<option  value="aluguel">Aluguel </option>
                <option value="compra" >Compra </option>
              </select> </table>
	
	
	</td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
            <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
            na garagem</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quartos" id="txt_quartos" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                    
					<option value="<% if rs444Comprador("quartos") <> "qqualquer" and rs444Comprador("quartos") <> "" then response.write rs444Comprador("quartos") else response.write "0" end if%>" selected><% if rs444Comprador("quartos") <> "qqualquer" and rs444Comprador("quartos") <> "" then response.write rs444Comprador("quartos") else response.write "0" end if%></option>
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quartos" value="<%=rs444Comprador("obs_quartos")%>" type="text" id="txt_obs_quartos" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     <option value="<% if rs444Comprador("vagas") <> "vgqualquer" and rs444Comprador("vagas") <> "" then response.write rs444Comprador("vagas") else response.write "0" end if%>" selected><% if rs444Comprador("vagas") <> "vgualquer" and rs444Comprador("vagas") <> "" then response.write rs444Comprador("vagas") else response.write "0" end if%></option>
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_vagas" value="<%=rs444Comprador("obs_vagas")%>" type="text" id="txt_obs_vagas" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_banheiros"   id="txt_banheiros" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      
					  <option value="<% if rs444Comprador("banheiros") <> "não informado" and rs444Comprador("banheiros") <> "" then response.write rs444Comprador("banheiros") else response.write "0" end if%>" selected><% if rs444Comprador("banheiros") <> "não informado" and rs444Comprador("banheiros") <> "" then response.write rs444Comprador("banheiros") else response.write "0" end if%></option>
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_banheiros" value="<%=rs444Comprador("obs_banheiros")%>" type="text" id="txt_obs_banheiros" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
          <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Su&iacute;tes</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ed&iacute;cula</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Entrada 
                  lateral </font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_suites" id="txt_suites" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					 <option value="<% if rs444Comprador("suites") <> "não informado" and rs444Comprador("suites") <> "" then response.write rs444Comprador("suites") else response.write "0" end if%>" selected><% if rs444Comprador("suites") <> "qqualquer" and rs444Comprador("suites") <> "" then response.write rs444Comprador("suites") else response.write "0" end if%></option>
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_suites" value="<%=rs444Comprador("obs_suites")%>" type="text" id="txt_obs_suites" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_edicula" id="txt_edicula" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      
					    <option value="<% if rs444Comprador("edicula") <> "não informado" and rs444Comprador("edicula") <> "" then response.write rs444Comprador("edicula") else response.write "0" end if%>" selected><% if rs444Comprador("edicula") <> "não informado" and rs444Comprador("edicula") <> "" then response.write rs444Comprador("edicula") else response.write "0" end if%></option>
					
					  
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_edicula" value="<%=rs444Comprador("obs_edicula")%>" type="text" id="txt_obs_edicula" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_entrada_lateral" id="txt_entrada_lateral" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                    
					
					 
				     <option value="<% if rs444Comprador("entrada_lateral") <> "não informado" and rs444Comprador("entrada_lateral") <> "" then response.write rs444Comprador("entrada_lateral") else response.write "0" end if%>" selected><% if rs444Comprador("entrada_lateral") <> "não informado" and rs444Comprador("entrada_lateral") <> "" then response.write rs444Comprador("entrada_lateral") else response.write "0" end if%></option>
					
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_entrada_lateral" value="<%=rs444Comprador("obs_entrada_lateral")%>" type="text" id="txt_obs_entrada_lateral" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sal&atilde;o 
                  de festas</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sal&atilde;o 
                  de Jogos</font></td>
          <td width="10">&nbsp;</td>
                <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Churrasqueira</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                  <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_salao_de_festas" id="txt_salao_de_festas" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      
					   <option value="<% if rs444Comprador("salao_de_festas") <> "não informado" and rs444Comprador("salao_de_festas") <> "" then response.write rs444Comprador("salao_de_festas") else response.write "0" end if%>" selected><% if rs444Comprador("salao_de_festas") <> "não informado" and rs444Comprador("salao_de_festas") <> "" then response.write rs444Comprador("salao_de_festas") else response.write "0" end if%></option>
					
					  <option value="00">00</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                      <option value="10">10</option>
                    </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_festas" value="<%=rs444Comprador("obs_salao_de_festas")%>" type="text" id="txt_obs_salao_de_festas" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_salao_de_jogos" id="txt_salao_de_jogos" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					  <option value="<% if rs444Comprador("salao_de_jogos") <> "não informado" and rs444Comprador("salao_de_jogos") <> "" then response.write rs444Comprador("salao_de_jogos") else response.write "0" end if%>" selected><% if rs444Comprador("salao_de_jogos") <> "não informado" and rs444Comprador("salao_de_jogos") <> "" then response.write rs444Comprador("salao_de_jogos") else response.write "0" end if%></option>
					
					 
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_jogos" value="<%=rs444Comprador("obs_salao_de_jogos")%>" type="text" id="txt_obs_salao_de_jogos" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_churrasqueira" id="txt_churrasqueira" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                    
					  <option value="<% if rs444Comprador("churrasqueira") <> "não informado" and rs444Comprador("churrasqueira") <> "" then response.write rs444Comprador("churrasqueira") else response.write "0" end if%>" selected><% if rs444Comprador("churrasqueira") <> "não informado" and rs444Comprador("churrasqueira") <> "" then response.write rs444Comprador("churrasqueira") else response.write "0" end if%></option>
					
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_churrasqueira" value="<%=rs444Comprador("obs_churrasqueira")%>" type="text" id="txt_obs_churrasqueira" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
          <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Piscina</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quintal</font></td>
          <td width="10">&nbsp;</td>
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quadras</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_piscina" id="txt_piscina" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                    
					 <option value="<% if rs444Comprador("piscina") <> "não informado" and rs444Comprador("piscina") <> "" then response.write rs444Comprador("piscina") else response.write "0" end if%>" selected><% if rs444Comprador("piscina") <> "não informado" and rs444Comprador("piscina") <> "" then response.write rs444Comprador("piscina") else response.write "0" end if%></option>
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_piscina" value="<%=rs444Comprador("obs_piscina")%>" type="text" id="txt_obs_piscina" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quintal" id="txt_quintal" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                   
				   
				     <option value="<% if rs444Comprador("quintal") <> "não informado" and rs444Comprador("quintal") <> "" then response.write rs444Comprador("quintal") else response.write "0" end if%>" selected><% if rs444Comprador("quintal") <> "não informado" and rs444Comprador("quintal") <> "" then response.write rs444Comprador("quintal") else response.write "0" end if%></option>
					
				   
				      <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quintal" value="<%=rs444Comprador("obs_quintal")%>" type="text" id="txt_obs_quintal" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quadras" id="txt_quadras" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                  
				    <option value="<% if rs444Comprador("quadras") <> "não informado" and rs444Comprador("quadras") <> "" then response.write rs444Comprador("quadras") else response.write "0" end if%>" selected><% if rs444Comprador("quadras") <> "não informado" and rs444Comprador("quadras") <> "" then response.write rs444Comprador("quadras") else response.write "0" end if%></option>
					
				  
				  
				      <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quadras" value="<%=rs444Comprador("obs_quadras")%>" type="text" id="txt_obs_quadras" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
                
          <td width="326"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Andares 
            do edif&iacute;cio</font></td>
          <td width="10">&nbsp;</td>
                
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quantidade 
            de elevadores</font></td>
          <td width="10">&nbsp;</td>
                
          <td width="327"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Portaria</font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="326"><table width="326" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_andares_edificio" id="txt_andares_edificio" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                     
					  <option value="<% if rs444Comprador("andares_edificio") <> "não informado" and rs444Comprador("andares_edificio") <> "" then response.write rs444Comprador("andares_edificio") else response.write "0" end if%>" selected><% if rs444Comprador("andares_edificio") <> "não informado" and rs444Comprador("andares_edificio") <> "" then response.write rs444Comprador("andares_edificio") else response.write "0" end if%></option>
					
					  <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    <option value="11">11</option>
                    <option value="12">12</option>
					<option value="13">13</option>
					<option value="14">14</option>
                    <option value="15">15</option>
					<option value="16">16</option>
					<option value="17">17</option>
                    <option value="18">18</option>
					<option value="19">19</option>
					<option value="20">20</option>
					<option value="21">21</option>
                    <option value="22">22</option>
					<option value="23">23</option>
					<option value="24">24</option>
                    <option value="25">25</option>
					<option value="26">26</option>
					<option value="27">27</option>
                    <option value="28">28</option>
					<option value="29">29</option>
					<option value="30">30</option>
					<option value="31">31</option>
                    <option value="32">32</option>
					<option value="33">33</option>
					<option value="34">34</option>
                    <option value="35">35</option>
					<option value="36">36</option>
					<option value="37">37</option>
                    <option value="38">38</option>
					<option value="39">39</option>
					<option value="40">40</option>
					<option value="41">41</option>
                    <option value="42">42</option>
					<option value="43">43</option>
					<option value="44">44</option>
                    <option value="45">45</option>
					<option value="46">46</option>
					<option value="47">47</option>
                    <option value="48">48</option>
					<option value="49">49</option>
					<option value="50">50</option>
					
					
					
                  </select></td>
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_andares_edificio" value="<%=rs444Comprador("obs_andares_edificio")%>" type="text" id="txt_obs_andares_edificio" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quantidade_elevadores" id="txt_quantidade_elevadores" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                      <option value="<% if rs444Comprador("quantidade_elevadores") <> "não informado" and rs444Comprador("quantidade_elevadores") <> "" then response.write rs444Comprador("quantidade_elevadores") else response.write "0" end if%>" selected>
                      <% if rs444Comprador("quantidade_elevadores") <> "não informado" and rs444Comprador("quantidade_elevadores") <> "" then response.write rs444Comprador("quantidade_elevadores") else response.write "0" end if%>
                      </option>
                      <option value="00">00</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                      <option value="10">10</option>
                    </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quantidade_elevadores" value="<%=rs444Comprador("obs_quantidade_elevadores")%>" type="text" id="txt_obs_quantidade_elevadores" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_portaria" id="txt_portaria" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
                  
				   <option value="<% if rs444Comprador("portaria") <> "não informado" and rs444Comprador("portaria") <> "" then response.write rs444Comprador("portaria") else response.write "0" end if%>" selected><% if rs444Comprador("portaria") <> "não informado" and rs444Comprador("portaria") <> "" then response.write rs444Comprador("portaria") else response.write "0" end if%></option>
					
				  
				  
				      <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
					<option value="03">03</option>
					<option value="04">04</option>
                    <option value="05">05</option>
					<option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
					<option value="09">09</option>
					<option value="10">10</option>
                    
					
					
                  </select></td>
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_portaria" value="<%=rs444Comprador("obs_portaria")%>" type="text" id="txt_obs_portaria" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
			
			
			
			
			
        </tr>
      </table>
        </td>
  </tr>
  <tr><td height="30">
  <table cellpadding="0" cellspacing="0" border="0">
  <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Condom&iacute;nio</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Condi&ccedil;&otilde;es 
              de pagamento</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
              Total </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
              &Uacute;til </font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="stage22" type="text" class="inputBox" id="stage22" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=formatnumber(rs444Comprador("valor"),2)%>" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_condominio" type="text" class="inputBox" id="txt_condominio" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=formatnumber(rs444Comprador("condominio"),2)%>" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_condicoes_pagamento" size="1" class="inputBox" id="txt_condicoes_pagamento"  style="HEIGHT: 18px; WIDTH: 190px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
              
			  <option value="<%=rs444Comprador("condicoes_pagamento")%>" selected><%=rs444Comprador("condicoes_pagamento")%></option>
					  
			   <option value="não informado">não informado</option>
			    <option value="à vista">à vista</option>
                <option  value="à vista e parte parcelado">à vista e parte parcelado </option>
                <option value="à vista e parte FGTS" >à vista e parte FGTS </option>
                <option  value="à vista , FGTS e financiamento">à vista , FGTS e financiamento" </option>
                <option value="FGTS" >FGTS </option>
                <option  value="financiamento">financiamento</option>
                <option value="parcelado" >parcelado </option>
                <option value="parte à vista mais imóvel">parte à vista mais imóvel</option>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
              <input name="txt_area_total" type="text" class="inputBox" id="txt_area_total" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%if rs444Comprador("area_total") <> "" then response.write rs444Comprador("area_total") else response.write "0" end if%>" size="12" maxlength="20">
              </font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_area_construida" type="text" class="inputBox" id="txt_area_construida" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="<%if rs444Comprador("area_construida") <> "" then response.write rs444Comprador("area_construida") else response.write "0" end if%>" size="12" maxlength="20"></td>
        </tr>
  
  </table>
  
  
  </td></tr>
  
  <tr><td>
  <table cellpadding="0" cellspacing="0" border="0">
  <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Situa&ccedil;&atilde;o 
              do cliente</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_standby" id="txt_standby" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
               
			   
			     <option value="<%if rs444Comprador("standby") <> "" then response.write rs444Comprador("standby") else response.write "não informado" end if %>" selected>
                <%if rs444Comprador("standby") <> "" then response.write rs444Comprador("standby") else response.write "não informado" end if %>
                </option>
			   
			    <option value="não informado">não informado</option>
                <option value="comprou com a Veja">comprou com a Veja</option>
                <option value="comprou com outro">comprou com outro</option>
                <option value="suspenso">suspenso</option>
                <option value="cliente inexistente">cliente inexistente</option>
                <option value="cliente para permuta">cliente para permuta</option>
                <option value="cliente com proposta">cliente com proposta</option>
				<option value="comprador OK" >comprador OK</option>
              <option value="comprador a contatar" >comprador a contatar</option>
			  </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" >&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" >&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" >&nbsp;</td>
          <td width="10">&nbsp;</td>
          <td width="192"  >&nbsp;</td>
        </tr>
  
  </table>
  
  </td></tr>
  
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dados 
              confidenciais desse cliente</font></td>
          <td width="10">&nbsp;</td>
            <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
              do im&oacute;vel desejado e forma de pagamento desse cliente.</font></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td width="495" height="80" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_descricao_confi" cols=20 rows=10 class="inputBox" id="textarea3" style="HEIGHT: 78px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%=rs444Comprador("descricao_confi")%></textarea></td>
          <td width="10">&nbsp;</td>
            <td width="495" height="80" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_descricao" COLS=20 ROWS=10 class="inputBox" id="txt_descricao" style="HEIGHT: 78px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%=rs444Comprador("descricao")%></textarea></td>
        </tr>
      </table></td>
  </tr>
  
  
  <% else %>
  
  
  
  <% end if %>
  
 
  
  
  <tr> 
    <td><div align="right">
          <input name="image" type="image"  src="bt_atualizar0011.jpg" width="145" height="18" border="0">
          <a href="atualizar_imovel33.asp?varCod_imovel=<%=rs("cod_imovel")%>&varNaoContatado=sim"><img src="bt_naocontatado0011.jpg" width="145" height="18" border="0"></a></div></td>
  </tr>
  
  
  
</table>
</form>




<%


 '-------------------------Colocar tarja verde nas contas procuradas----------- 
  
  dim rs444IndicacaoTarja01
  dim strSQL444IndicacaoTarja01
  dim vNumeroTarja01
  
  
  Set rs444IndicacaoTarja01 = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444IndicacaoTarja01 = "SELECT * from contas_procuradas where telefone like '"&rs("telefone")&"' or telefone like '"&rs("telefone02")&"' or telefone like '"&rs("telefone03")&"'"

         
         rs444IndicacaoTarja01.CursorLocation = 3
        rs444IndicacaoTarja01.CursorType = 3
        
 
            rs444IndicacaoTarja01.ActiveConnection = Conexao




	rs444IndicacaoTarja01.Open strSQL444IndicacaoTarja01, Conexao
	
	
	
	
	
	dim numTarja01
	numTarja01 = 0
	
	 if not rs444IndicacaoTarja01.eof  then 
	 vNumeroTarja01 = rs444IndicacaoTarja01.recordcount
	 
				     While   numTarja01 < int(vNumeroTarja01)
                   
                    'Conexao.execute"update imoveis set imovel_em_negociacao='"&"não informado"&"' where imovel_em_negociacao IS NULL" 
                   
				    Conexao.execute"update  contas_procuradas set Tarja01='"&"sim"&"' where cod_conta="&rs444IndicacaoTarja01("cod_conta")
                   
				   
				   
				   numTarja01 = numTarja01 +1
                   rs444IndicacaoTarja01.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	
  
  
  
  '-----------------------------------------------------------------------




%>






<%

'-------------------------------
   rs444IndicacaoTarja01.close
  
  
  set rs444IndicacaoTarja01 = nothing
 '------------------------------

%>






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


<%

'-----------------------------
           rs.Close           
		   
           Set rs = Nothing
		   
'---------------------------------





'-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------

'-----------------------------
            rs4.close        
		   
           Set rs4 = Nothing
		   
'---------------------------------




'-----------------------------
           rs33.Close           
		   
           Set rs33 = Nothing
		   
'---------------------------------


'-----------------------------
           rs44.Close           
		   
           Set rs44 = Nothing
		   
'---------------------------------




'-----------------------------
           rs333.Close           
		   
           Set rs333 = Nothing
		   
'---------------------------------

































'-----------------------------
           rs444Placa.Close           
		   
           Set rs444Placa = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Captacao.Close           
		   
           Set rs444Captacao = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Captacao22.Close           
		   
           Set rs444Captacao22 = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Origem.Close           
		   
           Set rs444Origem = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Tipo22.Close           
		   
           Set rs444Tipo22 = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Tipo23.Close           
		   
           Set rs444Tipo23 = Nothing
		   
'---------------------------------





'-----------------------------
           rs444Responsavel22.Close           
		   
           Set rs444Responsavel22 = Nothing
		   
'---------------------------------



'-----------------------------
           rs444ConseguiuProposta22.Close           
		   
           Set rs444ConseguiuProposta22 = Nothing
		   
'---------------------------------





'-----------------------------
           rs444Responsavel.Close           
		   
           Set rs444Responsavel = Nothing
		   
'---------------------------------





%>





 <% response.flush%>
  <%response.clear%>
  <%  EscreveFuncaoJavaScript2 ( Conexao3) %>
  <%EscreveFuncaoJavaScript999 ( Conexao3 )%>
  <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>




 <%
conexao3.close
set conexao3 = nothing


conexao333.close
set conexao333 = nothing


conexao.close
set conexao = nothing



set conexao33 = nothing



%>

	
	

</body>
</html>
