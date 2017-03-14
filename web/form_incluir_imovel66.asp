<!--#include file="dsn.asp"-->
<!--#include file="style_imoveis.asp"-->

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
	
	
	
	
	
	
	
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
	
		
	
	
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
	 
	 
	'-------------------------quem tirou a foto------------------------
	
	
	
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







	
	 
%>	





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Incluir imóvel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script>
function isValidDigitNumber (doublecombo)
{


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
	
	
	
	
	if (doublecombo.txt_tipo_vend.value == "Apartamento" && doublecombo.txt_condominio_vend.value == "0,00" ) {
        alert("Você precisa indicar o valor do condomínio!");
        doublecombo.txt_condominio_vend.focus();
		doublecombo.txt_condominio_vend.select();
        return false;
    }
	
	
	
	
	
	
	
	//------------Compradores condominio,área total e área construida----------------------

var strValidNumber8_ja="1234567890";
for (nCount=0; nCount < doublecombo.txt_area_total.value.length; nCount++) 
		{
strTempChar8_ja=doublecombo.txt_area_total.value.substring(nCount,nCount+1);
if (strValidNumber8_ja.indexOf(strTempChar8_ja,0)==-1) 
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
if (strValidNumber9_ja.indexOf(strTempChar9_ja,0)==-1) 
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
if (strValidNumber99_ja.indexOf(strTempChar99_ja,0)==-1) 
{
alert("O item condomínio só pode conter números e uma vírgula!");
doublecombo.txt_condominio.focus();
doublecombo.txt_condominio.select();
return false;
}
}
	
	
	
	
	
	


var strText2_4 = doublecombo.stage22.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.stage22.focus();
		
		doublecombo.stage22.select();
		
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






</head>

<body bgcolor="<%=escuro%>">





<div align="center"> <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi incluido com sucesso.</font> 
          <% end if %>
		  </div>
		  
		  
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="incluir_imovel33.asp">
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
          <td width="26"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_futuro_contato_vend" type="text" class="inputBox" id="txt_data_futuro_contato_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=now()%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
            <td width="798" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_assunto_futuro_contato_vend" type="text" class="inputBox" id="txt_assunto_futuro_contato_vend" style="HEIGHT: 18px; WIDTH: 796px; background:<%=medio%>" size="38" maxlength="50" align="left">
            </td>
          
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
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">n&atilde;o 
              informado </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
              <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
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
                <option value="aluguel">Aluguel</option>
                <option value="venda" selected>Venda</option>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_captacao_vend" type="text" class="inputBox" id="txt_data_captacao_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=now()%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_origem_captacao_vend" value="<%=session("nome_id")%>" type="text" id="txt_origem_captacao_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>"></td>
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
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
            &uacute;ltima atualiza&ccedil;&atilde;o</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
            cadastramento </font></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
				
				<select name="txt_captacao_vend" id="txt_captacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                 <option value="<%=session("nome_id")%>" selected ><%=session("nome_id")%></option>
                
				<option value="internet" >internet</option>
					  <option value="Internet">Internet</option>
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
                    </select>
					</td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input type="text" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=session("nome_id")%>"  size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input type="text" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=now()%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input type="text" class="inputBox" id="txt_data_captacao3" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="<%=now()%>" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_responsavel_cadastramento_vend" id="txt_responsavel_cadastramento_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="<%=session("nome_id")%>" selected ><%=session("nome_id")%></option>
               
              </select></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Propriet&aacute;rio</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
              1</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
              2 </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
              3</font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></td>
        </tr>
        <tr> 
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario_vend" type="text" id="txt_proprietario_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone_vend" type="text" id="txt_telefone_vend" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone03_vend" type="text" id="txt_telefone03_vend" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone02_vend" type="text" id="txt_telefone02_vend" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email_vend" type="text" id="txt_email_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px ; background:<%=medio%>"></td>
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_chaves_do_imovel_vend" type="text" id="txt_chaves_do_imovel_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_qualidade_vend" id="txt_qualidade_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="bom negócio" >Bom Negócio</option>
                <option value="negócio comum" selected>Negócio Comum</option>
              </select></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_melhor_horario_visita_vend" size="1" class="inputBox" id="txt_melhor_horario_visita_vend"  style="HEIGHT: 18px; WIDTH: 190px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
               
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_placa_vend" id="txt_placa_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="Sem Placa" selected >Sem Placa</option>
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
                <option value="não informado" selected>não informado</option>
                <option value="vago">vago</option>
                <option value="alugado">alugado</option>
                <option value="ocupado por terceiros">ocupado por terceiros</option>
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
                      <option value="cqualquer" selected>Cidade</option>
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
                <option value="bqualquer" selected>Bairro/Região</option>
                <% if not rs44.eof then%>
                <% While NOT Rs44.EoF %>
                <option value="<% = Rs44("id_combo2") %>"> 
                <% = Rs44("nome_combo2") %>
                </option>
                <% Rs44.MoveNext %>
                <% Wend %>
                <% else %>
                <option value=""></option>
                <% end if %>
              </select></td>
            <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                <option value="vlqualquer" selected>Vila</option>
                <option value="vlqualquer">qualquer um</option>
              </select></td>
            <td width="10">&nbsp;</td>
            <td width="394" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco_vend" type="text" id="txt_endereco_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 390px; background:<%=medio%>"></td>
          </tr>
        </table></td>
  </tr>
  <tr> 
    <td> <table width="1000" cellpadding="0" cellspacing="0">
        <tr> 
            <td width="192" height="21" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_condominio_vend" type="text" class="inputBox" id="txt_condominio_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_iptu_vend" type="text" class="inputBox" id="txt_valor_iptu_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_outros_vend" type="text" class="inputBox" id="txt_valor_outros_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_rateio_vend" type="text" class="inputBox" id="txt_rateio_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" size="12" maxlength="13"></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
              devedor</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
              devedor j&aacute; pago</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
              devedor a pagar</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">.&Agrave;rea 
              total apartamento ou terreno</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
              constru&iacute;da p/casas ou &uacute;til apart.</font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_saldo_devedor_vend" id="txt_saldo_devedor_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="sem saldo devedor" selected>Sem saldo devedor</option>
                <option value="com saldo devedor" >Com saldo devedor</option>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_ja_pago_devedor_vend" type="text" class="inputBox" id="txt_ja_pago_devedor_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="30"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
              <input name="txt_devendo_devedor_vend" type="text" class="inputBox" id="txt_devendo_devedor_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="30">
              </td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
              <input name="txt_a_total_vend" type="text" class="inputBox" id="txt_a_total_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="00" size="12" maxlength="20">
              </font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_a_constr_vend" type="text" class="inputBox" id="txt_a_constr_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="00" size="12" maxlength="20"></td>
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
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;Nome 
              do edif&iacute;cio</font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_de_frente_vend" type="text" class="inputBox" id="txt_metros_de_frente_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="00" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_de_fundo_vend" type="text" class="inputBox" id="txt_metros_de_fundo_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="00" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_lateral_esquerda_vend" type="text" class="inputBox" id="txt_metros_lateral_esquerda_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="00" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_metros_lateral_direita_vend" type="text" class="inputBox" id="txt_metros_lateral_direita_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="00" size="12" maxlength="20"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_nome_edificio_vend" type="text" class="inputBox" id="txt_nome_edificio_vend" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" size="12" maxlength="20"></td>
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
              conseguiu a proposta</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">vendido/suspenso/proposta&nbsp; 
              </font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_presenca_primeira_vend" size="1" class="inputBox" id="txt_presenca_primeira_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                <% if session("permissao") <> "6" or session("permissao") <> "5" then %>
                <option value="excluido">Excluído</option>
                <%else%>
                <option value="incluido">Incluído</option>
                <option value="excluido" selected>Excluído</option>
                <%end if%>
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_titulo_anuncio_vend" type="text" id="txt_titulo_anuncio_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_texto_anuncio_vend" type="text" id="txt_texto_anuncio_vend" size="38" maxlength="120" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>" style="border:1px solid #FFFFFF;"><select name="txt_conseguiu_proposta_vend" id="txt_conseguiu_proposta_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="Internet" selected >Internet</option>
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
                <option value="não informado" selected>não informado</option>
                <option value="Vendido pela Veja">Vendido pela Veja</option>
                <option value="Vendido por outros">Vendido por outros</option>
                <option value="Suspenso">Suspenso</option>
                <option value="Com proposta">Com proposta</option>
                <option value="Imóvel inexistente">Imóvel inexistente</option>
                <option value="Imóvel para permuta">Imóvel para permuta</option>
                <option value="alugado pela Veja" >alugado pela Veja</option>
				<option value="imóvel OK" >imóvel OK</option>
              <option value="Alugado por outros" >Alugado por outros</option>
			  </select></td>
        </tr>
  
  
  </table>
  
  </td>
  
  </tr>
  <tr><td>
  <table cellpadding="0" cellspacing="0" border="0">
  
  
  <tr> 
            <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quem 
              tirou a foto</font></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
        </tr>
        <tr> 
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_quem_tirou_foto_vend" id="txt_quem_tirou_foto_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>">
                <option value="não informado" selected >não informado</option>
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
              </select></td>
          <td width="10">&nbsp;</td>
            <td width="192" >&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" >&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" >&nbsp;</td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=escuro%>" >&nbsp;</td>
        </tr>
  
  
  </table>
  
  </td></tr>
  
  
  
  
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quartos_vend" type="text" id="txt_obs_quartos_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_vagas_vend" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_vagas_vend" type="text" id="txt_obs_vagas_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_banheiros_vend" id="txt_banheiros_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_banheiros_vend" type="text" id="txt_obs_banheiros_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_suites_vend" type="text" id="txt_obs_suites_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_edicula_vend" id="txt_edicula_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_edicula_vend" type="text" id="txt_obs_edicula_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_entrada_lateral_vend" id="txt_entrada_lateral_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_entrada_lateral_vend" type="text" id="txt_obs_entrada_lateral_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_festas_vend" type="text" id="txt_obs_salao_de_festas_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_salao_de_jogos_vend" id="txt_salao_de_jogos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_jogos_vend" type="text" id="txt_obs_salao_de_jogos_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_churrasqueira_vend" id="txt_churrasqueira_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_churrasqueira_vend" type="text" id="txt_obs_churrasqueira_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_piscina_vend" id="txt_piscina_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_piscina_vend" type="text" id="txt_obs_piscina_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quintal_vend" id="txt_quintal_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quintal_vend" type="text" id="txt_obs_quintal_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quadras_vend" id="txt_quadras_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quadras_vend" type="text" id="txt_obs_quadras_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_andares_edificio_vend" type="text" id="txt_obs_andares_edificio_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quantidade_elevadores_vend" id="txt_quantidade_elevadores_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quantidade_elevadores_vend" type="text" id="txt_obs_quantidade_elevadores_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_portaria_vend" id="txt_portaria_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_portaria_vend" type="text" id="txt_obs_portaria_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
            confidenciais sobre o propriet&aacute;rio</font></td>
          <td width="10">&nbsp;</td>
            <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
              sobre o im&oacute;vel e forma de pagamento que o propriet&aacute;rio 
              deseja</font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="495" height="102" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_obs_proprietario_vend" class="inputBox" id="txt_obs_proprietario_vend" style="HEIGHT: 100px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
          <td width="10">&nbsp;</td>
          <td width="495" height="80" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_obs_imovel_vend" class="inputBox" id="txt_obs_imovel_vend" style="HEIGHT: 100px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
        </tr>
      </table></td>
  </tr>
  
  <tr> 
    <td height="300"><div align="center">
	
	
	<table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
              <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Caso 
                o propriet&aacute;rio desse im&oacute;vel seja tamb&eacute;m um 
                comprador de im&oacute;veis, selecione sim e preencha a ficha 
                abaixo. </strong></font></td>
  </tr>
</table>

	
	<br><br> 
	
          <select name="txt_pergunta" id="txt_pergunta" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
           
		   <option value="não informado" selected>não informado</option>
			<option value="sim">Sim</option>
            <option value="nao">Não</option>
           
          </select>
        </div></td>
  </tr>
  
  
  
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_futuro_contato_comprador" type="text" class="inputBox" id="txt_data_futuro_contato_comprador" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0/0/2007 00:00:00" size="38" maxlength="50" align="left"></td>
          <td width="10">&nbsp;</td>
          <td width="798" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_assunto_futuro_contato_comprador" type="text" class="inputBox" id="txt_assunto_futuro_contato_comprador" style="HEIGHT: 18px; WIDTH: 796px; background:<%=medio%>" size="38" maxlength="50" align="left"></td>
          
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_origem" id="txt_origem" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_responsavel_cadastramento_comprador" id="txt_responsavel_cadastramento_comprador" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
               
			    <option value="<%=session("nome_id")%>" selected ><%=session("nome_id")%></option>
                
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
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_atendimento" id="txt_atendimento" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     
					  <option value="<%=session("nome_id")%>" selected ><%=session("nome_id")%></option>
                
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
          <td width="192" bgcolor="<%=medio%>" valign="top" style="border:1px solid #FFFFFF;" ><select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>">
                <option value="cqualquer" selected>Cidade</option>
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
            <td width="192" height="120" valign="top" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><select name="combo2" class="inputBox" style="HEIGHT: 130px; WIDTH: 190px; background:<%=medio%>" multiple size="8">
                <option value="bqualquer">Bairro/Região</option>
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
                <option  value="aluguel">Aluguel </option>
                <option value="compra" selected>Compra </option>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quartos" type="text" id="txt_obs_quartos" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_vagas" type="text" id="txt_obs_vagas" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_banheiros" id="txt_banheiros" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_banheiros" type="text" id="txt_obs_banheiros" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_suites" type="text" id="txt_obs_suites" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_edicula" id="txt_edicula" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_edicula" type="text" id="txt_obs_edicula" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_entrada_lateral" id="txt_entrada_lateral" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_entrada_lateral" type="text" id="txt_obs_entrada_lateral" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_festas" type="text" id="txt_obs_salao_de_festas" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_salao_de_jogos" id="txt_salao_de_jogos" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_salao_de_jogos" type="text" id="txt_obs_salao_de_jogos" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_churrasqueira" id="txt_churrasqueira" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_churrasqueira" type="text" id="txt_obs_churrasqueira" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_piscina" type="text" id="txt_obs_piscina" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quintal" id="txt_quintal" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quintal" type="text" id="txt_obs_quintal" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quadras" id="txt_quadras" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quadras" type="text" id="txt_obs_quadras" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
                <td width="276" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_andares_edificio" type="text" id="txt_obs_andares_edificio" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_quantidade_elevadores" id="txt_quantidade_elevadores" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_quantidade_elevadores" type="text" id="txt_obs_quantidade_elevadores" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
          <td width="327"><table width="327" border="0" cellspacing="0" cellpadding="0">
              <tr>
                 <td width="50" style="border:1px solid #FFFFFF;"><select name="txt_portaria" id="txt_portaria" class="inputBox" style="HEIGHT: 18px; WIDTH: 50px; background: <%=medio%>">
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
                <td width="277" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_obs_portaria" type="text" id="txt_obs_portaria" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 274px; background:<%=medio%>"></td>
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
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="stage22" type="text" class="inputBox" id="stage22" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_condominio" type="text" class="inputBox" id="txt_condominio" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="0,00" size="12" maxlength="13"></td>
          <td width="10">&nbsp;</td>
            <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_condicoes_pagamento" size="1" class="inputBox" id="txt_condicoes_pagamento"  style="HEIGHT: 18px; WIDTH: 190px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
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
              <input name="txt_area_total" type="text" class="inputBox" id="txt_area_total" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>" value="00" size="12" maxlength="20">
              </font></td>
          <td width="10">&nbsp;</td>
          <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_area_construida" type="text" class="inputBox" id="txt_area_construida" style="HEIGHT: 18px; WIDTH: 190px; background: <%=medio%>" value="00" size="12" maxlength="20"></td>
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
                <option value="não informado" selected>não informado</option>
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
            <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es 
              do im&oacute;vel desejado e forma de pagamento desse cliente.</font></td>
          <td width="10">&nbsp;</td>
          <td width="495"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dados 
            confidenciais desse cliente</font></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr> 
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="495" height="80" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_descricao" COLS=20 ROWS=10 class="inputBox" id="txt_descricao" style="HEIGHT: 78px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
          <td width="10">&nbsp;</td>
          <td width="495" height="80" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_descricao_confi" COLS=20 ROWS=10 class="inputBox" id="txt_descricao_confi" style="HEIGHT: 78px; WIDTH: 493px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr> 
    <td><div align="right">
          <input name="image" type="image"  src="bt_enviar0011.jpg" width="145" height="18" border="0"><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar0011.jpg" width="145" height="18" border="0"></a></div></td>
  </tr>
  
  
  
</table>
</form>



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
