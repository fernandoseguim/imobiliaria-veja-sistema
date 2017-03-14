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
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
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
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 
Sql5 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rs5 = Server.CreateObject("ADODB.RecordSet")

	rs5.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs5.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs5.ActiveConnection = Conexao3
	
	
	rs5.Open Sql5, Conexao3



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
SqlMarcas4 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


Set rsMarcas4 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas4.ActiveConnection = Conexao3
	
	
	rsMarcas4.Open SqlMarcas4, Conexao3


While NOT rsMarcas4.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas4("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas4("id_combo1")&" order by nome_combo2"


Set rsCarros4 = Server.CreateObject("ADODB.RecordSet")

	rsCarros4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros4.ActiveConnection = Conexao3
	
	
	rsCarros4.Open SqlCarros4, Conexao3

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros4.EoF

Response.Write "doublecombo.combo4.options[" & i & "] = new Option('" & rsCarros4("nome_combo2") & "','" & rsCarros4("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros4.MoveNext
Wend
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas4.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



rsMarcas4.close

set rsMarcas4 = nothing

rsCarros4.close

set rsCarros4 = nothing




End Function
%> 




















<%
'Criando conexão com o banco de dados! 

'Abrindo a tabela MARCAS!
Sql4 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 

Set rs4 = Server.CreateObject("ADODB.RecordSet")

	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao3
	
	
	rs4.Open Sql4, Conexao3
	

%> 














<!--#include file="cores03.asp"-->


<% response.buffer=True%>



<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")


 dim Conexao9,rs9

	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL9
	dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	 strSQL9 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where cod_permuta="&varCodPermuta
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	 rs9.Open strSQL9, Conexao3
	



dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where nome_combo3 ='"&rs9("vila_vend")&"' and bairro_combo3 ='"&rs9("bairro_vend")&"' and cidade_combo3 ='"&rs9("cidade_vend")&"'  ORDER BY nome_combo3" 
	 
	 
	 rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444.ActiveConnection = Conexao3
	 
	 
	 
	 
	 rs444.Open strSQL444, Conexao3		
	





dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where nome_combo3 ='"&rs9("vila_comp")&"' and bairro_combo3 ='"&rs9("bairro_comp")&"' and cidade_combo3 ='"&rs9("cidade_comp")&"'  ORDER BY nome_combo3" 
	
	
	rs555.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs555.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs555.ActiveConnection = Conexao3
	
	
	 rs555.Open strSQL555, Conexao3		
	








   
   Set rs = Server.CreateObject("ADODB.RecordSet")
   dim rs4,strSQL4,strSQL6,rs6
   
    
	Set rs6 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 like '"& rs9("bairro_vend") &"' ORDER BY nome_combo2" 
	strSQL6 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 like '"& rs9("bairro_comp") &"' ORDER BY nome_combo2"  
	
    
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
  
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao3 
		
		rs6.Open strSQL6, Conexao3
		
	
	 dim Conexao2,rs7
 
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL7
	
	if rs9("cod_imovel") <> "não informado" then
	
	 strSQL7 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where cod_imovel="&rs9("cod_imovel")
	 rs7.CursorLocation = 3
      rs7.CursorType = 3
	 rs7.Open strSQL7, Conexao3
   if not rs7.eof then
   vimagem = rs7("foto_grande")
   else
   vimagem = "imovel00000.jpg"
  end if
	
	else
	
	vimagem = "imovel00000.jpg"
	
	end if
	
	
	
	
		
%>		




<%
Function EscreveFuncaoJavaScript888 ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros888 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo2.options[doublecombo.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 ORDER BY nome_combo2" 


Set rsMarcas888 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas888.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas888.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas888.ActiveConnection = Conexao3
	
	
	rsMarcas888.Open SqlMarcas888, Conexao3


While NOT rsMarcas888.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas888("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros888 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where id_combo2 =" & rsMarcas888("id_combo2")&""

Set rsCarros888 = Server.CreateObject("ADODB.RecordSet")

	rsCarros888.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros888.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros888.ActiveConnection = Conexao3
	
	
	rsCarros888.Open SqlCarros888, Conexao3
'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros888.EoF

Response.Write "doublecombo.combo5.options[" & i & "] = new Option('" & rsCarros888("nome_combo3") & "','" & rsCarros888("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros888.MoveNext
Wend


Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas888.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 


rsMarcas888.close

set rsMarcas888 = nothing

rsCarros888.close


set rsCarros888 = nothing




End Function
%> 


<%

'Criando conexão com o banco de dados! 


'

Sql888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 


Set rs888 = Server.CreateObject("ADODB.RecordSet")

	rs888.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs888.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs888.ActiveConnection = Conexao3
	
	
	rs888.Open Sql888, Conexao3
	



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
SqlMarcas999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 


Set rsMarcas999 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas999.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas999.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas999.ActiveConnection = Conexao3
	
	
	rsMarcas999.Open SqlMarcas999, Conexao3



While NOT rsMarcas999.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""

Set rsCarros999 = Server.CreateObject("ADODB.RecordSet")

	rsCarros999.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros999.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros999.ActiveConnection = Conexao3
	
	
	rsCarros999.Open SqlCarros999, Conexao3
'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo7.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros999.EoF

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




rsMarcas999.close


set rsMarcas999 = nothing


rsCarros999.close


set rsCarros999 = nothing

End Function
%> 


<%

'Criando conexão com o banco de dados! 


'

Sql999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 

Set rs999 = Server.CreateObject("ADODB.RecordSet")

	rs999.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs999.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs999.ActiveConnection = Conexao3
	
	
	rs999.Open Sql999, Conexao3
	



 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")




dim rs666,strSQL666
   
    Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL666 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 where nome_combo1 ='"&rs9("cidade_vend")&"'  ORDER BY nome_combo1" 
	
	
	rs666.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs666.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs666.ActiveConnection = Conexao3
	
	
	
	 rs666.Open strSQL666, Conexao3		


dim rs777,strSQL777
   
    Set rs777 = Server.CreateObject("ADODB.RecordSet")
	strSQL777 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 where nome_combo1 ='"&rs9("cidade_comp")&"'  ORDER BY nome_combo1" 
	
	
	rs777.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs777.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs777.ActiveConnection = Conexao3
	
	
	
	 rs777.Open strSQL777, Conexao3		



dim rs8888,strSQL8888
   
    Set rs8888 = Server.CreateObject("ADODB.RecordSet")
	strSQL8888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 ='"&rs9("bairro_vend")&"' and cidade_combo2 ='"&rs9("cidade_vend")&"'  ORDER BY nome_combo2" 
	
	
	rs8888.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs8888.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs8888.ActiveConnection = Conexao3
	
	
	 rs8888.Open strSQL8888, Conexao3		



dim rs9999,strSQL9999
   
    Set rs9999 = Server.CreateObject("ADODB.RecordSet")
	strSQL9999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 ='"&rs9("bairro_comp")&"' and cidade_combo2 ='"&rs9("cidade_comp")&"'  ORDER BY nome_combo2" 
	
	
	rs9999.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs9999.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs9999.ActiveConnection = Conexao3
	
	
	
	 rs9999.Open strSQL9999, Conexao3






'---------------------Jogar valores nas variáveis---------------------------







 dim varIndicacaoCidadeVend
 dim varIndicacaoBairroVend
 dim varIndicacaoVilaVend
 dim varIndicacaoQuartosVend
 dim varIndicacaoVagasVend
 dim varIndicacaoValorVend
 dim varIndicacaoTipoVend


 dim varIndicacaoCidadeComp
 dim varIndicacaoBairroComp
 dim varIndicacaoVilaComp
 dim varIndicacaoQuartosComp
 dim varIndicacaoVagasComp
 dim varIndicacaoValorComp
 dim varIndicacaoTipoComp
 dim varIndicacaoCodigo	








      varIndicacaoCodigo = rs9("cod_permuta")
	 varIndicacaoCidadeVend=rs9("cidade_vend")
 varIndicacaoBairroVend=rs9("bairro_vend")
 varIndicacaoVilaVend=rs9("vila_vend")
 varIndicacaoQuartosVend=rs9("quartos_vend")
 varIndicacaoVagasVend=rs9("vagas_vend")
 varIndicacaoValorVend=rs9("valor_vend")
 varIndicacaoTipoVend=rs9("tipo_vend")


 varIndicacaoCidadeComp=rs9("cidade_comp")
 varIndicacaoBairroComp=rs9("bairro_comp")
 varIndicacaoVilaComp=rs9("vila_comp")
 varIndicacaoQuartosComp=rs9("quartos_comp")
 varIndicacaoVagasComp=rs9("vagas_comp")
 varIndicacaoValorComp=rs9("valor_comp")
 varIndicacaoTipoComp=rs9("tipo_comp")
	





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







 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	
	rs444Tipo23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo23.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo23.ActiveConnection = Conexao3
	
	
	 rs444Tipo23.Open strSQL444Tipo23, Conexao3




'-------------------------------------------------------------------------------------------------




'--------------------atualizar ultimo acesso-------------------------




dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou FROM compradores where telefone ='"&varTelefone&"'" 
	
	
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



	
var strValidNumber1_77="1234567890,";
for (nCount=0; nCount < doublecombo.txt_cod_imovel.value.length; nCount++) 
		{
strTempChar1_77=doublecombo.txt_cod_imovel.value.substring(nCount,nCount+1);
if (strValidNumber1_77.indexOf(strTempChar1_77,0)==-1) 
{
alert("O formulário cod imovel só pode conter números!");
doublecombo.txt_cod_imovel.focus();
doublecombo.txt_cod_imovel.select();
return false;
}
}






if (doublecombo.txt_proprietario.value == "") {
        alert("Você precisa indicar o nome do proprietário!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do proprietário!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
    }
	
	
	
	
	
var strValidNumber1_7="1234567890,";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Telefone só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}


if (doublecombo.txt_endereco.value == "") {
        alert("Você precisa indicar o endereço do proprietário!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }




if (doublecombo.txt_valor_vend.value == "") {
        alert("O formulário valor do seu Imóvel está vazio!");
        doublecombo.txt_valor_vend.focus();
		doublecombo.txt_valor_vend.select();
        return false;
    }
	
	
	if (doublecombo.txt_valor_comp.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.txt_valor_comp.focus();
		doublecombo.txt_valor_comp.select();
        return false;
    }


var strText2_4 = doublecombo.txt_valor_vend.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.txt_valor_vend.focus();
		
		doublecombo.txt_valor_vend.select();
		
return false;

}



var strText2_5 = doublecombo.txt_valor_comp.value;
var s_strText2_5 = strText2_5.length
if (strText2_5.substring((s_strText2_5 - 3), (s_strText2_5 - 2)) != ","){

       alert("A vírgula do formulário Valor do imóvel pretendido está fora do lugar!");
       doublecombo.txt_valor_comp.focus();
		
		doublecombo.txt_valor_comp.select();
		
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




<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow444(abrejanela444) {
   openWindow444 = window.open(abrejanela444,'openWin444','width=603,height=500,resizable=yes,scrollbars=yes,left=100')
   openWindow444.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow555(abrejanela555) {
   openWindow555 = window.open(abrejanela555,'openWin555','width=603,height=500,resizable=yes,scrollbars=yes,left=100')
   openWindow555.focus( )
   }

</SCRIPT>


</head>

<!--#include file="style_imoveis02.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >

<center>
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_permuta02.asp?varCodPermuta=<%=varCodPermuta%>"><table width="590" border="0" cellspacing="0" cellpadding="0"> 
  <table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="590" height="48"><a href="acessoLink01.asp?varTelefone=<%=session("varTelefone")%>"><img  src="bt_voltar002.jpg" border="0" ></img></a></td>
  </tr>
  
   
  
  
  
  
 
  <tr>
    <td width="590" height="18"><table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td width="5">&nbsp;</td>
    <td><div align="center"> </div></td>
          <td width="5">&nbsp;</td>
  </tr>
</table>
</td>
  </tr>
  <tr>
      <td width="590" height="190"><table width="590" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="190">&nbsp;</td>
            <td width="580" height="190" style="border:1px solid #FFFFFF;"><table width="580" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  
                <td width="290" height="190" bgcolor="<%=medio%>" >&nbsp;</td>
                  <td width="290" height="190" ><table width="290" height="190" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        
                      <td width="290" height="170"><% if vimagem <> "imovel00000.jpg" then %><% If objFSO.FileExists(Server.MapPath(vimagem)) = True Then%><img src="<%=vimagem%>" name="photoslider" width="290" height="170"></img><% else %><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div><% end if %><% else %><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div><% end if %></td>
                      </tr>
                      <tr>
                        
                      <td width="290" height="20" bgcolor="<%=claro%>" >
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                          do im&oacute;vel</font></div></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
            <td width="5" height="190">&nbsp;</td>
          </tr>
        </table></td>
  </tr>
  
  
  
  <tr>
    <td>&nbsp;<div align="center">
          <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi atualizado  com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
  
  
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atendimento</font></div></td>
                      <td height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><% if session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then%><input name="txt_atendimento" type="text" class="inputBox" id="txt_atendimento" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left"><%else%><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("atendimento")%></font><input name="txt_atendimento" type="hidden" class="inputBox" id="txt_atendimento" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left"><%end if%></td>
              </tr>
			
			 <tr>
                      <td height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                          de inclus&atilde;o</font></div></td>
                      <td height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs9("data")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			   <tr>
                      <td height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                          da &uacute;ltima atualiza&ccedil;&atilde;o</font></div></td>
                      <td height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("data_atualizacao")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			
			  <tr>
                      <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                          da permuta</font></div></td>
                      <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rs9("cod_permuta")%></font></td>
              </tr>
			 
			 
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      do im&oacute;vel do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_cod_imovel" type="text" class="inputBox" id="txt_cod_imovel" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<% if rs9("cod_imovel") = "não informado" or rs9("cod_imovel") = "" then response.write "00" else response.write rs9("cod_imovel") end if%>" size="38" maxlength="20" align="left"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link 
                      de visualiza&ccedil;&atilde;o do im&oacute;vel do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_link" type="text" class="inputBox" id="txt_link" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>; " value="<%=rs9("link_imovel")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			 
			 
			 
			 
			 
			    <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("nome")%>" size="38" maxlength="50" align="left"></td>
              </tr>
              <tr>
                      <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                          do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
				<div align="left">
				<input name="txt_telefone" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>; " value="<%=rs9("telefone")%>" size="38" maxlength="20" align="left">
	            </div>
	            </td>
              </tr>
              <tr>
                      <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">email 
                          do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
				<div align="left">
				<input name="txt_email" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_email" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>;" value="<%=rs9("email")%>" size="38" maxlength="50" align="left">
	            </div>
	           </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="HEIGHT: 18px; WIDTH: 290px ; background: <%=medio%>;" value="<%=rs9("endereco_vend")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			 
              
			  
			  
             
			  
			  
			  
                
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> <a href="javascript:newWindow3('form_incluir_cidade.asp')"></a><font color="#FFFFFF"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <select name="combo1" class="inputBox" id="combo1" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" onChange="javascript:atualizacarros(this.form);">
                     <option value="<% if rs9("cidade_vend") = "não informado" or rs9("cidade_vend") = "qualquer um" or   rs666.eof  then response.write "cqualquer" else response.write rs666("id_combo1") end if  %>" select><% if rs9("cidade_vend") <> "cqualquer" and rs9("cidade_vend") <> "" then response.write rs9("cidade_vend") else response.write "não informado" end if  %></option>
					  
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
					  <option value="cqualquer">qualquer um</option>
                    </select>
                    </font></font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                          do im&oacute;vel atual</font></div></td>
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo2" onChange="javascript:atualizacarros888(this.form);" class="inputBox" id="combo2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                          <option value="<% if rs9("bairro_vend") = "não informado" or rs9("bairro_vend") = "qualquer um" or rs9("bairro_vend") = "bqualquer" or  rs888.eof  then response.write "bqualquer" else response.write rs8888("id_combo2") end if  %>" select><% if rs9("bairro_vend") <> "bqualquer" and rs9("bairro_vend") <> "" then response.write rs9("bairro_vend") else response.write "não informado" end if  %></option>
                        </select> </td>
              </tr>
			  
			  
			   <tr> 
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                       
                   <option value="<%if rs9("vila_vend") <> "não informado" and  not rs444.eof then response.write rs444("id_combo3") else response.write "vlqualquer" end if%>"  selected><% if rs9("vila_vend") <> "vlqualquer" and rs9("vila_vend") <> "" then response.write rs9("vila_vend") else response.write "não informado" end if  %></option>
					
					    </select> </td>
              </tr>
			  
			  
			  
			  
              <tr>
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                   <option value="<%if rs9("tipo_vend") <> "tqualquer" then response.write rs9("tipo_vend") else response.write "tqualquer" end if%>"><%if rs9("tipo_vend") <> "tqualquer" then response.write rs9("tipo_vend") else response.write "não informado" end if%></option>
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
              </tr>
			  
			  
			  
			  <tr>
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de dormit&oacute;rios im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_quartos_vend" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                   
					<option value="<%=rs9("quartos_vend")%>" selected><% if rs9("quartos_vend") = "0" then response.write "não informado" else response.write rs9("quartos_vend") end if%></option>
					 <option value="não informado" >Não informado</option>
					
					<option value="01" >01</option>                    
					<option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07 </option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                    
                    
				  
				  
				  </select>
                    </font></td>
              </tr>
			  
			  <tr>
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de vagas na garagem do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_vagas_vend" size="1" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                   
					<option value="<%=rs9("vagas_vend")%>" selected><% if rs9("vagas_vend") = "0" then response.write "não informado" else response.write rs9("vagas_vend") end if%></option>
					 <option value="não informado" >Não informado</option>
					
					<option value="01" >01</option>                    
					<option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07 </option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                   
                    
				  
				  
				  </select>
                    </font></td>
              </tr>
			  
			  
			  <tr>
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
                        <input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="<%=FormatNumber(rs9("valor_vend"),2)%>" size="12" maxlength="30">
                    </font></td>
              </tr>
                <tr> 
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descrição do imóvel atual</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=claro%>" >&nbsp;</td>
                    </tr>
                  </table></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    
                    <textarea name="txt_descricao_vend" class="inputBox" id="txt_descricao_vend" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 800)"><%=rs9("descricao_vend")%></textarea>
                    </td>
              </tr>
			  <tr><td height="40"></td></tr>
              <tr>
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      pretendida </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> <font color="#FFFFFF"> 
                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" onChange="javascript:atualizacarros2(this.form);">
                     <option value="<% if rs9("cidade_comp") = "não informado" or rs9("cidade_comp") = "qualquer um" or   rs777.eof  then response.write "cqualquer" else response.write rs777("id_combo1") end if  %>" select><% if rs9("cidade_comp") <> "cqualquer" and rs9("cidade_comp") <> "" then response.write rs9("cidade_comp") else response.write "não informado" end if  %></option>
					 
					  <% if not rs5.eof then %>
                      <% While NOT Rs5.EoF %>
                      <option value="<% = Rs5("id_combo1") %>" <% if rs5("nome_combo1") = rs9("cidade_comp") then%>selected<%else%><%end if%>>
                    <% = Rs5("nome_combo1") %>
                    </option>
                    <% Rs5.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
					<option value="cqualquer">qualquer um</option>
                  </select>
                    </font></font> </td>
              </tr>
                <tr> 
                      <td bgcolor="<%=medio%>" height="120" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      pretendido </font></div></td>
                  <td bgcolor="<%=medio%>" height="120" style="border:1px solid #FFFFFF;"><select name="combo4"  class="inputBox" id="combo4" style="HEIGHT: 120px; WIDTH: 150px; background:<%=medio%>" multiple size="8">
                    
					
					<%				 
	  '-----------------------pegar vários bairros-----------
  
  
  
dim Variavel
dim Retorno
dim i
Variavel = rs9("bairro_comp")
Retorno = Split(Variavel,", ")

i=0

Set rs4 = Server.CreateObject("ADODB.RecordSet")


for i=0 to UBound(Retorno)



strSQL4 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where nome_combo2 like '"& Retorno(i) &"' and cidade_combo2 ='"&rs9("cidade_comp")&"' "

 
 

rs4.open strSQL4,Conexao3,2,1

while not rs4.eof

%>

<option value="<%=rs4("id_combo2")%>" selected><%=rs4("nome_combo2")%></option>


<%
rs4.MoveNext
Wend

rs4.close




%>


<%
next



%>
					
					
					
					
					
                  </select></td>
              </tr>
			  
			  <tr> 
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                          pretendida</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                          <option value="<%if rs9("vila_comp") <> "não informado" and  not rs555.eof then response.write rs555("id_combo3") else response.write "vlqualquer" end if%>"  selected><% if rs9("vila_comp") <> "vlqualquer" and rs9("vila_comp") <> "" then response.write rs9("vila_comp") else response.write "não informado" end if  %></option> 
					
                  </select></td>
              </tr>
			  
			  
			  
              <tr>
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      de im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
                    <select name="txt_tipo2" size="1" id="select" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="<%if rs9("tipo_comp") <> "tqualquer" then response.write rs9("tipo_comp") else response.write "tqualquer" end if%>"><%if rs9("tipo_comp") <> "tqualquer" then response.write rs9("tipo_comp") else response.write "não informado" end if%></option>
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
                    </font> </td>
              </tr>
                
				
				
				<tr>
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                      de dormit&oacute;rios do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_quartos_comp" size="1" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                     
					 <option value="<%=rs9("quartos_comp")%>" selected><% if rs9("quartos_comp") = "0" then response.write "não informado" else response.write rs9("quartos_comp") end if%></option>
					  <option value="não informado">Não informado</option>
                      <option value="01" >01</option>
                      <option value="02">02 </option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07 </option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                     
                    </select>
                    </font> </td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de vagas do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_vagas_comp" size="1" id="txt_vagas_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     
					 <option value="<%=rs9("vagas_comp")%>" selected><% if rs9("vagas_comp") = "0" then response.write "não informado" else response.write rs9("vagas_comp") end if%></option>
					  <option value="não informado">Não informado</option>
                      <option value="01" >01</option>
                      <option value="02">02 </option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07 </option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                     
                    </select>
                    </font> </td>
              </tr>
                
                
				
				<tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
                        <input name="txt_valor_comp" type="text" class="inputBox" id="txt_valor_comp" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="<%=FormatNumber(rs9("valor_comp"),2)%>" size="12" maxlength="30">
                    </font> </td>
              </tr>
                
				<tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy 
                      do im&oacute;vel</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby" id="txt_standby" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     
					  <option value="<%if rs9("standby") <> "" then response.write rs9("standby") else response.write "excluido" end if %>" selected><%if rs9("standby") <> "" then response.write rs9("standby") else response.write "excluido" end if %></option>
					  <option value="excluido" >Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr> 
				
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                            <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                                do im&oacute;vel pretendido</font></div></td>
                    </tr>
                    <tr> 
                            <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao_comp" class="inputBox" id="txt_descricao_comp" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; " onKeyPress="return limitfield(this, 800)"><%=rs9("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                      <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="145"><input name="image" type="image"  src="bt_atualizar002.jpg" width="145" height="18" border="0"></td>
                            <td width="145"><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
                          </tr>
                        </table> </td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</table>
</form>

<%
           
           %>
 


<br>

<table width="740" height="1000" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td><iframe src="indicacao_permuta02.asp?varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoVagasVend=<%=varIndicacaoVagasVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoVagasComp=<%=varIndicacaoVagasComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" name="meio" width="740px" height="2050px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>


 



<%

'-----------------------------------------------

rs.close

set rs = nothing

'------------------------------------------------


'-----------------------------------------------

rs3.close

set rs3 = nothing

'------------------------------------------------



'-----------------------------------------------



set rs4 = nothing

'------------------------------------------------


'-----------------------------------------------

rs5.close

set rs5 = nothing

'------------------------------------------------



'-----------------------------------------------

rs6.close

set rs6 = nothing

'------------------------------------------------



'-----------------------------------------------

'rs7.close

'set rs7 = nothing

'------------------------------------------------




'-----------------------------------------------

rs9.close

set rs9 = nothing

'------------------------------------------------


'-----------------------------------------------

rs444.close

set rs444 = nothing

'------------------------------------------------



'-----------------------------------------------

rs555.close

set rs555 = nothing

'------------------------------------------------


'-----------------------------------------------

rs666.close

set rs666 = nothing

'------------------------------------------------



'-----------------------------------------------

rs777.close

set rs777 = nothing

'------------------------------------------------



'-----------------------------------------------

rs888.close

set rs888 = nothing

'------------------------------------------------


'-----------------------------------------------

rs999.close

set rs999 = nothing

'------------------------------------------------


'-----------------------------------------------

rs444Tipo22.close

set rs444Tipo22 = nothing

'------------------------------------------------


'-----------------------------------------------

rs444Tipo23.close

set rs444Tipo23 = nothing

'------------------------------------------------

'-----------------------------------------------

rs444VerificaConta.close

set rs444VerificaConta = nothing

'------------------------------------------------




'-----------------------------------------------

rs8888.close

set rs88888 = nothing

'------------------------------------------------


'-----------------------------------------------

rs9999.close

set rs9999 = nothing

'------------------------------------------------









%>

<% response.flush%>
  <%response.clear%>
 
 
 <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript2 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript888 ( Conexao3 ) %>
 

<%

conexao3.close

set conexao3 = nothing

%>

</center>
</body>
</html>
