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

'Criando conexão com o banco de dados! 


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






<!--#include file="cores03.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
 
 
 
 
 
dim varIndicacaoCidade
dim varIndicacaoBairro
dim varIndicacaoNegociacao
dim varIndicacaoQuartos
dim varIndicacaoVagas
dim varIndicacaoValor
dim varIndicacaoTipo

 
 
 
 
 
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
 
	
   
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
  
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao3 
		
		
	
	
	Set rs9 = Server.CreateObject("ADODB.RecordSet") 
dim varCodCompradores
	varCodCompradores=request.form("CodComprador")
	
	if varCodCompradores = "" then
	varCodCompradores=request.QueryString("varCodCompradores")
    end if
	
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
	
	
	 
	 if varNome <> "" then
	 strSQL9 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where nome='"&varNome&"' and telefone='"&varTelefone&"' and cod_compradores="&varCodCompradores
	else
	
	strSQL9 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where  cod_compradores="&varCodCompradores
	
	end if
	
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	  
	  
	 rs9.Open strSQL9, Conexao3
	 
	 if not rs9.eof then
	 
	 dim EnderecoIP
	 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
	 
	Conexao3.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs9("nome") &"','"& rs9("telefone") &"','"& rs9("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP &"','"& now() &"')"
	
	 
	 
	 
	 
	 
	 dim vValor
	  vValor=rs9("valor")
   session("vValor")=vValor
   session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)
   
   
   dim rs4,strSQL4
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 ='"&rs9("bairro")&"'  ORDER BY nome_combo2" 
	 
	 
	 rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao3
	 
	 
	 
	 
	 
	 rs4.Open strSQL4, Conexao3		



dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where nome_combo3 ='"&rs9("vila")&"' and cidade_combo3 ='"&rs9("cidade")&"' and bairro_combo3 ='"&rs9("bairro")&"' ORDER BY nome_combo3" 
	
	rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444.ActiveConnection = Conexao3
	
	
	
	
	 rs444.Open strSQL444, Conexao3		




dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 ='"&rs9("bairro")&"' and cidade_combo2 ='"&rs9("cidade")&"' ORDER BY nome_combo2" 
	
	rs555.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs555.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs555.ActiveConnection = Conexao3
	
	

	
	
	
	
	
	 rs555.Open strSQL555, Conexao3		


 varIndicacaoCidade = rs9("cidade")
	 varIndicacaoBairro = rs9("bairro")
	 varIndicacaoNegociacao = rs9("negociacao")
	 varIndicacaoTipo = rs9("tipo")
	 varIndicacaoQuartos = rs9("quartos")
	 varIndicacaoVagas = rs9("vagas")
	 varIndicacaoValor = rs9("Valor")


dim rs666,strSQL666
   
    Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL666 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 where nome_combo1 ='"&rs9("cidade")&"'  ORDER BY nome_combo1" 
	
	
	
	rs666.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs666.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs666.ActiveConnection = Conexao3
	
	
	rs666.Open strSQL666, Conexao3
	
	
	
	
	
	
			







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




	' Conexao3.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta("cod_compradores")
	 


end if





'--------------------------------------------------------------------



if session("nome") = "" then
session("nome") = rs9("nome")
end if



if session("telefone") = "" then
session("telefone") = rs9("telefone")
end if



if session("email") = "" then
session("email") = rs9("email")
end if






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
  
  if vNumero_acessos = "" then
  
  vNumero_acessos = "0"
  
  end if
  
  
  vNumero_acessos = int(vNumero_acessos) + 1


if  not rs444VerificaConta022.eof then




	 Conexao3.execute"update compradores set data_ultimo_acesso='"&now()&"',acessos='"&vNumero_acessos&"'  where cod_compradores="&rs444VerificaConta022("cod_compradores")
	end if 
      
	  
	  session("acessos") = "acessado"
	  
	  

end if


'-----------------------------------fim no cadastro de número de acessos-------------
	

%>		



<script>
function isValidDigitNumber (doublecombo)
{






if (doublecombo.txt_proprietario.value == "") {
        alert("Você precisa indicar o nome do comprador!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do comprador!");
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
alert("O telefone do comprador só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}







	
	if (doublecombo.stage22.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
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


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>



<html>


<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>
<title>Conta de comprador</title>
<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

</head>

<!--#include file="style_imoveis02.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="20" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >

<br>
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><a href="acessoLink02.asp?varTelefone=<%=session("varTelefone")%>"><img  src="bt_voltar002.jpg" border="0" ></img></a></td>
    </tr>
  </table>
</center>

<br>
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Aqui 
          você poderá atualizar os seus dados de comprador de imóveis, modificar 
          seus interesses ou simplesmente verificar se novos imóveis foram indicados 
          pelo nosso sistema de acordo com seu interesse.</strong></font></div></td>
    </tr>
  </table>
</center>

<div align="center"><br>
  <font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
  </strong></font> <br>
</div>
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_compradores02.asp?varCodCompradores=<%=varCodCompradores%>">
<table width="590" border="0" cellspacing="0" cellpadding="0" align="center">
  
  <tr>
      <td height="18">
<div align="center"> 
          <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi atualizado com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
           
		   
		    <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      atendente</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs9("atendimento")%>" size="38" maxlength="33" align="left"></td>
              </tr>
		   
		   
		     <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de inclus&atilde;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs9("data")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de atualiza&ccedil;&atilde;o</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs9("data_atualizacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
		       
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do comprador</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("nome")%>" size="38" maxlength="50" align="left"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      do comprador</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>; " value="não informado" size="38" maxlength="20" align="left"></td>

              </tr>
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                      do comprador</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>;" value="não informado" size="38" maxlength="50" align="left"></td>
              </tr>
             
			  
			 
              
			  
			  
              
			  
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="<% if rs9("cidade") = "não informado" or rs666.eof then response.write "cqualquer" else response.write rs666("id_combo1") end if  %>" select><%=rs9("cidade")%></option>
					 
					 <% if not rs3.eof then %>
				    <% While NOT Rs3.EoF %>
                    <option value="<% = Rs3("id_combo1") %>" >
                    <% = Rs3("nome_combo1") %>
                    </option>
                    <% Rs3.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					
					<%end if%>
					<option value="cqualquer">qualquer um</option>
                  </select>
                    <a href="javascript:newWindow3('form_incluir_cidade.asp')"></img></a></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" height="120" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" height="120" style="border:1px solid #FFFFFF;"> 
                    <select name="combo2" class="inputBox"  style="HEIGHT: 120px; WIDTH: 150px; background:<%=medio%>" multiple size="8">
	<%				 
	  '-----------------------pegar vários bairros-----------
  
  
  
dim Variavel
dim Retorno
dim i
Variavel = rs9("bairro")
Retorno = Split(Variavel,", ")

i=0

Set rs4 = Server.CreateObject("ADODB.RecordSet")


for i=0 to UBound(Retorno)



strSQL4 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where nome_combo2 like '"& Retorno(i) &"' and cidade_combo2 ='"&rs9("cidade")&"' "

 rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao3
 
 
 
 

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

	</select>
                  </td>
              </tr>
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                   <option value="<%if rs9("vila") <> "não informado" and not rs444.eof then response.write rs444("id_combo3") else response.write "vlqualquer" end if%>" selected><%=rs9("vila")%></option>
					 
					
                  </select>
                  </td>
              </tr>
			  
			  
			  
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" multiple size="8" id="txt_tipo" class="inputBox" style="HEIGHT: 130
					px; WIDTH: 150px; background: <%=claro%>">
                     
	<%				 '-----------------------pegar vários tipos-----------
  
  
  
dim VariavelTipo
dim RetornoTipo
dim iTipo
VariavelTipo = rs9("tipo")
RetornoTipo = Split(VariavelTipo,", ")

iTipo=0

Set rs04Tipo = Server.CreateObject("ADODB.RecordSet")


for iTipo=0 to UBound(RetornoTipo)



strSQL04Tipo = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo where tipo like '"& RetornoTipo(iTipo) &"'  ORDER BY tipo ASC"

 
 

rs04Tipo.open strSQL04Tipo,Conexao3,2,1

while not (rs04Tipo.eof)

%>
                <option value="<%=rs04Tipo("tipo")%>" selected><%=rs04Tipo("tipo")%></option>
                <%
rs04Tipo.MoveNext
Wend

rs04Tipo.close




%>
                <%
next



%>
					 
					 
					 
					 
					 
					 
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de quartos do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos" id="select4" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="<%=rs9("quartos")%>" selected><%=rs9("quartos")%></option>
					<option value="não informado" >não informado</option>
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
                  </td>
              </tr>
              
			 
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                   <option value="<%=rs9("vagas")%>" selected><%=rs9("vagas")%></option>
				    <option value="não informado" >não informado</option>
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
                  </td>
              </tr>
              
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                   <option value="<%=rs9("ocupacao")%>" selected><%=rs9("ocupacao")%></option>
				    <option value="não informado" >não informado</option>
                    <option value="ocupado">Ocupado</option>
                    <option value="vago">Vago</option>
				  </select>
                  </td>
              </tr>
              
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      que deseja</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="example2" size="1" class="inputBox" id="select7"  style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px; background: <%=claro%>; color:FFFFFF;">
                      <option value="<%=rs9("negociacao")%>" selected><%=rs9("negociacao")%></option>
					 					  
                      <option value="nqualquer" >Qualquer um</option>
                      <option  value="Aluguel">Aluguel</option>
                      <option value="Compra">Compra</option>
                    </select> </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Faixa 
                      de pre&ccedil;o desejada</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="stage22" type="text" id="txt_valor2" size="12" maxlength="12" value="<%=formatnumber(rs9("valor"),2)%>" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>"> 
                  </td>
              </tr>
             
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            do im&oacute;vel desejado</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; " onKeyPress="return limitfield(this, 800)"><%=rs9("descricao")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><input name="image" type="image" src="bt_atualizar002.jpg" width="145" height="18" border="0"></td>
                        <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
                      </tr>
                    </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
<br>
<center>
  <font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
  abaixo indicações de imóveis para você.</strong></font>
</center>
<br>
<table width="740" height="1000" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td><iframe src="indicacao_compradores02.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>" name="meio" width="740px" height="2050px" frameborder="0" scrolling="no"></iframe></td>
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

<%  Response.Redirect "acesso02.asp?SecondTry=True&WrongPW=True" %> 

<% end if %>



<%
           rs.Close
           'fecha a conexão
          
           Set rs = Nothing
		   
		   
		   
		   '----------------------------------------
		   
		   
		     rs3.Close
           'fecha a conexão
          
           Set rs3 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   
		  '----------------------------------------
		   
		   
		     rs333.Close
           'fecha a conexão
          
           Set rs333 = Nothing
		   
		   
		   '---------------------------------------- 
		   
		   
		   
		   '----------------------------------------
		   
		   
		     rs9.Close
           'fecha a conexão
          
           Set rs9 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   
		   '----------------------------------------
		   
		   
		    
           'fecha a conexão
          
           Set rs4 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   
		   '----------------------------------------
		   
		   
		     rs444.Close
           'fecha a conexão
          
           Set rs444 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   '----------------------------------------
		   
		   
		     rs555.Close
           'fecha a conexão
          
           Set rs555 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   '----------------------------------------
		   
		   
		     rs666.Close
           'fecha a conexão
          
           Set rs666 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   
		   '----------------------------------------
		   
		   
		     rs444Tipo22.Close
           'fecha a conexão
          
           Set rs444Tipo22 = Nothing
		   
		   
		   '----------------------------------------
		   
		   
		   
		   '----------------------------------------
		   
		   
		     rs444VerificaConta.Close
           'fecha a conexão
          
           Set rs444VerificaConta = Nothing
		   
		   
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

