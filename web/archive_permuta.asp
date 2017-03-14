<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="cores.asp"-->


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
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


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


dim varCidade,stringCidade,varBairro,stringBairro,varNegociacao
dim stringNegociacao,varQuartos,stringQuartos,varCidade2
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

dim stringIndex



 



Set Conexao4 = Server.CreateObject("ADODB.Connection")
Conexao4.open dsn

 Set rrs2 = Server.CreateObject("ADODB.RecordSet")
 SSQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1="&varCidade2
 
 


	rrs2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rrs2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rrs2.ActiveConnection = Conexao3
	
	

	
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 

	 
	
  
 
 
 
	                                      
									
	 



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
 
 
 
 
 
 
 '------------------------------Números de quartos--------------------------------

 
 
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
 
 
 
 
 '------------------------------Números de Vagas--------------------------------

 
 
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
 
 
 
 '-------------------Valor-----------------------------
 
 vValor_vend=request.querystring("txt_valor_vend")
  
  session("vValor_vend") = vValor_vend
  
  session("vValor_vend1")=left(vValor_vend,10)
   session("vValor_vend2")=right(vValor_vend,10)
 
  
  
  vValor_comp=request.querystring("txt_valor_comp")
  
  if vValor_comp = "vqualquer" then
  vValor_comp = "0000000000 0000000000"
  session("vValor_comp") = "0000000000 0000000000" 
  end if
  
  if vValor_vend = "vqualquer" then
  vValor_vend = "0000000000 0000000000"
  session("vValor_vend") = "0000000000 0000000000" 
  end if
  
  
  
  session("vValor_comp") = vValor_comp
  
   session("vValor_comp1")=left(vValor_comp,10)
   session("vValor_comp2")=right(vValor_comp,10)
 
 
 
 
 
 
 dim vValorMedio_vend
 dim vValorMedio_comp
 
 
 if session("vValor_vend") = "vqualquer" or session("vValor_vend") = "" then
 vValorMedio_vend = "0"
 else
 vValorMedio_vend = (int(session("vValor_vend1")) + int(session("vValor_vend2")))/2
 end if
 
 
  
 if session("vValor_comp") = "vqualquer" or session("vValor_comp") = "" then
 vValorMedio_comp = "0"
 else
 vValorMedio_comp = (int(session("vValor_comp1")) + int(session("vValor_comp2")))/2
 end if
 
 session("vValorMedio_comp") = vValorMedio_comp
 
 

 session("vValorMedio") = vValorMedio
 
 
 
 
 
 
 
 '------------------------Sua Cidade--------------------------

stringIndex = " where cod_permuta<>"&"0"&"" 

if  session("vCidade_vend") <> "cqualquer" and session("vCidade_vend") <> "qualquer um" and session("vCidade_vend") <> "não informado" and session("vCidade_vend") <> "" then
stringCidadeVend = " and (cidade_comp='"& session("vCidade_vend")&"' or cidade_comp='"& "não informado" &"' or cidade_comp='"& "cqualquer" &"' or cidade_comp='"&"qualquer um"&"')"
else
stringCidadeVend = ""
end if
 
 
 
 	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend

 if   session("vBairro_vend") <> "bqualquer"  and session("vBairro_vend") <> "" and session("vBairro_vend") <> "não informado" and session("vBairro_vend") <> "qualquer um" then
	stringBairroVend = " and (Bairro_comp like '%"&session("vBairro_vend")&"%' or Bairro_comp like '%"&"não informado"&"%' or Bairro_comp like'%"&"bqualquer"&"%' or Bairro_comp like'%"&"qualquer um"&"%')"
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


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend
 
 
 if session("vQuartos_vend") <> "qqualquer" and session("vQuartos_vend") <> "" then

stringQuartosVend = " and Quartos_comp <="&session("vQuartos_vend")&""
else
stringQuartosVend = ""
 end if
 


 
 
 '-----------------------Número de Vagas do seu imóvel-----------------
 
 
 
 
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
	'stringValorVend = " and Valor_comp >="& session("vValor_vend1") &" and Valor_comp <="& session("vValor_vend2") &""
	stringValorVend = " and Valor_comp >="& session("vValor_vend1") &" and Valor_comp <="& session("vValor_vend2") &""
	
	else	
	stringValorVend = ""
  end if
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp
  if session("vCidade_comp")<>"cqualquer" and session("vCidade_comp")<>"cqualquer" and session("vCidade_comp")<>"não informado" and session("vCidade_comp")<>"qualquer um" and session("vCidade_comp")<>"" then
	stringCidadeComp = " and Cidade_vend ='"& session("vCidade_comp") &"'"
	else
	
	stringCidadeComp = ""
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp

	if session("vBairro_comp") <> "bqualquer" and session("vBairro_comp") <> "bqualquer" and session("vBairro_comp") <> "não informado" and session("vBairro_comp") <> "qualquer um" and session("vBairro_comp") <> "" then
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
	

 



















dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")



'----------------Fazer listagem de cidade atual-------------
dim rs3330
dim Sql3330


Set rs3330 = Server.CreateObject("ADODB.RecordSet")
'Abrindo a tabela MARCAS!
Sql3330 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



 rs3330.CursorLocation = 3
rs3330.CursorType = 3

rs3330.ActiveConnection = Conexao3



        rs3330.Open sql3330, Conexao3





'--------------------------------------------------------------------





'--------------------------Fazer listagem do bairro atual--------
dim rs4440
dim strSQL4440



 Set rs4440 = Server.CreateObject("ADODB.RecordSet")



if session("vCidade_vend2") <> "cqualquer" and session("vCidade_vend2") <> "" then

strSQL4440 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = '"&session("vCidade_vend2")&"'  ORDER BY nome_combo2" 

else

strSQL4440 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 



end if

 rs4440.CursorLocation = 3
rs4440.CursorType = 3

rs4440.ActiveConnection = Conexao3



		
		rs4440.open strSQL4440, Conexao3





'------------------------------------------------------------



'------------------------Listagem da cidade pretendida------------

dim rs5550
dim Sql5550



Set rs5550 = Server.CreateObject("ADODB.RecordSet")

Sql5550 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


 rs5550.CursorLocation = 3
rs5550.CursorType = 3

rs5550.ActiveConnection = Conexao3



        rs5550.Open Sql5550, Conexao3





'-------------------------------------------------------------------


'-----------------------Listagem do bairro pretendido----------------
dim rs6660
dim strSQL6660

Set rs6660 = Server.CreateObject("ADODB.RecordSet")



if session("vCidade_comp2") <> "cqualquer" and session("vCidade_comp2") <> "" then

 strSQL6660 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = '"&session("vCidade_comp2")&"'  ORDER BY nome_combo2" 
	
else

strSQL6660 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 =4  ORDER BY nome_combo2" 
		
end if


 rs6660.CursorLocation = 3
rs6660.CursorType = 3

rs6660.ActiveConnection = Conexao3



        
		
		
		
		
		
		rs6660.Open strSQL6660, Conexao3



'-------------------------------------------------------------------







'------------------------Listagem do tipo atual -------------------

dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 



 rs444Tipo23.CursorLocation = 3
rs444Tipo23.CursorType = 3

rs444Tipo23.ActiveConnection = Conexao3



       
	 
	 
	 
	 
	 
	 
	 rs444Tipo23.Open strSQL444Tipo23, Conexao3






'--------------------------------------------------------------------






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




<html>
<head>
<title>Permuta</title>

<script>

function check(acao){
if(document.Formulario.selTodos.checked){
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked = acao;
}
}
else
{
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked =! acao;
}
}



}





</script>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=800,height=600,resizable=yes,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>



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

function newWindow6666(abrejanela6666) {
   openWindow6666 = window.open(abrejanela6666,'openWin6666','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow6666.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow7777(abrejanela7777) {
   openWindow7777 = window.open(abrejanela7777,'openWin7777','width=600,height=500,resizable=yes,scrollbars=yes')
   openWindow7777.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow5555(abrejanela5555) {
   openWindow5555 = window.open(abrejanela5555,'openWin22','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow5555.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3333(abrejanela3333) {
   openWindow3333 = window.open(abrejanela3333,'openWin22','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow3333.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2222(abrejanela2222) {
   openWindow2222 = window.open(abrejanela2222,'openWin22','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow2222.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3121(abrejanela3121) {
   openWindow3121 = window.open(abrejanela3121,'openWin3121','width=620,height=500,resizable=yes,scrollbars=yes')
   openWindow3121.focus( )
   }

</SCRIPT>

</head>
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="675" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
          <tr> 
            
                <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis.asp" target="_blank">Im&oacute;veis</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores.asp" target="_blank">Compradores</a></strong></font></div></td>
                <td width="135" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta.asp" target="_blank">Permuta</a></strong></font></div></td>
            
                <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta.asp" target="_blank">Proposta</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email.asp" target="_blank">Email</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow7777('procurar_avaliacao_corretor.asp')" style="color:#FFFFFF">Avaliação </a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ligar_urgente_comprador.asp" target="_blank" style="color:#FFFFFF">Ligar 
                urgente</a></strong></font></div></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imovel_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Imóveis 
                clicados</a></strong></font></div></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_contas_procuradas_corretor.asp" target="_blank" style="color:#FFFFFF">Contas 
                acessadas</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_imovel.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                imóvel</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_comprador.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                compradores</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo02.asp" target="_blank" style="color:#FFFFFF">Captação 
                bloco</a></strong></font></div>
				<%else%>
				<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação 
                bloco</strong></font></div>
				
				
				<%end if%></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo01.asp" target="_blank" style="color:#FFFFFF">Atendente 
                bloco</a></strong></font></div>
			<%else%>
			
			<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente 
                bloco</strong></font></div>
			
			<%end if%></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_financiamentos.asp" target="_blank" style="color:#FFFFFF">Financiamentos</a></strong></font></div>
			<%else%>
			<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Financiamentos</strong></font></div>
			<%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp" target="_blank" style="color:#FFFFFF">Cidade</a></strong></font></div>
			  <%else%>
             
			  <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div>
			 
			  <%end if%></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro.asp" target="_blank" style="color:#FFFFFF">Bairro</a></strong></font></div>
			  <%else%>
			  <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div>
			  
              <%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_vila.asp" target="_blank" style="color:#FFFFFF">Vila</a></strong></font></div>
			  <%else%>
			  
			   <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div>
			  
              <%end if%></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_comprador_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Compradores 
                Clicados</a></strong></font></div></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_procurados.asp" target="_blank">Im&oacute;veis 
          procurados</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_referencia_procurados.asp" target="_blank">Refer&ecirc;ncias 
          procuradas</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_procurados.asp" target="_blank">Permutantes 
          procurados</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_origem.asp" target="_blank">Origem</a></strong></font></div>
	  <%else%>
	  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Origem</strong></font></div>
      <%end if%>
	  
	  </td>
            <td width="135" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_tipo.asp" target="_blank">Tipos de imóveis</a></strong></font></div>
			  <%else%>
              
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipos de imóveis</strong></font></div>
		<% end if %>  
		</td> 
		   
		    <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_enviado.asp" target="_blank">Emails 
                enviados </a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_oficial.asp" target="_blank">Proposta oficial 
                 </a></strong></font></div></td>
    
  </tr>
  
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_procurados.asp" target="_blank">Compradores procurados</a></strong></font></div></td>
            
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_imoveis.asp" target="_blank" style="color:#FFFFFF">corretores externos imóveis</a></strong></font></div>
			  <%else%>
              <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>corretores externos imóveis</strong></font></div>
			
			<%end if%> 
            </td>
            
            <td width="135" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_compradores.asp" target="_blank" style="color:#FFFFFF">Corretores externos compradores</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Corretores externos compradores</strong></font></div>
			
			<%end if%> </td> 
		   
		    
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_visualiza_paginas.asp" target="_blank" style="color:#FFFFFF">Visualização de páginas</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Visualização página</strong></font></div>
			
			<%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if  (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_de_fora.asp" target="_blank" style="color:#FFFFFF">Proposta 
                de Fora</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Proposta de fora</strong></font></div>
			
			<%end if%></td>
    
  </tr>
   <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_interno.asp" target="_blank">Email interno</a></strong></font></div></td>
            
            <td width="134" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <div align="center"></div>
			  </td>
            
            <td width="135" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td> 
		   
		    
            <td width="136" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td>
            <td width="134" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td>
    
  </tr>
  
</table></td>
  </tr>
  </table></td>
  </tr></table>

<center>
<br>
<img src="simbol_permuta.jpg"></img>
<br>

<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>A sua permissão é <%=session("permissao")%></strong></font>
<br>
<table width="120" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" ><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="1956C6"></td>
          </tr>
        </table></td>
      <td width="100" height="20" bgcolor="#FFFFFF"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="estatistica_compradores.asp" style="color:black">Standby</a></strong></font></td>
  </tr>
</table>
<br>

<br>
<table width="790" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><% if    session("permissao") = "5" or  session("permissao") = "6"  then %>
        <div align="center"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="estatistica_permuta.asp" style="color:black">Estat&iacute;stica 
          da busca de permutantes</a></strong></font></div>
        <%else%>
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estat&iacute;stica 
          da busca de permutantes</strong></font></div>
        <%end if%></td>
    </tr>
  </table>


  <br>

</center>


<center>
<form name="doublecombo2" method="get" action="archive_permuta.asp">
<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
        <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr>
              <td width="115" height="25" bgcolor="<%=claro%>"> 
                <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Im&oacute;vel 
                  pretendido </strong></font></div></td>
            <td width="115" bgcolor="<%=claro%>"><select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 115px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;" onChange="javascript:atualizacarros2(this.form);">
                                              <option value="cqualquer" selected>Cidade</option>
              <option value="cqualquer" >Qualquer cidade</option>
             
			 
			  <% if not rs5550.eof then %>
              <% While NOT Rs5550.EoF %>
             
			 <option value="<% = Rs5550("id_combo1") %>"<%if session("vCidade_comp2")<> "cqualquer" then%><%if int(rs5550("id_combo1")) = int(session("vCidade_comp2")) then response.write "selected" else response.write "" end if %><%end if%>> 
		   
            <% = Rs5550("nome_combo1") %>
            </option>
			  
			  
			  <% Rs5550.MoveNext %>
              <% Wend %>
              <%else%>
              <option value=""></option>
              <%end if%>
            </select></td>
            <td width="115" bgcolor="<%=claro%>"><select name="combo4" size="1"  class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 115px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                              <option value="bqualquer" selected >Bairro</option>
               <option value="bqualquer">Qualquer bairro</option>
              
			  
			  <% if not rs6660.eof then%>
              <% While NOT Rs6660.EoF %>
             
			   <option value="<% = Rs6660("id_combo2") %>" <% if session("vBairro_comp2") <> "bqualquer" then if Rs6660("id_combo2") = int(session("vBairro_comp2"))  then response.write "selected" else response.write "" end if end if %>> 
                
            
			
			<% = Rs6660("nome_combo2") %>
            </option>
              
			  <% Rs6660.MoveNext %>
              <% Wend %>
              <% else %>
              <option value=""></option>
              <% end if %>
            </select></td>
            <td width="100" bgcolor="<%=claro%>"><select name="txt_tipo_comp" size="1" id="txt_tipo_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 100px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                             
											   <option value="<%=session("vTipo_comp")%>" selected><%if session("vTipo_comp") <> "tqualquer" and session("vTipo_comp") <> "" then  response.write session("vTipo_comp") else response.write "Tipo" end if%></option>
				 
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
            </select></td>
            <td width="70" bgcolor="<%=claro%>"><select name="txt_quartos_comp" size="1" id="txt_quartos_comp"  style="HEIGHT: 18px; WIDTH: 70px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                             <option value="<%=session("vQuartos_comp")%>"><% if session("vQuartos_comp") <> "qqualquer" and session("vQuartos_comp") <> "" then response.write session("vQuartos_comp") else response.write "Quartos" end if%></option>
										 
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
            </select></td>
            <td width="70" bgcolor="<%=claro%>"><select name="txt_vagas_comp" size="1" id="txt_vagas_comp"   style="HEIGHT: 18px; WIDTH: 70px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                              <option value="<%=session("vVagas_comp")%>"><% if session("vVagas_comp") <> "vgqualquer" and session("vVagas_comp") <> "" then response.write session("vVagas_comp") else response.write "Vagas" end if%></option>
									   
			  
			  
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
            </select></td>
            <td width="160" bgcolor="<%=claro%>"><select name="txt_valor_comp" size="1" id="txt_valor_comp" style="HEIGHT: 18px; WIDTH: 160px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                              
								 <option value="<%=session("vValor_comp")%>" selected><% if session("vValor_comp") <> "vqualquer" and session("vValor_comp") <> "" then response.write FormatNumber(session("vValor_comp1"),2)&" até "&FormatNumber(session("vValor_comp2"),2) else response.write "Valor" end if%></option>
			 			  
											  
                  <option value="0000000000 0000020000">Até 20.000,00</option>
                  <option value="0000000000 0000050000">até 50.000,00</option>
                  <option value="0000000000 0000080000">até 80.000,00</option>
                  <option value="0000000000 0000110000">até 110.000,00</option>
                  <option value="0000000000 0000150000">até 150.000,00</option>
                  <option value="0000000000 0000200000">até 200.000,00</option>
                  <option value="0000000000 0000250000">até 250.000,00</option>
                  <option value="0000000000 0000300000">até 300.000,00</option>
                  <option value="0000000000 0000350000">até 350.000,00</option>
                  <option value="0000000000 0000400000">até 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
               
            </select></td>
              <td width="255">&nbsp;</td>
          </tr>
        </table></td>
  </tr>
  <tr><td height="20"></td></tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
          <tr>
              <td width="115" height="25" bgcolor="<%=claro%>"> 
                <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Imóvel 
                  atual </strong></font></div></td>
            <td width="115" bgcolor="<%=claro%>"><select name="combo1"  id="combo1" class="inputBox" style="HEIGHT: 18px; WIDTH: 115px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;" onChange="javascript:atualizacarros(this.form);">
                                              <option value="cqualquer" selected>Cidade</option>
            <option value="cqualquer">Qualquer cidade</option>
           
		   
		    <% if not rs3330.eof then %>
            <% While NOT Rs3330.EoF %>
           <option value="<% = Rs3330("id_combo1") %>"<%if session("vCidade_vend2")<> "cqualquer" then%><%if int(rs3330("id_combo1")) = int(session("vCidade_vend2")) then response.write "selected" else response.write "" end if %><%end if%>> 
		   
            <% = Rs3330("nome_combo1") %>
            </option>
            </option>
            <% Rs3330.MoveNext %>
            <% Wend %>
            <%else%>
            <option value=""></option>
            <%end if%>
          </select></td>
            <td width="115" bgcolor="<%=claro%>"><select name="combo2" id="combo2"  class="inputBox"  style="HEIGHT: 18px; WIDTH: 115px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                             <option value="bqualquer" selected>Bairro</option>
              <option value="bqualquer">Qualquer bairro</option>
             
			 
			  <% if not rs4440.eof then%>
              <% While NOT Rs4440.EoF %>
             
			  <option value="<% = Rs4440("id_combo2") %>" <% if session("vBairro_vend2") <> "bqualquer" then if Rs4440("id_combo2") = int(session("vBairro_vend2"))  then response.write "selected" else response.write "" end if end if %>> 
                
            
			
			<% = Rs4440("nome_combo2") %>
            </option>
             
			  <% Rs4440.MoveNext %>
              <% Wend %>
              <% else %>
              <option value=""></option>
              <% end if %>
            </select></td>
            <td width="100" bgcolor="<%=claro%>"><select name="txt_tipo_vend" size="1" id="txt_tipo_vend" style="HEIGHT: 18px; WIDTH: 100px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                              <option value="<%=session("vTipo_vend")%>" selected><%if session("vTipo_vend") <> "tqualquer" and session("vTipo_vend") <> "" then  response.write session("vTipo_vend") else response.write "Tipo" end if%></option>
				 
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
                </select></td>
            <td width="70" bgcolor="<%=claro%>"><select name="txt_quartos_vend" size="1" id="txt_quartos_vend" style="HEIGHT: 18px; WIDTH: 70px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                           <option value="<%=session("vQuartos_vend")%>"><% if session("vQuartos_vend") <> "qqualquer" and session("vQuartos_vend") <> "" then response.write session("vQuartos_vend") else response.write "Quartos" end if%></option>
										 
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
            </select></td>
            <td width="70" bgcolor="<%=claro%>"><select name="txt_vagas_vend" size="1" id="txt_vagas_vend"  style="HEIGHT: 18px; WIDTH: 70px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                             
											  <option value="<%=session("vVagas_vend")%>"><% if session("vVagas_vend") <> "vgqualquer" and session("vVagas_vend") <> "" then response.write session("vVagas_vend") else response.write "Vagas" end if%></option>
									   
			  
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
            </select></td>
            <td width="160" bgcolor="<%=claro%>"><select name="txt_valor_vend" size="1" id="txt_valor_vend"  style="HEIGHT: 18px; WIDTH: 160px; ; font-size : 11px; background: <%=medio%>; color:#FFFFFF;">
                                             
				  
           
			 
			   <option value="<%=session("vValor_vend")%>" selected><% if session("vValor_vend") <> "vqualquer" and session("vValor_vend") <> "" then response.write FormatNumber(session("vValor_vend1"),2)&" até "&FormatNumber(session("vValor_vend2"),2) else response.write "Valor" end if%></option>
			 
			  <option value="vqualquer">Qualquer um</option>
			   <option value="0000000000 0000020000">Até 20.000,00</option>
                  <option value="0000000000 0000050000"> até 50.000,00</option>
                  <option value="0000000000 0000080000"> até 80.000,00</option>
                  <option value="0000000000 0000110000"> até 110.000,00</option>
                  <option value="0000000000 0000150000"> até 150.000,00</option>
                  <option value="0000000000 0000200000"> até 200.000,00</option>
                  <option value="0000000000 0000250000"> até 250.000,00</option>
                  <option value="0000000000 0000300000"> até 300.000,00</option>
                  <option value="0000000000 0000350000"> até 350.000,00</option>
                  <option value="0000000000 0000400000"> até 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
			  
			   
			    </select></td>
            <td width="255"><table width="255" border="0" cellspacing="0" cellpadding="0">
  <tr>
                    <td width="65" height="25" bgcolor="<%=claro%>"> <div align="center"><input name="submit" type="submit" class="inputSubmit" style="background:<%=escuro%>;" value="Buscar"></div></td>
    <td>&nbsp;</td>
  </tr>
</table></td>
          </tr>
        </table></td>
  </tr>
  
</table>
</form>

</center>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere
dim varCod_imovel

SearchWhere = request.querystring("SearchWhere")
SearchFor = request.querystring("SearchFor")

session("SearchWhere") = SearchWhere
session("SearchFor") = SearchFor


Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = claro




Set rs = Server.CreateObject("ADODB.RecordSet")
 
	
	 
   
   

       
	if request.querystring("combo1")= "" and  request.querystring("SearchWhere")<>"" then	

SQL = "Select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia  from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY cod_permuta DESC"

if session("SearchFor") = "" and session("SearchWhere") = "Data" then
SQL = "Select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais ,permuta.origem_franquia  from permuta where  origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER  BY Cod_permuta DESC"
end if

if session("SearchFor") <>"" and session("SearchWhere") = "Data" then
SQL = "select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "data_atualizacao like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "data_atualizacao like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if

end if



'---------------------------especial proprietário-----------------------------




if session("SearchFor") <>"" and session("SearchWhere") = "proprietario" then


SQL = "select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "nome like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "nome like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if






end if




'-------------------------------------------------------------------------







if session("SearchFor") ="" and session("SearchWhere") = "proprietario" then
SQL = "Select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"'   ORDER  BY Cod_permuta DESC"
end if



if session("SearchFor") <>"" and session("SearchWhere") = "endereco" then





SQL = "select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "endereco_vend like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "endereco_vend like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if






end if


if session("SearchFor") <>"" and session("SearchWhere") = "atendimento" then





SQL = "select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "atendimento like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "atendimento like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if






end if


 if session("SearchFor") ="" and session("SearchWhere") = "endereco" then
SQL = "Select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"'   ORDER BY Cod_permuta   DESC"
end if

if session("SearchFor") ="" and session("SearchWhere") = "telefone" then
SQL = "Select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"'    ORDER BY Cod_permuta   DESC"
end if

if session("SearchFor") <>"" and session("SearchWhere") = "telefone" then


SQL = "select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "telefone like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "telefone like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if



end if

if session("SearchFor") <>"" and session("SearchWhere") = "cod" then
SQL = "Select permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia   from permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and cod_permuta like '"& session("SearchFor") &"'   ORDER BY Cod_permuta   DESC"
end if





else

if session("vCidade_vend") <> "" then
	
	SQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia  FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringTipoComp&stringQuartosComp&stringVagasComp&stringValorComp&" and origem_franquia like '"&session("vOrigem_Franquia")&"' "
	
	else
	
	SQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia  FROM permuta where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"'"
	
	end if



end if
























%>
<form action="archive_permuta.asp?SearchFor=<%=SearchFor%>" Method="GET" name="b2" >

<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td bgcolor="#DAE3F0">
<table border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="<%=claro%>">
          <tr>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>Procurar</b></font></td>
            <td bgcolor="<%=claro%>">
<input type="text" name="SearchFor" class="inputBox" style="background:<%=medio%>;" value="<%=SearchFor%>">
            </td>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>em</b></font></td>
            <td bgcolor="<%=claro%>">
	
	
	<% 
	if SearchWhere = "" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background: <%=medio%>;">
<option value="proprietario" selected >Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone"  >Telefone</option>
<option value="Data" >Data</option>
<option value="cod">Código da permuta</option>
<option value="atendimento">Atendimento</option>
</select>

<%
end if
%>

	
			
			
	
<!-------------------------------------------------- -->

<% 
	if SearchWhere = "endereco" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario"  >Proprietário</option>
<option value="endereco" selected >Endereço</option>
<option value="telefone"  >Telefone</option>
<option value="Data" >Data</option>

<option value="cod">Código da permuta</option></option>
<option value="atendimento">Atendimento</option>
</select>

<%
end if
%>

<!-------------------------------------------------- -->

<% 
	if SearchWhere = "Data" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario">Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="endereco"  >Telefone</option>
<option value="Data" selected>Data</option>

<option value="cod">Código da permuta</option></option>
<option value="atendimento">Atendimento</option>
</select>

<%
end if
%>

<!-- --------------------------------------------------------- -->

<% 
	if SearchWhere = "proprietario" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" selected>Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone"  >Telefone</option>
<option value="Data">Data</option>

<option value="cod">Código da permuta</option></option>
<option value="atendimento">Atendimento</option>
</select>

<%
end if
%>

<% 
	if SearchWhere = "telefone" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone" selected>Telefone</option>
<option value="Data">Data</option>

<option value="cod">Código da permuta</option></option>
<option value="atendimento">Atendimento</option>
</select>

<%
end if
%>

<% 
	if SearchWhere = "cod" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>

<option value="cod" selected>Código da permuta</option></option>
<option value="atendimento">Atendimento</option>
</select>

<%
end if
%>

<% 
	if SearchWhere = "atendimento" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Proprietário</option>
<option value="endereco"  >Endereço</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>

<option value="cod">Código da permuta</option></option>
<option value="atendimento" selected>Atendimento</option>
</select>

<%
end if
%>


















            </td>
            <td bgcolor="<%=claro%>">
<input type="submit" value="Buscar" class="inputSubmit" style="background:<%=escuro%>;"></td>
</tr>
</table>
</td>
</tr>
</table>
</form>
           
<%

Dim LinkTemp
'essa variável vai ser usada como contador


'as variáveis acima são usadas para trocar a cor das tabelas que conterão os valores
'dos recordsets.






dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeira página.

dim intPageCount
'Essa variável vai receber o valor da quantidade de páginas do recordset.

dim intRecordCount
'Essa variável vai receber o número de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a variável intPage recebe o valor "1" na primeira página.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	
RS.Open SQL, Conn, 1, 3
'o recordset é aberto
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%>

<center>
<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Foram encontrados  <%=intRecordCount%> registros na busca.</strong></font>
</center>

           
 <form  Method="Post" name="Formulario" action="multi_excluir_permuta.asp?varCodPermuta=<%=varCodPermuta%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&page=<%=cInt(intPage)%>" >
  <table width="934" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);">
      </td>
      <td width="95" height="18"><% if  session("permissao") = "4" or session("permissao") = "5" or  session("permissao") = "6" then %><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"><%else%><img src="bt_excluir002.jpg" width="95" height="18" border="0"></img><%end if%></td>
      <td width="95" height="18"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></td>
      
	  
	  <td width="23" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;">&nbsp;</td>
	  <td width="43" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cod</strong></font></div></td>
	 
	   <td width="140" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendimento</strong></font></div></td>
   <td width="25" bgcolor="#000000" style="border:1px solid #FFFFFF;"></td>
	  
	  <td width="220" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Permutante</strong></font></div></td>
      <td width="120" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
	  
	  <td width="40" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Indica</strong></font></div></td>
      
	 
   
   
   <td width="170" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;tima 
          atualiza&ccedil;&atilde;o </strong></font></div></td>
  
    </tr>
   
   
   
   
  <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage é menor ou igual a zero então intPage igual a "1"
	'a variável intPage sempre vai ser forçada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados então.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a página exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posição exata do primeiro registro da página correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual ao número de páginas no recordset , estamos na última 
			'página então.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor do número do último recordset.
				'intFinish corresponde ao valor do último registro da página correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'do número de registros na página menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros então
		For intRecord = 1 to RS.PageSize
		'um contador inRecord é colocado até o número de registros na página.
%> 
   
   
    <%










'------------------------------------------------

%>
<%
If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If


if rs("StandBy") = "incluido" then
color1 = "1956C6"
end if


dim varCodCompradores
%>




	<% session("page")=intPage%>
	<% varCodPermuta = rs("COD_permuta") %>
	<tr> 
      <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("COD_permuta")%>"></td>
      <td width="95" height="18" bgcolor="<%=color1%>"><% if  session("permissao") = "4" or session("permissao") = "5" or  session("permissao") = "6"  then %><a href="excluir_permuta.asp?varCodPermuta=<%=varCodPermuta%>&page=<%=cInt(intPage)%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a><%else%><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img><%end if%></td>
      <td width="95" height="18" bgcolor="<%=color1%>"> 
        <% if    session("permissao") <> "4" and session("permissao") <> "5" and session("permissao") <> "6" and session("permissao") <> "2"  then %>
        <% if  UCase(rs("atendimento")) <> UCase(Session("Admin_ID")) then%>
        <img src="bt_atualizar22.jpg" width="95" height="18" border="0"></img> 
        <%else%>
        <a href="javascript:newWindow2('visualizar_permuta33.asp?varCodPermuta=<%=varCodPermuta%>')"><img src="bt_atualizar22.jpg" width="95" height="18" border="0"></img></a> 
        <%end if%>
        <%else%>
        <a href="javascript:newWindow2('visualizar_permuta33.asp?varCodPermuta=<%=varCodPermuta%>')"><img src="bt_atualizar22.jpg" width="95" height="18" border="0"></img></a> 
        <%end if%>
      </td>
     
	  <td width="23" align="center" height="18"  bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
	  
	  </td>
	  <td width="43"  bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cod_permuta")%></font></div></td>
	  
	  <td width="140" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("atendimento")%></font></div></td>
	 
	 <td width="25" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6"  then %><%if  UCase(rs("atendimento")) <> UCase(Session("nome_id")) then %>&nbsp;<%else%><a href="javascript:newWindow2('form_enviar_email33.asp?varCodPermuta=<%=varCodPermuta%>')"><img src="bt_email22.jpg" width="25" height="18" border="0"></a><%end if%><%else%><a href="javascript:newWindow2('form_enviar_email33.asp?varCodPermuta=<%=varCodPermuta%>')"><img src="bt_email22.jpg" width="25" height="18" border="0"></a><%end if%></td>
	 
	  <td width="220" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("nome")%></font> 
        </div></td>
		 <td width="120" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6" then %><% if  UCase(rs("atendimento")) <> UCase(Session("Admin_ID")) then response.write "Não informado" else response.write rs("telefone") end if %><%else%><%response.write rs("telefone") end if %></font></div></td>
      <td width="40" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%

'------------------------Sua Cidade--------------------------

stringIndex = " where cod_permuta<>"&"0"&""
 dim stringCidadeVend002
 
 
  if   rs("cidade_vend") = "não informado" or rs("cidade_vend") = "" or rs("cidade_vend") = "cqualquer" or  rs("cidade_vend") = "qualquer um" then
	stringCidadeVend002 = ""
 else

stringCidadeVend002 = " and (Cidade_comp='"&rs("cidade_vend")&"' or Cidade_comp='"&"não informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend002

 if   rs("bairro_vend") = "não informado" or rs("bairro_vend") = "" or rs("bairro_vend") = "bqualquer" or  rs("bairro_vend") = "qualquer um" then
	stringBairroVend002 = ""
 else
'stringBairroVend = ""
stringBairroVend002 = " and (Bairro_comp like'%"&rs("bairro_vend")&"%' or Bairro_comp like'%"&"não informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend002

'" and (Vila_comp='"&rs("vila_vend")&"' or Vila_comp='"&"não informado"&"' or Vila_comp='"&"vlqualquer"&"' or Vila_comp='"&"qualquer um"&"' )"

 if   rs("vila_vend") = "não informado" or rs("vila_vend") = "" or rs("vila_vend") = "vlqualquer" or rs("vila_vend") = "qualquer um" then
	stringVilaVend002 =  ""
 else

stringVilaVend002 = ""

end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend002
 
 
 if rs("tipo_vend") = "não informado" or rs("tipo_vend") = "" or rs("tipo_vend") = "tqualquer" or rs("tipo_vend") = "qualquer um"  then

stringTipoVend002 = ""

else
stringTipoVend002 = " and Tipo_comp like '%"&rs("tipo_vend")&"%'"
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend002
 
 
 

stringQuartosVend002 = " and Quartos_comp <="&int(rs("quartos_vend"))&""

 


 '-----------------------Número de Vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend002
 
 
 



stringVagasVend002 = " and vagas_comp <="&int(rs("vagas_vend"))&""

 




 
 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
dim PorcentualVend002

dim vValorMenorVend002
dim vValorMaiorVend002

PorcentualVend002 = int(rs("valor_vend"))*20/100

   


   vValorMenorVend002 = int(rs("valor_vend")) - int(PorcentualVend002)
   vValorMaiorVend002 = int(rs("valor_vend")) + int(PorcentualVend002)

 
 
 
 
 
	 dim stringValorVend002
  
	
	
	
	'stringValorVend002 = " and Valor_comp >="&  vValorMenorVend002 &" and Valor_comp <="& vValorMaiorVend002&""
 
     stringValorVend002 = " and Valor_comp >="& int(vValorMenorVend002)&""
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp002
  if rs("cidade_comp")="não informado" or rs("cidade_comp")="" or rs("cidade_comp")="cqualquer" or rs("cidade_comp") = "qualquer um" then
	stringCidadeComp002 = ""
	else
	
	stringCidadeComp002 = " and Cidade_vend ='"& rs("cidade_comp") &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp002

	if rs("bairro_comp") = "não informado" or  rs("bairro_comp") = "" or  rs("bairro_comp") = "bqualquer" or rs("bairro_comp") = "qualquer um" then
	
	
	
	
	
	stringBairroComp002 = ""
	
	
	
	
	else
	
	
	
	'stringBairroComp = " and Bairro_vend ='"& rs("bairro_comp") &"'"
	
	
	
	
 
dim Numero_Indicacoes002
dim Numero_Indicacoes02002




Numero_Indicacoes002 = 0
Numero_Indicacoes02002 = 0


dim soma02002
dim soma002

soma002 = 0
soma02002 = 0

dim Variavel002
dim Retorno002
dim contar002
Variavel002 = rs("bairro_comp")
Retorno002 = Split(rs("bairro_comp"),", ")

contar002=0

dim stringBairro3002
dim stringBairro4002
dim stringBairro5002

for contar002=0 to UBound(Retorno002)

stringBairro3002 = "and ( "
stringBairro4002 = " Bairro_vend='"&Retorno002(contar002)&"'or  " &stringBairro4002

stringBairro5002 = " cod_permuta=0)"


stringBairroComp002 = stringBairro3002&stringBairro4002&stringBairro5002



next


stringBairro3002 = ""
stringBairro4002 = ""
stringBairro5002 = ""

	
	
	

	
	
	end if
	
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 'and Vila_vend ='"& rs("vila_comp") &"'
	 dim stringVilaComp002

	if rs("vila_comp") <> "não informado" and rs("vila_comp") <> "" and rs("vila_comp") <> "vlqualquer" and rs("vila_comp") <> "qualquer um" then
	stringVilaComp002 = ""
	else
	
	stringVilaComp002 = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp
  'if rs("tipo_comp")="não informado" or rs("tipo_comp")="" or rs("tipo_comp")="tqualquer" or rs("tipo_comp") = "qualquer um" then
	'stringTipoComp = ""
	'else
	
	
	'stringTipoComp = " and Tipo_vend ='"& rs("tipo_comp")&"'"
	'end if
	
	
	
	'--------------------------Tipo----------------------------

if rs("tipo_comp") <> "qualquer um" and rs("tipo_comp") <> "não informado" and rs("tipo_comp") <> "" then




 
dim Numero_IndicacoesTipoComp002
dim Numero_Indicacoes02TipoComp002




Numero_IndicacoesTipoComp002 = 0
Numero_Indicacoes02TipoComp002 = 0


dim soma02TipoComp002
dim somaTipoComp002

somaTipoComp002 = 0
soma02TipoComp002 = 0

dim VariavelTipoComp002
dim RetornoTipoComp002
dim contarTipoComp002
VariavelTipoComp002 =  rs("tipo_comp")
RetornoTipoComp002 = Split(rs("tipo_comp"),", ")

contarTipoComp002=0

dim stringTipo3Comp002
dim stringTipo4Comp002
dim stringTipo5Comp002

for contarTipoComp002=0 to UBound(RetornoTipoComp002)

stringTipo3Comp002 = "and ( "
stringTipo4Comp002 = " tipo_vend='"&RetornoTipoComp002(contarTipoComp002)&"'or  " &stringTipo4Comp002

stringTipo5Comp002 = " cod_permuta=0)"


stringTipo2Comp002 = stringTipo3Comp002&stringTipo4Comp002&stringTipo5Comp002







next

stringTipo3Comp002 = ""
stringTipo4Comp002 = ""
stringTipo5Comp002 = ""


else
stringTipo2Comp002 = ""
end if

	
	
	
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp002
  
	
	stringQuartosComp002 = " and Quartos_vend >="& int(rs("quartos_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp002
 
	
	stringVagasComp002 = " and vagas_vend >="& int(rs("vagas_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------

dim PorcentualComp002

dim vValorMenorComp002
dim vValorMaiorComp002

PorcentualComp002 = int(rs("valor_comp"))*20/100

   


   vValorMenorComp002 = int(rs("valor_comp")) - int(PorcentualComp002)
   vValorMaiorComp002 = int(rs("valor_comp")) + int(PorcentualComp002)


	 dim stringValorComp002
  
	
	
	'stringValorComp002 = " and Valor_vend >="& vValorMenorComp002 &" and Valor_vend <="& vValorMaiorComp002 &""
	
	stringValorComp002 = "  and Valor_vend <="& int(vValorMaiorComp002) &""
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	'strSQL444 = "SELECT * FROM permuta"
	'&stringVilaVend
	'&stringVilaComp
	varIndicacaoCodigo002=rs("cod_permuta")
	
	strSQL444 = "SELECT permuta.cod_permuta   FROM permuta"&stringIndex&stringCidadeVend002&stringBairroVend002&stringTipoVend002&stringQuartosVend002&stringVagasVend002&stringValorVend002&stringCidadeComp002&stringBairroComp002&stringTipo2Comp002&stringQuartosComp002&stringVagasComp002&stringValorComp002&" and standby <> 'incluido' and cod_permuta not like "&varIndicacaoCodigo002
	
	
	 varIndicacaoCidadeVend002=rs("cidade_vend")
 varIndicacaoBairroVend002=rs("bairro_vend")
 varIndicacaoVilaVend002=rs("vila_vend")
 varIndicacaoQuartosVend002=rs("quartos_vend")
 varIndicacaoVagasVend002=rs("vagas_vend")
 varIndicacaoValorVend002=rs("valor_vend")
 varIndicacaoTipoVend002=rs("tipo_vend")


 varIndicacaoCidadeComp002=rs("cidade_comp")
 varIndicacaoBairroComp002=rs("bairro_comp")
 varIndicacaoVilaComp002=rs("vila_comp")
 varIndicacaoQuartosComp002=rs("quartos_comp")
 varIndicacaoVagasComp002=rs("vagas_comp")
 varIndicacaoValorComp002=rs("valor_comp")
 varIndicacaoTipoComp002=rs("tipo_comp")
	
	
	
	
	 
rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444.ActiveConnection = Conexao3
	 
	 rs444.Open strSQL444,Conexao3 
	   
     %>
	 <% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('indicacao_permuta22.asp?varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend002%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend002%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend002%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend002%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend002%>&varIndicacaoVagasVend=<%=varIndicacaoVagasVend002%>&varIndicacaoValorVend=<%=varIndicacaoValorVend002%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp002%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp002%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp002%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp002%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp002%>&varIndicacaoVagasComp=<%=varIndicacaoVagasComp002%>&varIndicacaoValorComp=<%=varIndicacaoValorComp002%>&varIndicacaoCodigo=<%=varIndicacaoCodigo002%>')"><%=rs444.RecordCount%></a><strong></font><%else%><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444.RecordCount%></strong></font><%end if%>
	 <%
	 
 do while not rs444.eof 

 
 
 rs444.movenext
loop
 
 rs444.close
  
 
 









%> 


       <%'strSQL444%>  </font></div></td>
     
	<%
	
	%> 
	 
     
	 
	 <td width="170" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("data_atualizacao")%></font></div></td>
    
	
	
	</tr>
	
	 <%
'-----------------------------------------------









rs.movenext
If RS.EOF Then Exit for
Next

%>


	
	
  </table>
</form>





<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000"> 
        <%If cInt(intPage) > 1 Then%>
        <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
        <a href="?page=<%=intPage - 1%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>" style="color:#000000"> 
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF"> <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
       
	   <%dim cont,cont2,i %>
	 
	 
	 <%if int(intPageCount) > 1 then%>
<%
If int(intPage)-5 > 1 then
cont=int(intPage)-5
else
cont=1
end if
%>
<%if cint(cont+10) > cint(intPageCount) then 
cont2=int(intPageCount)
else
cont2=int(cont)+10
end if
%>
<%for i=int(cont) to int(cont2)%>
<%

%>
<a href="?page=<%=i%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>"><%if int(intPage) = int(i) then %><font color="#FF0000"><%else%><font color="#000000"><%end if%><%=i%></font>
</a> 
<%next%>
<%end if%>

	 
	   
	   
	   
	   
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
        <a href="?page=<%=intPage + 1%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Próximo</b></font> 
        </a> 
        <%End If%>
        </font></div></td>
        </tr>
      </table>










 
<%'else%>
<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font> 
<%
'End If%>
<%'else%>
<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font> 
<%else%>
<%end if%>
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
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000199")
group2[2][3]=new Option("200,00 até 500,00","0000000200 0000000500")
group2[2][4]=new Option("500,00 até 1000,00","0000000500 0000001000")
group2[2][5]=new Option("1000,00 até 2000,00","0000001000 0000002000")
group2[2][6]=new Option("Mais de 2000,00","0000002001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Menos de 20.000,00","0000000000 0000019999")
group2[3][3]=new Option("20.000,00 até 50.000,00","0000020000 0000050000")
group2[3][4]=new Option("50.000,00 até 100.000,00","0000050000 0000100000")
group2[3][5]=new Option("100.000,00 até 150.000,00","0000100000 0000150000")
group2[3][6]=new Option("150.000,00 até 200.000,00","0000150000 0000200000")
group2[3][7]=new Option("200.000,00 até 250.000,00","0000200000 0000250000")
group2[3][8]=new Option("250.000,00 até 300.000,00","0000250000 0000300000")
group2[3][9]=new Option("300.000,00 até 350.000,00","0000300000 0000350000")
group2[3][10]=new Option("350.000,00 até 400.000,00","0000350000 0000400000")


group2[3][11]=new Option("Mais de 400.000,00","0000400001 1000000000")









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
            rs.Close
           
          
           Set rs = Nothing
		   


      '------------------------------------------------


         
           
          
           Set rs = Nothing
		   
	'---------------------------------------	   
	
	
	  
	
	
	  '------------------------------------------------


         
           
          
           Set rs444 = Nothing
		   
	'---------------------------------------	 	 
		   
		   
		   
		   
		   
		   
		   
		   
           %>
  <% response.flush%>
  <%response.clear%>
 
  <%else%>
  <table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      <td width="95" height="18"><% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %><a href="javascript:newWindow2('form_permuta_incluir22.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a><%else%><img src="bt_incluir001.jpg" width="95" height="18" border="0"><%end if%></td>
      
    </tr>
 </table>
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Permutante </font><font color="<%=escuro%>"> 
  n&atilde;o encontrado</font></div>
</font> 
 
  <%end if%>
  
  <%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo2) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo2.combo1.options[doublecombo2.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
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
Response.Write "doublecombo2.combo2.options[" & i  & "] = new Option('" & "Bairro" & "','" & "bqualquer" & "');"
i = 1 
While NOT rsCarros3.EoF

Response.Write "doublecombo2.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo2.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
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
Function EscreveFuncaoJavaScript222 ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo2) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo2.combo3.options[doublecombo2.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
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
Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & "Bairro" & "','" & "bqualquer" & "');" & vbcrlf
i = 1
While NOT rsCarros444.EoF


Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & rsCarros444("nome_combo2") & "','" & rsCarros444("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros444.MoveNext
Wend
Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" & vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas444.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas444.close

set rsMarcas444 = nothing

rsCarros444.close

set rsCarros444 = nothing


End Function
%>





  
  <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
  
  <%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>
  <%
  
 
  
 conexao3.close
 
 set conexao3 = nothing
  
  
  %>
  
  
</body>
</html>

