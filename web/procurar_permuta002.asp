





<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%response.Buffer = true %>


<%
Function EscreveFuncaoJavaScript111( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros4362 (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas333 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


Set rsMarcas333 = Server.CreateObject("ADODB.RecordSet")


 rsMarcas333.CursorLocation = 3
rsMarcas333.CursorType = 3

rsMarcas333.ActiveConnection = Conexao3



        rsMarcas333.Open SqlMarcas333, Conexao3




While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas333("id_combo1")&" order by nome_combo2"



Set rsCarros333 = Server.CreateObject("ADODB.RecordSet")


 rsCarros333.CursorLocation = 3
rsCarros333.CursorType = 3

rsCarros333.ActiveConnection = Conexao3



        rsCarros333.Open SqlCarros333, Conexao3






'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 

Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros333.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros333("nome_combo2") & "','" & rsCarros333("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros333.MoveNext
Wend

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" & vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas333.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas333.close

set rsMarcas333 = nothing

rsCarros333.close

set rsCarros333 = nothing



End Function
%> 




















<%
'Criando conex�o com o banco de dados! 

Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn






'Abrindo a tabela MARCAS!
Sql333 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 
Sql555 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 




Set rs555 = Server.CreateObject("ADODB.RecordSet")


 rs555.CursorLocation = 3
rs555.CursorType = 3

rs555.ActiveConnection = Conexao3



        rs555.Open Sql555, Conexao3




Set rs333 = Server.CreateObject("ADODB.RecordSet")


 rs333.CursorLocation = 3
rs333.CursorType = 3

rs333.ActiveConnection = Conexao3



        rs333.Open sql333, Conexao3


 Set rsss = Server.CreateObject("ADODB.RecordSet")
   dim rs444,strSQL444,strSQL666,rs666
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	strSQL666 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
    
	
		
		
		
		



 rs444.CursorLocation = 3
rs444.CursorType = 3

rs444.ActiveConnection = Conexao3



		
		rs444.Open strSQL444, Conexao3
		
		
		
		


 rs666.CursorLocation = 3
rs666.CursorType = 3

rs666.ActiveConnection = Conexao3



        
		
		
		
		
		
		rs666.Open strSQL666, Conexao3
		



%> 










<%
Function EscreveFuncaoJavaScript222 ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo2) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo2.combo3.options[doublecombo2.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
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




'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
 i = 0 
Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & "Em qual bairro o sr(a) quer adquirir um im�vel ?" & "','" & "bqualquer" & "');" & vbcrlf
i = 1
While NOT rsCarros444.EoF


Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & rsCarros444("nome_combo2") & "','" & rsCarros444("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros444.MoveNext
Wend
Response.Write "doublecombo2.combo4.options[" & i & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" & vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas444.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas444.close

set rsMarcas444 = nothing

rsCarros444.close

set rsCarros444 = nothing


End Function
%> 




















<%

'Abrindo a tabela MARCAS!
Sql444 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 





Set rs555 = Server.CreateObject("ADODB.RecordSet")


 rs555.CursorLocation = 3
rs555.CursorType = 3

rs555.ActiveConnection = Conexao3



        rs555.Open sql444, Conexao3

%> 

























<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs44,strSQL44,Conexaoo
  
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	strSQL44 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
	


 rs44.CursorLocation = 3
rs44.CursorType = 3

rs44.ActiveConnection = Conexao3



        
	
	
	
	
	
	rs44.Open strSQL44, Conexao3



%>














<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo2) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo2.combo1.options[doublecombo2.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
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



'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo2.combo2.options[" & i  & "] = new Option('" & "Qual o bairro do seu im�vel ?" & "','" & "bqualquer" & "');"
i = 1 
While NOT rsCarros3.EoF

Response.Write "doublecombo2.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo2.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas3.close

set rsMarcas3 = nothing

rsCarros3.close

set rsCarros3 = nothing



End Function
%> 


<%
'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rs3 = Server.CreateObject("ADODB.RecordSet")


 rs3.CursorLocation = 3
rs3.CursorType = 3

rs3.ActiveConnection = Conexao3



        rs3.Open sql3, Conexao3




dim Sql33 ,Rs33

Sql33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 





Set rs33 = Server.CreateObject("ADODB.RecordSet")


 rs33.CursorLocation = 3
rs33.CursorType = 3

rs33.ActiveConnection = Conexao3



        rs33.Open sql33, Conexao3




dim Sql333 ,Rs333

Sql333 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 




Set rs333 = Server.CreateObject("ADODB.RecordSet")


 rs333.CursorLocation = 3
rs333.CursorType = 3

rs333.ActiveConnection = Conexao3



        rs333.Open sql333, Conexao3

%> 


<%

dim rs4,strSQL4,Conexao
  
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
		
		
	


 rs4.CursorLocation = 3
rs4.CursorType = 3

rs4.ActiveConnection = Conexao3



       	
		
		
		
		
		
		
	rs4.Open strSQL4, Conexao3



%>







<%
Function EscreveFuncaoJavaScript888 ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros888 (doublecombo2) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo2.combo2.options[doublecombo2.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 



Set rsMarcas888 = Server.CreateObject("ADODB.RecordSet")


 rsMarcas888.CursorLocation = 3
rsMarcas888.CursorType = 3

rsMarcas888.ActiveConnection = Conexao3



        rsMarcas888.Open sqlMarcas888, Conexao3





While NOT rsMarcas888.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas888("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo2.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
sqlCarros888 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 FROM combo3 where id_combo2 =" & rsMarcas888("id_combo2")&""





Set rsCarros888 = Server.CreateObject("ADODB.RecordSet")


 rsCarros888.CursorLocation = 3
rsCarros888.CursorType = 3

rsCarros888.ActiveConnection = Conexao3



        rsCarros888.Open sqlCarros888, Conexao3





'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo2.combo5.options[" & i  & "] = new Option('" & "Qual a vila do seu im�vel ?" & "','" & "vlqualquer" & "');"
i = 1 
While NOT rsCarros888.EoF

Response.Write "doublecombo2.combo5.options[" & i & "] = new Option('" & rsCarros888("nome_combo3") & "','" & rsCarros888("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros888.MoveNext
Wend


Response.Write "doublecombo2.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas888.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%



Sql888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 





Set rs888 = Server.CreateObject("ADODB.RecordSet")


 rs888.CursorLocation = 3
rs888.CursorType = 3

rs888.ActiveConnection = Conexao3



        rs888.Open sql888, Conexao3



%> 




<%
Function EscreveFuncaoJavaScript999 ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros999 (doublecombo2) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo2.combo4.options[doublecombo2.combo4.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 





Set rsMarcas999 = Server.CreateObject("ADODB.RecordSet")


 rsMarcas999.CursorLocation = 3
rsMarcas999.CursorType = 3

rsMarcas999.ActiveConnection = Conexao3



        rsMarcas999.Open sqlMarcas999, Conexao3





While NOT rsMarcas999.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo2.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""




Set rsCarros999 = Server.CreateObject("ADODB.RecordSet")


 rsCarros999.CursorLocation = 3
rsCarros999.CursorType = 3

rsCarros999.ActiveConnection = Conexao3



        rsCarros999.Open sqlCarros999, Conexao3




'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo2.combo7.options[" & i  & "] = new Option('" & "Em qual vila o sr(a) quer adquirir um im�vel ?" & "','" & "vlqualquer" & "');"
i = 1 
While NOT rsCarros999.EoF

Response.Write "doublecombo2.combo7.options[" & i & "] = new Option('" & rsCarros999("nome_combo3") & "','" & rsCarros999("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros999.MoveNext
Wend


Response.Write "doublecombo2.combo7.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas999.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas999.close

set rsMarcas999 = nothing

rsCarros999.close

set rsCarros999 = nothing


End Function
%> 


<%



Sql999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 





Set rs999 = Server.CreateObject("ADODB.RecordSet")


 rs999.CursorLocation = 3
rs999.CursorType = 3

rs999.ActiveConnection = Conexao3



        rs999.Open sql999, Conexao3



%> 







<%
Function EscreveFuncaoJavaScript777 ( Conexao3)
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros777 (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo2.options[doublecombo.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
sqlMarcas777 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 ORDER BY nome_combo2" 


Set rsMarcas777 = Server.CreateObject("ADODB.RecordSet")


 rsMarcas777.CursorLocation = 3
rsMarcas777.CursorType = 3

rsMarcas777.ActiveConnection = Conexao3



        rsMarcas777.Open sqlMarcas777, Conexao3



While NOT rsMarcas777.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas777("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros777 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where id_combo2 =" & rsMarcas777("id_combo2")&""




Set rsCarros777 = Server.CreateObject("ADODB.RecordSet")


 rsCarros777.CursorLocation = 3
rsCarros777.CursorType = 3

rsCarros777.ActiveConnection = Conexao3



        rsCarros777.Open sqlCarros777, Conexao3





'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"
i = 1 
While NOT rsCarros777.EoF

Response.Write "doublecombo.combo5.options[" & i & "] = new Option('" & rsCarros777("nome_combo3") & "','" & rsCarros777("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros777.MoveNext
Wend


Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas777.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

rsMarcas777.close

set rsMarcas777 = nothing

rsCarros777.close

set rsCarros777 = nothing

End Function
%> 


<%



Sql777 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 





Set rs777 = Server.CreateObject("ADODB.RecordSet")


 rs777.CursorLocation = 3
rs777.CursorType = 3

rs777.ActiveConnection = Conexao3



        rs777.Open sql777, Conexao3





'------------------------------selecionar os tipos de im�vel para o formul�rio-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 



 rs444Tipo22.CursorLocation = 3
rs444Tipo22.CursorType = 3

rs444Tipo22.ActiveConnection = Conexao3



	 
	 
	 
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao3







 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 



 rs444Tipo23.CursorLocation = 3
rs444Tipo23.CursorType = 3

rs444Tipo23.ActiveConnection = Conexao3



       
	 
	 
	 
	 
	 
	 
	 rs444Tipo23.Open strSQL444Tipo23, Conexao3





 dim rs444Tipo24,strSQL444Tipo24
   
    Set rs444Tipo24 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo24 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	



 rs444Tipo24.CursorLocation = 3
rs444Tipo24.CursorType = 3

rs444Tipo24.ActiveConnection = Conexao3



        
	
	
	
	
	
	
	 rs444Tipo24.Open strSQL444Tipo24, Conexao3




'-------------------------------------------------------------------------------------------------








%> 











<html>

<!--#include file="style4_imoveis.asp"-->
<head>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=600,height=510,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>


<script>

// Verifica se somente n�meros foram digitados no campo
function isValidDigitNumber (doublecombo) 



{




{


if (doublecombo.txt_nome.value == "Seu nome:") {
		alert("Por favor,deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}

if (doublecombo.txt_nome.value == "") {
		alert("Por favor,deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}


if (doublecombo.stage22.value == "vqualquer") {
		alert("Por favor, escolha um faixa de valor na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_nome.focus();
		
		return false;
}










if (doublecombo.txt_telefone.value == "Seu telefone:") {
		alert("Por favor, coloque seu telefone , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_telefone.value == "") {
		alert("Por favor, coloque seu telefone , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_telefone.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "Seu email:") {
		alert("Por favor, coloque seu email , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}

if (doublecombo.txt_email.value == "") {
		alert("Por favor, coloque seu email , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.txt_email.focus();
		
		return false;
}





var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas n�meros!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}



if (doublecombo.combo1.value == "cqualquer") {
		alert("Voc� precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}



if (doublecombo.example2.value == "nqualquer") {
		alert("Por favor, escolha um tipo de negocia��o , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo.example2.focus();
		
		return false;
}












}
}


</script>








<script>

// Verifica se somente n�meros foram digitados no campo
function isValidDigitNumber2 (doublecombo2) 



{




{


if (doublecombo2.txt_nome.value == "Seu nome:") {
		alert("Por favor,deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.txt_nome.focus();
		
		return false;
}

if (doublecombo2.txt_nome.value == "") {
		alert("Por favor,deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.txt_nome.focus();
		
		return false;
}









if (doublecombo2.txt_telefone.value == "Seu telefone:") {
		alert("Por favor, coloque seu telefone , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.txt_telefone.focus();
		
		return false;
}

if (doublecombo2.txt_telefone.value == "") {
		alert("Por favor, coloque seu telefone , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.txt_telefone.focus();
		
		return false;
}

if (doublecombo2.txt_email.value == "Seu email:") {
		alert("Por favor, coloque seu email , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.txt_email.focus();
		
		return false;
}

if (doublecombo2.txt_email.value == "") {
		alert("Por favor, coloque seu email , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.txt_email.focus();
		
		return false;
}





if (doublecombo2.txt_valor_vend.value == "vqualquer") {
		alert("Por favor, coloque o valor do seu im�vel.");
		doublecombo2.txt_valor_vend.focus();
		
		return false;
}


if (doublecombo2.txt_valor_comp.value == "vqualquer") {
		alert("Por favor, coloque o valor do im�vel que pretende adquirir.");
		doublecombo2.txt_valor_comp.focus();
		
		return false;
}






var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo2.txt_telefone.value.length; nCount++) 
		{
strTempChar1_4=doublecombo2.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas n�meros!");
doublecombo2.txt_telefone.focus();

return false;
}
}


























}
}


</script>







<script>
function isValidDigitNumber3 (form)
{

if (form.txt_nome.value == "") {
		alert("Por favor , deixe seu nome na busca , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		form.txt_nome.focus();
		
		return false;
}

if (form.txt_telefone.value == "") {
		alert("Por favor , deixe seu telefone na busca, pois assim, voc� ter� um atendimento preferencial e exclusivo.");
		form.txt_telefone.focus();
		
		return false;
}



if (form.txt_email.value == "") {
		alert("Por favor , deixe seu email na busca, pois assim, voc� ter� um atendimento preferencial e exclusivo.");
		form.txt_email.focus();
		
		return false;
}























	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}

//-----------------------------------------------










</script>








<script language="javascript">
function funScroll()
{
window.scrollTo(0,321)

}		
</script>


<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>}
</STYLE>



</head>












<body onLoad="funScroll()" bgcolor="E17508"  topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginheight="0" marginwidth="0">
<table width="755" border="0" cellspacing="0" cellpadding="0" bgcolor="EAA813">
  <tr>
    <td><form name="doublecombo"   method="post" onSubmit="return isValidDigitNumber(this);" action="listar_imoveis.asp">

<table width="755" border="0" cellspacing="0" cellpadding="0" >
  <tr>
      <td width="755" height="78"><table width="755" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="755" height="51"><a href="primeira.asp"><img src="top_page001.jpg" width="755" height="51" border="0"></a></td>
          </tr>
          <tr>
            <td width="755" height="14" bgcolor="#000000"><div align="center"><table width="600" xmlns=""><tr><td style="width:600; color:#000000;)"><marquee width="100%" scrolldelay="10" scrollamount="2">
                      <font face="Verdana" size="1" color="#FFFFFF"><B>Imobili�ria 
                      Veja: Av.Ant�rtico 315 - Jardim do Mar - SBC - CEP 09726-150. 
                      Tel: 4123-72-44. CRECI: 11.676-J. Atuando no mercado imobili�rio do grande ABC desde fevereiro de 1991.</B></font>
</marquee></td></tr></table></div></td>
          </tr>
          <tr>
            <td width="755" height="13"><img src="top_page002.jpg" width="755" height="13"></td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td width="755" height="243"><table width="755" height="243" border="0" cellpadding="0" cellspacing="0">
        <tr>
                  <td width="176" height="243" align="center" background="fundo_primeira.jpg" bgcolor="#000000"> 
                    <div align="center">
                      <table width="149" border="0"  cellspacing="0" cellpadding="0" height="172">
                        <tr>
							
                          <td width="149" height="14">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Busca 
                                  de im&oacute;veis</strong> </font> </div></td>
							</tr>
							<tr>
                                  <td height="11"><input name="txt_nome" onfocus="doublecombo.txt_nome.value=''"  type="text" class="inputBox" id="txt_nome"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;" value="Seu nome:"></td>
                            </tr>
							<tr>
                                  <td><input name="txt_telefone" onfocus="doublecombo.txt_telefone.value=''" type="text" class="inputBox" id="txt_telefone"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;" value="Seu telefone:"></td>
                            </tr>
							<tr>
                                  <td><input name="txt_email" onfocus="doublecombo.txt_email.value=''" type="text" class="inputBox" id="txt_email"  style="HEIGHT: 16px; WIDTH: 149px; ; font-size : 9px; background: FFFFFF; color:000000;" value="Seu email:"></td>
                            </tr>
							
							
							<tr>
                                  <td>
								  <select name="combo1" onChange="javascript:atualizacarros4362(this.form);" class="inputBox" style="HEIGHT: 11px; WIDTH: 149px; background:white;color:black;">
                  <option value="cqualquer" selected>Cidade</option>
				  <% if not rs3.eof then %>
                  <% While NOT Rs3.EoF %>
                  <option value="<% = Rs3("id_combo1") %>" > 
                  <% = Rs3("nome_combo1") %>
                  </option>
                  <% Rs3.MoveNext %>
                  <% Wend %>
				  <option value="cqualquer">qualquer uma</option>
                  <%else%>
                  <option value=""></option>
                  <%end if%>
                </select>
								  
								   </td>
                            </tr>
                            <tr>
                                  <td><select name="combo2" class="inputBox" onChange="javascript:atualizacarros777(this.form);" style="HEIGHT: 11px; WIDTH: 149px; background:white;color:black;">
                   <option value="bqualquer" selected>Bairro/Regi�o</option>
				  <% if not rs4.eof then%>
                  <% While NOT Rs4.EoF %>
                  <option value="<% = Rs4("id_combo2") %>"> 
                  <% = Rs4("nome_combo2") %>
                  </option>
                  <% Rs4.MoveNext %>
				  
                  <% Wend %>
				   <option value="bqualquer">qualquer um</option>
				  
                  <% else %>
                  <option value=""></option>
                  <% end if %>
                </select> </td>
                            </tr>
							
							 <tr>
                                  <td><select name="combo5" class="inputBox" style="HEIGHT: 11px; WIDTH: 149px; background:white;color:black;">
                   <option value="vlqualquer" selected>Vila</option>
				 <option value="vlqualquer">qualquer um</option>
                </select> </td>
                            </tr>
                            <tr>
                                  <td><select name="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="tqualquer">Tipo</option>
				   <option value="tqualquer">Qualquer um</option>
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
                  
                 
                </select></td>
                            </tr>
							<tr>
                                  <td><select name="txt_Quartos" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ;  background:#FFFFFF; color:#000000;">
                  <option value="qqualquer">Quartos</option>
				   <option value="qqualquer">Qualquer um</option>
                  <option value="01">01</option>
				   <option value="02">02</option>
				   <option value="03">03</option>
                  <option value="04">04</option>
				  <option value="05">05</option>
                  <option value="06">06</option>
                   <option value="07">07</option>
				  <option value="08">08</option>
                  <option value="09">09</option>
                  
                 
                </select></td>
                            </tr>
							
							<tr>
                                  <td><select name="txt_garagem" size="1"  class="inputBox" style="HEIGHT: 11px; WIDTH: 149px ;  background:#FFFFFF; color:#000000;">
                  <option value="gqualquer">Vagas na Garagem</option>
				   <option value="gqualquer">Qualquer um</option>
                  <option value="01">01</option>
				   <option value="02">02</option>
				   <option value="03">03</option>
                  <option value="04">04</option>
				  <option value="05">05</option>
                  <option value="06">06</option>
                   <option value="07">07</option>
				  <option value="08">08</option>
                  <option value="09">09</option>
                  
                 
                </select></td>
                            </tr>
							
							
							
                            <tr>
                                  <td><select name="example2" size="1" class="inputBox" id="select7" onChange="redirect2(this.options.selectedIndex)" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="nqualquer">Negocia��o </option>
                  <option value="nqualquer" >Qualquer um </option>
				  <option  value="Aluguel">Aluguel </option>
                  <option value="Venda">Venda </option>
                  
                </select></td>
                            </tr>
                            <tr>
                                  <td><select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 11px; WIDTH: 149px ; background:#FFFFFF; color:#000000;">
                  <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000020000">At� 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 at� 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 at� 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 at� 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 at� 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 at� 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 at� 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 at� 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 at� 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 at� 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
               
			   
			    </select></td>
                            </tr>
                            <tr>
                              <td><input name="image" type="image"  src="bt_procurar002.jpg" width="149" height="15" border="0"></td>
                            </tr>
                            
                          </table></div></td>
            <td width="579" height="243"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="579" height="243">
                <param name="movie" value="front_page.swf">
                <param name="quality" value="high">
                <embed src="front_page.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="579" height="243"></embed></object></td>
        </tr>
      </table></td>
  </tr>
  <tr>
  <td width="755" height="10" bgcolor="863F15"><table width="755" height="10" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="136"> <div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="quem_somos.asp" style="color:#FFCC00">Quem somos</a></strong></font></div></td>
            <td width="116"> <div align="right"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="onde_estamos.asp" style="color:#FFCC00">Onde 
                      estamos </a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="servicos.asp" style="color:#FFCC00">Servi&ccedil;os</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="financiamento.asp" style="color:#FFCC00">Financiamento/FGTS</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="dicas.asp" style="color:#FFCC00">Dicas</a></strong></font></div></td>
            <td width="126"><div align="center"><font color="#FFCC00" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('form_enviar_email.asp')" style="color:#FFCC00">Contato</a></strong></font></div></td>
          </tr>
        </table></td>
  </tr>
  
</table></form>

      <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
       
  
  <tr> 
          <td width="750" height="40"> <div align="center"> 
            <div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Na 
              busca abaixo procure interessados em permutar im&oacute;vel com 
              voc&ecirc;. </strong></font></div></td>
  </tr>
  
  <tr> 
          <td width="750" height="40"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Aten&ccedil;&atilde;o!!</font></strong> 
              <strong>Na busca abaixo deixe seu nome e telefone para podermos 
              lhe atender com mais rapidez e agilidade, os clientes que fornecerem 
              essas informa&ccedil;&otilde;es ser&atilde;o clientes com atendimento 
              preferencial.</strong></font></div></td>
  </tr>
</table>
  
 <form name="doublecombo2"  onSubmit="return isValidDigitNumber2(this);" method="get" action="listar_permuta002.asp">









    
        <table width="351" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="1" height="120" > 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
        
            <td width="608" height="50" > 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="2">QUALIFIQUE 
                AGORA SEUS DADOS E OS DADOS DO SEU IM&Oacute;VEL.  </font></strong></font><font color="#FFFFFF" size="2"> 
                </font> </div></td>
              </tr>
  <tr>
  
  
  <tr>
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
        
            <td width="608" > 
              <div align="center"><font color="#FFFFFF">
          <input name="txt_nome" value="<% if session("nome") = "" then %>Seu nome:<%else%><%response.write session("nome")%><%end if%>"  onfocus="form.txt_nome.value=''" type="text" class="inputBox"  id="txt_nome" style="HEIGHT: 16px; WIDTH: 350px; background:#FFFFFF;" size="12" maxlength="12">
          </font> </div></td>
              </tr>
  <tr>
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
        
            <td width="608" > 
              <div align="center"><font color="#FFFFFF">
          <input name="txt_telefone" type="text" onfocus="form.txt_telefone.value=''" class="inputBox" value="<% if session("telefone") = "" then %>Seu telefone:<%else%><%response.write session("telefone")%><%end if%>" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 350px; background:#FFFFFF;" size="12" maxlength="12">
          </font> </div></td>
              </tr>
 
   <tr>
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                </strong></font></div></td>
                  
        
            <td width="608" > 
              <div align="center"><font color="#FFFFFF">
                <input name="txt_email" type="text" onfocus="form.txt_email.value=''" class="inputBox" value="<% if session("email") = "" then %>Seu email:<%else%><%response.write session("email")%><%end if%>" id="txt_email" style="HEIGHT: 18px; WIDTH: 350px; background:#FFFFFF;" size="12" maxlength="50">
          </font> </div></td>
              </tr>
 
  
 
 
 
 
  <tr>
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
        
            <td width="608" > 
              <div align="center">
          <select name="combo1"  id="combo1" class="inputBox" style="HEIGHT: 16px; WIDTH: 350px; background:white;color:black;" onChange="javascript:atualizacarros(this.form);">
            <option value="cqualquer" selected>Qual a cidade do seu im�vel ?</option>
            <option value="cqualquer">Qualquer cidade</option>
           
		   
		    <% if not rs333.eof then %>
            <% While NOT Rs333.EoF %>
            <option value="<% = Rs333("id_combo1") %>" <% if rs333("nome_combo1") = "Santo Andr�" then%><%else%><%end if%>> 
            <% = Rs333("nome_combo1") %>
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
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" > 
              <div align="center">
            <select name="combo2" id="combo2"  class="inputBox"  style="HEIGHT: 18px; WIDTH: 350px; background:white;color:black;" onChange="javascript:atualizacarros888(this.form);">
              <option value="bqualquer" selected>Qual o bairro do seu im�vel ?</option>
              <option value="bqualquer">Qualquer bairro</option>
             
			 
			  <% if not rs444.eof then%>
              <% While NOT Rs444.EoF %>
              <option value="<% = Rs444("id_combo2") %>"<%if rs444("nome_combo2") = "Bairro Campestre" then%><%end if%>> 
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
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" > 
              <div align="center">
            <select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 350px; background:white;color:black;">
            <option value="vlqualquer" selected>Qual a vila do seu im�vel ?</option>
            <option value="vlqualquer">qualquer um</option>
          </select>
          </div></td>
              </tr>
			  
			  
			  
              <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
                  <option value="tqualquer" selected>Qual o tipo do seu im�vel?</option>
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
            </font></div></td>
              </tr>
               
			   
			    <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_quartos_vend" size="1" id="txt_quartos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
              <option value="qqualquer" selected>Quantos quartos tem o seu im�vel ?</option>
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
            </font></div></td>
              </tr>
			  
			   <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_vagas_vend" size="1" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
              <option value="vgqualquer" selected>Quantas vagas na garagem tem o seu im�vel ?</option>
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
            </font></div></td>
              </tr>
			   
			    <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_valor_vend" size="1" id="txt_valor_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
             <option value="vqualquer" selected>Qual a faixa de valor do seu im�vel ?</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000020000">At� 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 at� 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 at� 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 at� 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 at� 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 at� 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 at� 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 at� 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 at� 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 at� 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
               
			</select>
            </font></div></td>
              </tr>
			   
			   
			   <tr>
                  
            <td width="1" height="120" > 
              <div align="center">
<p></p>
                <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                </strong></font></div></td>
                  
            <td width="608" height="20" > 
              <div align="center"><font color="#FFFFFF"> <font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>QUALIFIQUE 
                AGORA OS DADOS DO IMOVEL QUE O SR(A) QUER ADQUIRIR NA PERMUTA.</strong></font> 
                <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                </font></font> </div></td>
              </tr>
			   
              <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                </strong></font></div></td>
                  
            <td width="608" > 
              <div align="center"><font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 350px; background:white;color:black;" onChange="javascript:atualizacarros2(this.form);">
              <option value="cqualquer" selected>Em qual cidade o sr(a) quer adquirir um im�vel ?</option>
              <option value="cqualquer" >Qualquer cidade</option>
             
			 
			  <% if not rs555.eof then %>
              <% While NOT Rs555.EoF %>
              <option value="<% = Rs555("id_combo1") %>" <% if rs555("nome_combo1") = "Santo Andr�" then%><%else%><%end if%>> 
              <% = Rs555("nome_combo1") %>
              </option>
              <% Rs555.MoveNext %>
              <% Wend %>
              <%else%>
              <option value=""></option>
              <%end if%>
            </select>
            </font></font> </div></td>
              </tr>
                <tr> 
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" > 
              <div align="center">
            <select name="combo4" onChange="javascript:atualizacarros999(this.form);" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 350px; background:white;color:black;">
              <option value="bqualquer" selected>Em qual bairro o sr(a) quer adquirir um im�vel ?</option>
               <option value="bqualquer">Qualquer bairro</option>
              
			  
			  <% if not rs666.eof then%>
              <% While NOT Rs666.EoF %>
              <option value="<% = Rs666("id_combo2") %>"<%if rs666("nome_combo2") = "Bairro Campestre" then%><%end if%>> 
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
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" > 
              <div align="center">
            <select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 350px; background:white;color:black;">
            <option value="vlqualquer" selected>Em qual vila o sr(a) quer adquirir um im�vel ? </option>
            <option value="vlqualquer">qualquer um</option>
          </select>
          </div></td>
              </tr>
			  
			  
			  
              <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_tipo_comp" size="1" id="txt_tipo_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
              <option value="tqualquer" selected>Que tipo de im�vel o sr(a) quer ?</option>
              
			  <option value="tqualquer">Qualquer tipo</option>
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
            </font> </div></td>
              </tr>
              
			   <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_quartos_comp" size="1" id="txt_quartos_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
              <option value="qqualquer" selected>De quantos quartos o sr(a) precisa ?</option>			  
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
            </font></div></td>
              </tr>
			  <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_vagas_comp" size="1" id="txt_vagas_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
              <option value="vgqualquer" selected>De quantas vagas na garagem o sr(a) precisa ?</option>			  
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
            </font></div></td>
              </tr>
			   
			    <tr>
                  
            <td width="1" > 
              <div align="right"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  
            <td width="608" >
<div align="center"><font color="#FFFFFF"> 
            <select name="txt_valor_comp" size="1" id="txt_valor_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 350px; background: white;color:black;">
             <option value="vqualquer" selected>Qual a faixa de valor do im�vel que o sr(a) quer ?</option>
                  <option value="vqualquer">Qualquer um</option>
                  <option value="0000000000 0000020000">At� 20.000,00</option>
                  <option value="0000020001 0000050000">20.001,00 at� 50.000,00</option>
                  <option value="0000050001 0000080000">50.001,00 at� 80.000,00</option>
                  <option value="0000080001 0000110000">80.001,00 at� 110.000,00</option>
                  <option value="0000110001 0000150000">110.001,00 at� 150.000,00</option>
                  <option value="0000150001 0000200000">150.001,00 at� 200.000,00</option>
                  <option value="0000200001 0000250000">200.001,00 at� 250.000,00</option>
                  <option value="0000250001 0000300000">250.001,00 at� 300.000,00</option>
                  <option value="0000300001 0000350000">300.001,00 at� 350.000,00</option>
                  <option value="0000350001 0000400000">350.001,00 at� 400.000,00</option>
                  <option value="0000400001 1000000000">Acima de 400.000,00</option>
               
            </select>
            </font></div></td>
              </tr>
			   
			  
			  
			  
              <tr>
                
            <td width="1">&nbsp;</td>
                  
            <td width="608"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                <input name="image22" type="image"  src="bt_procurar001.jpg" width="149" height="18" border="0">
                </strong></font> </div></td>
              </tr>
            </table>
			</form>
         
       
      </table></td>
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
/* Aqui � criada uma vari�vel "groups" que receber� os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a vari�vel group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero at� o n�mero de elementos do array "groups" */

group2[i2]=new Array()
/* aqui � criado o array "group" que receber� valores conforme o n�mero de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receber� valores de op��es. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("201,00 at� 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 at� 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 at� 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 at� 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 at� 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 at� 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 at� 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 at� 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 at� 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("At�  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 at� 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 at� 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 at� 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 at� 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 at� 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 at� 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 at� 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 at� 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 at� 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("Acima de 400.000,00","0000400001 1000000000")









/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


var temp2=document.doublecombo.stage22
/* aqui a vari�vel "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui � criada a fun��o "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que d� um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que � escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" ser� o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a vari�vel "location" recebe os valores de "stage2" que corresponde ao endere�o de
link para o carregamento de p�gina. */


//-->
</script>
  <%





%>


<%

'------------------------------------


rs888.close

set rs888 = nothing


'------------------------------------------


'------------------------------------


rs333.close

set rs333 = nothing


'------------------------------------------


'------------------------------------


rs33.close

set rs33 = nothing


'------------------------------------------

'------------------------------------


rs44.close

set rs44 = nothing


'------------------------------------------


'------------------------------------


rs999.close

set rs999 = nothing


'------------------------------------------


'------------------------------------


rs444Tipo22.close

set rs444Tipo22 = nothing


'------------------------------------------


'------------------------------------


rs444Tipo23.close

set rs444Tipo23 = nothing


'------------------------------------------

'------------------------------------


rs444.close

set rs444 = nothing


'------------------------------------------


'------------------------------------


rs444Tipo24.close

set rs444Tipo24 = nothing


'------------------------------------------


'------------------------------------


rs555.close

set rs555 = nothing


'------------------------------------------


'------------------------------------


rs666.close

set rs666 = nothing


'------------------------------------------


'------------------------------------


rs777.close

set rs777 = nothing


'------------------------------------------































%>








  <% response.flush%>
  <%response.clear%>
  


<table width="755" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50" bgcolor="<%=escuro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Os 
        dados dispon&iacute;veis neste site s&atilde;o de inteira responsabilidade 
        dos internautas</strong></font></div></td>
  </tr>
</table>

</td>
  </tr>
</table>



<%  EscreveFuncaoJavaScript111 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript777 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript ( Conexao3) %>
<%  EscreveFuncaoJavaScript888 ( Conexao3) %>
<%  EscreveFuncaoJavaScript222 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript999 ( Conexao3 ) %>

<!--#include file="dsn2.asp"-->
</body>
</html>

