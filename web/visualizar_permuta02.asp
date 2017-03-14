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
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas3 = Conexao3.Execute ( SqlMarcas3 )

While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"
Set rsCarros3 = Conexao3.Execute ( SqlCarros3 )

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

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Sql5 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs5 = Conexao3.Execute ( Sql5 )
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 










<%
Function EscreveFuncaoJavaScript2 ( Conexao4 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo3.options[form.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas4 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas4 = Conexao4.Execute ( SqlMarcas4 )

While NOT rsMarcas4.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas4("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros4 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas4("id_combo1")&" order by nome_combo2"
Set rsCarros4 = Conexao4.Execute ( SqlCarros4 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros4.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros4("nome_combo2") & "','" & rsCarros4("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros4.MoveNext
Wend
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas4.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao4 = Server.CreateObject("ADODB.Connection")
Conexao4.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql4 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs5 = Conexao4.Execute ( Sql4 ) 
%> 













<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->


<% response.buffer=True%>



<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")


 dim Conexao9,rs9
 Set Conexao9 = Server.CreateObject("ADODB.Connection")
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	Conexao9.Open dsn
	dim strSQL9
	dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	 strSQL9 = "SELECT * FROM permuta where cod_permuta="&varCodPermuta
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	 rs9.Open strSQL9, Conexao9
	



dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT * FROM combo3 where nome_combo3 ='"&rs9("vila_vend")&"' and bairro_combo3 ='"&rs9("bairro_vend")&"' and cidade_combo3 ='"&rs9("cidade_vend")&"'  ORDER BY nome_combo3" 
	 rs444.Open strSQL444, Conexao9		
	





dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT * FROM combo3 where nome_combo3 ='"&rs9("vila_comp")&"' and bairro_combo3 ='"&rs9("bairro_comp")&"' and cidade_combo3 ='"&rs9("cidade_comp")&"'  ORDER BY nome_combo3" 
	 rs555.Open strSQL555, Conexao9		
	








   
   Set rs = Server.CreateObject("ADODB.RecordSet")
   dim rs4,strSQL4,strSQL6,rs6
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	Set rs6 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where nome_combo2 like '"& rs9("bairro_vend") &"' ORDER BY nome_combo2" 
	strSQL6 = "SELECT * FROM combo2 where nome_combo2 like '"& rs9("bairro_comp") &"' ORDER BY nome_combo2"  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		rs4.Open strSQL4, Conexao
		rs6.Open strSQL6, Conexao
		
	
	 dim Conexao2,rs7
 Set Conexao2 = Server.CreateObject("ADODB.Connection")
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	Conexao2.Open dsn
	dim strSQL7
	
	if rs9("cod_imovel") <> "não informado" then
	
	 strSQL7 = "SELECT * FROM imoveis where cod_imovel="&rs9("cod_imovel")
	 rs7.CursorLocation = 3
      rs7.CursorType = 3
	 rs7.Open strSQL7, Conexao2
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
Function EscreveFuncaoJavaScript888 ( Conexao888 )
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
SqlMarcas888 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas888 = Conexao888.Execute ( SqlMarcas888 )

While NOT rsMarcas888.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas888("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros888 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas888("id_combo2")&""

Set rsCarros888 = Conexao888.Execute ( SqlCarros888 )

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

End Function
%> 


<%

'Criando conexão com o banco de dados! 
Set Conexao888 = Server.CreateObject("ADODB.Connection")
Conexao888.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'

Sql888 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs888 = Conexao888.Execute ( Sql888 ) 




%> 




<%
Function EscreveFuncaoJavaScript999 ( Conexao999 )
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
SqlMarcas999 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas999 = Conexao999.Execute ( SqlMarcas999 )

While NOT rsMarcas999.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""

Set rsCarros999 = Conexao999.Execute ( SqlCarros999 )

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

End Function
%> 


<%

'Criando conexão com o banco de dados! 
Set Conexao999 = Server.CreateObject("ADODB.Connection")
Conexao999.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'

Sql999 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs999 = Conexao999.Execute ( Sql999 ) 


 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")




dim rs666,strSQL666
   
    Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL666 = "SELECT * FROM combo1 where nome_combo1 ='"&rs9("cidade_vend")&"'  ORDER BY nome_combo1" 
	 rs666.Open strSQL666, Conexao		


dim rs777,strSQL777
   
    Set rs777 = Server.CreateObject("ADODB.RecordSet")
	strSQL777 = "SELECT * FROM combo1 where nome_combo1 ='"&rs9("cidade_comp")&"'  ORDER BY nome_combo1" 
	 rs777.Open strSQL777, Conexao		



dim rs8888,strSQL8888
   
    Set rs8888 = Server.CreateObject("ADODB.RecordSet")
	strSQL8888 = "SELECT * FROM combo2 where nome_combo2 ='"&rs9("bairro_vend")&"' and cidade_combo2 ='"&rs9("cidade_vend")&"'  ORDER BY nome_combo2" 
	 rs8888.Open strSQL8888, Conexao		



dim rs9999,strSQL9999
   
    Set rs9999 = Server.CreateObject("ADODB.RecordSet")
	strSQL9999 = "SELECT * FROM combo2 where nome_combo2 ='"&rs9("bairro_comp")&"' and cidade_combo2 ='"&rs9("cidade_comp")&"'  ORDER BY nome_combo2" 
	 rs9999.Open strSQL9999, Conexao




%> 








<html>


<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript2 ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript888 ( Conexao888 ) %>
<%  EscreveFuncaoJavaScript999 ( Conexao999 ) %>
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






</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  
  <tr>
    <td width="590" height="20"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="imprimir_permuta22.asp?varCodPermuta=<%=varCodPermuta%>" style="color:#FFFFFF">Visualizar 
        impressão</a></strong></font></div></td>
  </tr>
  <tr>
    <td width="590" height="18"><table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td width="5">&nbsp;</td>
    <td><div align="center">
              <table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="290"> <%
				   dim varRs444Imovel
	if rs9("cod_imovel") <> "" then
	varRs444Imovel = rs9("cod_imovel")
	else
	varRs444Imovel = "0"
	end if
				   
				   
				   
				   dim rs444Imovel,SQL444Imovel
 Set rs444Imovel = Server.CreateObject("ADODB.RecordSet")
 SQL444Imovel = "SELECT * FROM imoveis where cod_imovel="&varRs444Imovel 
	
	
	
	
	
	rs444Imovel.open SQL444Imovel,Conexao,2,1  
	
			
	if  not rs444Imovel.eof then
				  
				  
				  %>
				  
				  
				  
				  
				  
<div align="center"><a href="visualizar_imovel22.asp?varCod_imovel=<%=rs9("cod_imovel")%>"><img src="bt_foto22imovel.jpg" width="290" height="18" border="0"></a></div>
<%else%>

<%end if%></td>
                  <td width="290"><%
				   dim varRs444Comprador
	if rs9("cod_comprador") <> "" then
	varRs444Comprador = rs9("cod_comprador")
	else
	varRs444Comprador = "0"
	end if
				   
				   
				   
				   dim rs444Comprador,SQL444Comprador
 Set rs444Comprador = Server.CreateObject("ADODB.RecordSet")
 SQL444Comprador = "SELECT * FROM compradores where cod_compradores="&varRs444Comprador 
	
	
	
	
	
	rs444Comprador.open SQL444Comprador,Conexao,2,1  
	
			
	if  not rs444Comprador.eof then
				  
				  
				  %>
                    <div align="center"><a href="visualizar_compradores22.asp?varCodCompradores=<%=rs9("cod_comprador")%>"><img src="bt_foto22Compr.jpg" width="290" height="18" border="0"></a></div>
                    <%else%>
                    <%end if%></td>
                </tr>
              </table>
            </div></td>
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
                        
                      <td width="290" height="170"><% if not rs7.eof then %><% If objFSO.FileExists(Server.MapPath(rs7("foto_grande"))) = True Then%><img src="<%=rs7("foto_grande")%>" name="photoslider" width="290" height="170"></img><% else %><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
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
                      <td height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><% if session("permissao") = "4" then%><input name="txt_atendimento" type="text" class="inputBox" id="txt_atendimento" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left"><%else%><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("atendimento")%></font><input name="txt_atendimento" type="hidden" class="inputBox" id="txt_atendimento" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left"><%end if%></td>
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
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs9("nome")%>" size="38" maxlength="35" align="left"></td>
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
                     <option value="<% if rs9("cidade_vend") = "não informado" or rs9("cidade_vend") = "qualquer um" or   rs666.eof  then response.write "cqualquer" else response.write rs666("id_combo1") end if  %>" select><%=rs9("cidade_vend")%></option>
					  
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
                          <option value="<% if rs9("bairro_vend") = "não informado" or rs9("bairro_vend") = "qualquer um" or  rs888.eof  then response.write "bqualquer" else response.write rs8888("id_combo2") end if  %>" select><%=rs9("bairro_vend")%></option>
                        </select> </td>
              </tr>
			  
			  
			   <tr> 
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                       
                   <option value="<%if rs9("vila_vend") <> "não informado" and  not rs444.eof then response.write rs444("id_combo3") else response.write "vlqualquer" end if%>"  selected><%=rs9("vila_vend")%></option>
					
					    </select> </td>
              </tr>
			  
			  
			  
			  
              <tr>
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                          do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                   <option value="<%=rs9("tipo_vend")%>"><%=rs9("tipo_vend")%></option>
				     <option value="Apartamento">Apartamento </option>
				   <option value="Térrea/Sobrado">Térrea/Sobrado</option>
				   <option value="Chácara">Chácara</option>
                  <option value="Flat">Flat</option>
				  <option value="Fazenda">Rural</option>
                  <option value="Prédio Comercial">Prédio Comercial</option>
                  <option value="Galpões">Galpões</option>
                  <option value="Sala Comercial">Sala Comercial</option>
				  <option value="Salão Comercial">Salão Comercial</option>
                  <option value="Terreno/Área">Terreno/Área</option>
                  <option value="Ponto Comercial">Ponto Comercial</option>
				  <option value="Cobertura">Cobertura</option>
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
                    
                    <textarea name="txt_descricao_vend" class="inputBox" id="txt_descricao_vend" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 200)"><%=rs9("descricao_vend")%></textarea>
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
                     <option value="<% if rs9("cidade_comp") = "não informado" or rs9("cidade_comp") = "qualquer um" or   rs777.eof  then response.write "cqualquer" else response.write rs777("id_combo1") end if  %>" select><%=rs9("cidade_comp")%></option>
					 
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
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      pretendido </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="combo4" onChange="javascript:atualizacarros999(this.form);" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      
					 <option value="<% if rs9("bairro_comp") = "não informado" or rs9("bairro_comp") = "qualquer um" or rs9999.eof  then response.write "bqualquer" else response.write rs9999("id_combo2") end if  %>" select><%=rs9("bairro_comp")%></option>
					
					
                  </select></td>
              </tr>
			  
			  <tr> 
                      <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                          pretendida</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                          <option value="<%if rs9("vila_comp") <> "não informado" and  not rs555.eof then response.write rs555("id_combo3") else response.write "vlqualquer" end if%>"  selected><%=rs9("vila_comp")%></option> 
					
                  </select></td>
              </tr>
			  
			  
			  
              <tr>
                      <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      de im&oacute;vel pretendido</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF">
                    <select name="txt_tipo2" size="1" id="select" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="<%=rs9("Tipo_comp")%>"><%=rs9("Tipo_comp")%></option>
					  <option value="Apartamento">Apartamento </option>
				   <option value="Térrea/Sobrado">Térrea/Sobrado</option>
				   <option value="Chácara">Chácara</option>
                  <option value="Flat">Flat</option>
				  <option value="Fazenda">Rural</option>
                  <option value="Prédio Comercial">Prédio Comercial</option>
                  <option value="Galpões">Galpões</option>
                  <option value="Sala Comercial">Sala Comercial</option>
				  <option value="Salão Comercial">Salão Comercial</option>
                  <option value="Terreno/Área">Terreno/Área</option>
                  <option value="Ponto Comercial">Ponto Comercial</option>
				  <option value="Cobertura">Cobertura</option>
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
                
				
				
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            do im&oacute;vel pretendido</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=claro%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao_comp" class="inputBox" id="txt_descricao_comp" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 200)"><%=rs9("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td></td>
                        <td></td>
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

<%
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>
