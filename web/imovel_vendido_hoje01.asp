<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->

<!--#include file="style6_imoveis.asp"-->

<%response.Buffer = true %>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel vendido hoje</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela2) {
   openWindow2 = window.open(abrejanela2,'openWin','width=800,height=600,resizable=yes,scrollbars=yes')
   openWindow2.focus( )
   }

</SCRIPT>

</head>

<body>
<%

dim Conexao
dim rs
dim Sql


dim hora, dia, mes, ano

hora = hour(now())
dia = day(now())
mes = month(now())
ano = year(now())


'Criando conexão com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn











'Abrindo a tabela MARCAS!

Sql = "SELECT * FROM imoveis where data_atualizacao like '%"&dia&"/"&mes&"/"&ano&"%' and captacao like '"&session("nome_id")&"' and imovel_em_negociacao like '"&"Vendido por outros"&"' ORDER BY cod_imovel ASC" 

Set rs = Server.CreateObject("ADODB.RecordSet")

	rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao
	
	
	rs.Open sql, Conexao
%>

<center><table>
<%


if not rs.eof then 
                  While NOT Rs.EoF%>
                  <tr>
    <td width="30" height="30"><img src="bola_piscando01.gif" width="30" height="30" border="0"></img></td>
    <td width="150"><a href="javascript:newWindow2('visualizar_imovel33.asp?varCod_Imovel=<%=rs("cod_imovel")%>&Atendido=<%="sim"%>')" style="color:#000000"><%=rs("proprietario")%></a></td>
	
	</tr>
				  
                 <% 
                   Rs.MoveNext 
                   Wend 
				   
				   else
				   response.write "não há registros"
	end if



%>

</table></center>
<% response.flush%>
<%response.clear%>


</body>
</html>
