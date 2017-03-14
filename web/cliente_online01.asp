<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="style6_imoveis.asp"-->

<%response.Buffer = true %>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Cliente online</title>
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



'---------------------excluir clientes pela hora-----------

dim Sql01
dim rs01

'Abrindo a tabela MARCAS!
Sql01 = "SELECT * FROM cliente_online where data_full <> '"&hora&"/"&dia&"/"&mes&"/"&ano&"' ORDER BY cod_online ASC" 

Set rs01 = Server.CreateObject("ADODB.RecordSet")

	rs01.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs01.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs01.ActiveConnection = Conexao
	
	
	rs01.Open sql01, Conexao



if not rs01.eof then 
                  While NOT Rs01.EoF
                  
				 Conexao.execute"Delete from cliente_online where cod_online="&rs01("cod_online")

                  
                   Rs01.MoveNext 
                   Wend 
				   
				   else
				   
	end if






'---------------------------------------------------------------







'Abrindo a tabela MARCAS!

Sql = "SELECT * FROM cliente_online where data_full like '"&hora&"/"&dia&"/"&mes&"/"&ano&"' and atendimento like '"&session("nome_id")&"' ORDER BY cod_online ASC" 

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
    <td width="150"><a href="javascript:newWindow2('visualizar_compradores33.asp?varCodCompradores=<%=rs("cod_cliente")%>')" style="color:#000000"><%=rs("nome")%></a></td>
	
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
