<% response.Buffer = true %>
<!--#include file="dsn.asp"-->


<!--#include file="cores02.asp"-->
<html>
<head>
<title>Conta de comprador</title>

<!--#include file="style_imoveis02.asp"-->

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
   openWindow = window.open(abrejanela,'openWin','width=345,height=180,resizable=yes')
   openWindow.focus( )
   }

</SCRIPT>






</head>
<body  topmargin="0" bgcolor="#FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">

<br>
<center>
<table width="710" height="60" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td style="border:1px solid <%=claro%>;"> 
        <table width="700" height="50" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="<%=escuro%>"> 
              <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acesso01.asp" style="color:#FFFFFF;text-decoration:none;">Sua 
                conta de permutante</a></strong></font></div></td>
            <td bgcolor="<%=claro%>"> 
              <div align="center"><font color="<%=escuro%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acesso03.asp" style="color:<%=escuro%>;text-decoration:none;">Sua 
                conta de vendedor de im&oacute;veis</a></strong></font></div></td>
            <td bgcolor="<%=escuro%>"> 
              <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="acesso02.asp" style="color:#FFFFFF;text-decoration:none;">Sua 
                conta de comprador</a></strong></font></div></td>
  </tr>
</table>
      </td>
  </tr>
</table>
</center>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere

dim varNome
dim varTelefone

varNome = request.form("nome")
varTelefone = request.form("telefone")

if varTelefone = "" then
varTelefone = request.querystring("varTelefone")
end if




session("varNome") = varNome
session("varTelefone") = varTelefone

dim varCodOrigem

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = escuro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou from compradores where  (telefone like '"&varTelefone&"' or telefone02 like'"&varTelefone&"' or telefone03 like '"&varTelefone&"') and atendimento <>'"&"internet"&"' and atendimento <>'"&"não informado"&"'  ORDER BY cod_compradores DESC"



'--------------------------acrescentar acesso-----------------------------

dim rs444VerificaConta02,strSQL444VerificaConta02
   
    Set rs444VerificaConta02 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta02 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou FROM compradores where telefone like '%"&varTelefone&"%' or telefone02 like '%"&varTelefone&"%' or telefone03 like '%"&varTelefone&"%'" 
	 rs444VerificaConta02.Open strSQL444VerificaConta02, Conexao
	

if  not rs444VerificaConta02.eof then




	 Conexao.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta02("cod_compradores")
	end if 





'----------------------------------------------------------------------------


rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao



rs.Open SQL, Conexao

%>

<br>
<center>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>As 
  suas contas est&atilde;o listadas abaixo.<br> Clique no nome do comprador e veja as indicaçoes de imóveis disponíveis.</strong></font>
</center>
<div align="left"><br>
  <br>
  <br>
  <%
If NOT (rs.BOF AND rs.EOF) Then
%>
</div>
<center>
  <table width="600" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Comprador</strong></font></div></td>
      <td height="30" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
      <td height="30" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>C&oacute;digo 
          de comprador</strong></font></div></td>
   
   
    </tr>
   <%
dim varEmail

if rs("email") <> "" then
varEmail =  rs("email")
session("email") = varEmail
else
varEmail = "wentznico@terra.com.br"
session("email") = varEmail
end if







Do While not rs.eof

'------------------------------------------------

%>
   
    <tr> 
      <td width="200" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="conta_comprador01.asp?varCodCompradores=<%=rs("cod_compradores")%>&varNome=<%=rs("nome")%>&varTelefone=<%=rs("telefone")%>&varEmail=<%=varEmail%>" style="text-decoration: none; color: #FFFFFF;"><%=rs("nome")%></a></font></div></td>
      <td width="200" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="conta_comprador01.asp?varCodCompradores=<%=rs("cod_compradores")%>&varNome=<%=rs("nome")%>&varTelefone=<%=rs("telefone")%>&varEmail=<%=varEmail%>" style="text-decoration: none; color: #FFFFFF;"><%=rs("telefone")%></a></font></div></td>
	    <td width="200" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="conta_comprador01.asp?varCodCompradores=<%=rs("cod_compradores")%>&varNome=<%=rs("nome")%>&varTelefone=<%=rs("telefone")%>&varEmail=<%=varEmail%>" style="text-decoration: none; color: #FFFFFF;"><%=rs("cod_compradores")%></a></font></div></td>
	   
		<%
'-----------------------------------------------






rs.movenext


If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = escuro
Else
	colorchanger = 1
	color1 = escuro
	color2 = medio
End If




loop
%>
    </tr>
  </table>
</center>






 <%else%>
 
 
 
 
 
 
 <br>
 <center>
<%  

Response.Redirect "acesso02.asp?SecondTry=True&WrongPW=True"

'response.write SQL

 %> 
</center>
 <br>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center">Comprador não encontrado</div>
</font>            
 
           
            <%
End If

%>
        
<%
  rs.Close
           'fecha a conexão

           Set rs = Nothing
		   
		   
		   rs444Verificaconta02.close
		   
		   set rs444Verificaconta02 = nothing
		   
		   conexao.close
		   
		   set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
