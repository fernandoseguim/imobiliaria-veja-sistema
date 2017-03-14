<% response.Buffer = true %>
<!--#include file="dsn.asp"-->


<!--#include file="cores02.asp"-->
<html>
<head>
<title></title>



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
<body bottommargin="30"  topmargin="0" bgcolor="#FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">

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
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQL = "Select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  from imoveis where (telefone like '"&varTelefone&"' or telefone02 like '"&varTelefone&"' or telefone03 like '"&varTelefone&"') and captacao <>'"&"internet"&"' and captacao <>'"&"não informado"&"' ORDER BY cod_imovel DESC"


'--------------------------acrescentar acesso-----------------------------

dim rs444VerificaConta02,strSQL444VerificaConta02
   
    Set rs444VerificaConta02 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta02 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like '%"&varTelefone&"%' or telefone02 like'%"&varTelefone&"%' or telefone03 like'%"&varTelefone&"%'" 
	
	
	rs444VerificaConta02.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

    rs444VerificaConta02.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

     rs444VerificaConta02.ActiveConnection = Conexao
	
	
	 rs444VerificaConta02.Open strSQL444VerificaConta02, Conexao
	

if  not rs444VerificaConta02.eof then




	 Conexao.execute"update compradores set data_ultimo_acesso='"&now()&"' where cod_compradores="&rs444VerificaConta02("cod_compradores")
	end if 





'----------------------------------------------------------------------------









rs.Open SQL, Conexao

%>

<br>

<div align="left"><br>
  <br><table width="400" height="80" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><font color="<%=escuro%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>As 
        suas contas est&atilde;o listadas abaixo.<br>
          Clique no endere&ccedil;o do im&oacute;vel e veja as indica&ccedil;oes 
          de compradores dispon&iacute;veis.</strong></font></div></td>
  </tr>
</table>
  <br>
  <%
If NOT (rs.BOF AND rs.EOF) Then
%>
</div>
<center>
  <table width="600" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="30" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Proprietário</strong></font></div></td>
      <td height="30" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
      <td height="30" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endereço</strong></font></div></td>
   
   
    </tr>
   <%








Do While not rs.eof

'------------------------------------------------

%>
   
    <tr> 
      <td width="200" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="conta_imovel01.asp?varCod_imovel=<%=rs("cod_imovel")%>&varNome=<%=rs("proprietario")%>&varTelefone=<%=rs("telefone")%>" style="text-decoration: none; color: #FFFFFF;"><%=rs("proprietario")%></a></font></div></td>
      <td width="200" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="conta_imovel01.asp?varCod_imovel=<%=rs("cod_imovel")%>&varNome=<%=rs("proprietario")%>&varTelefone=<%=rs("telefone")%>" style="text-decoration: none; color: #FFFFFF;"><%=rs("telefone")%></a></font></div></td>
	    <td width="200" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="conta_imovel01.asp?varCod_imovel=<%=rs("cod_imovel")%>&varNome=<%=rs("proprietario")%>&varTelefone=<%=rs("telefone")%>" style="text-decoration: none; color: #FFFFFF;"><%=rs("endereco")%></a></font></div></td>
	   
		<%
'-----------------------------------------------






rs.movenext


If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
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
<%  Response.Redirect "acesso03.asp?SecondTry=True&WrongPW=True" %> 
</center>
 <br>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center">Imóvel não encontrado</div>
</font>            
 
           
            <%
End If

%>
        
<%
  rs.Close
           'fecha a conexão
          
           Set rs = Nothing
		   
		   
		   rs444VerificaConta02.close
		   
		   set rs444VerificaConta02 = nothing
		   
		   
           %>
  <% response.flush%>
  <%response.clear%>
  
  <%
  
 conexao.close
  
  set conexao = nothing
  
  %>

</body>
</html>
