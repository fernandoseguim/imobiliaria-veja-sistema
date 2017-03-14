

<%
Option Explicit
%>
<!--#include file="cores02.asp"-->
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%


if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if

Dim Conexao,strSQL,rs,varCodImovel,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,vHorario,vNegociacao
 
 varCodImovel = request.QueryString("varCodImovel")
   
   Dim vdata2

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
  vNome=request.form("txtNome")
      vTelefone=request.form("txtTelefone")
      vEmail=request.form("txtEmail")
	  
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  vProposta=request.form("txtProposta")                                                           
		vHorario=request.form("txtHorario")
		vNegociacao=request.form("txtNegociacao")													  
	Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis  Where cod_imovel = "&varCodImovel
	
	
	 Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
	vimagem=rs("foto_grande")
	
	
	
	'Conexao.execute"Insert into proposta(Foto_proposta, Nome_proposta, Email_proposta, Telefone_proposta, Proposta_proposta, data_proposta,horario_proposta,interesse_proposta,cod_imovel_proposta) values( '"& vimagem &"','"& vNome &"','"& vEmail &"','"& vTelefone &"','"& vProposta &"','"& now() &"','"& vHorario &"','"& vNegociacao &"','"& varCodImovel &"')" 
	 
	 
	
	  
     
	  
	   dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&vTelefone&"%' or telefone02 like'%"&vTelefone&"%' or telefone03 like'%"&vTelefone&"%'" 
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao
	 
	 
	 
	 
	 
	 
    dim vOrigem
	 
	 vOrigem = "Busca de imóvel"
	 
	 dim vAtendimento
	 
	 if  not rs444VerificaConta2.eof then
	 
	 vAtendimento = rs444VerificaConta2("atendimento")
	 else
	 vAtendimento = "não informado"
	 end if
   
   
   
   
   Conexao.execute"Insert into email(nome, telefone, email ,assunto,mensagem,data,cod_imovel,atendimento,origem,origem_franquia) values( '"& vNome &"','"& vTelefone &"','"& vEmail &"','"& vNegociacao &"','"& vProposta &"','"& now() &"','"& varCodImovel &"','"& vAtendimento &"','"& vOrigem &"','"& session("vOrigem_Franquia") &"')" 
	 
	
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Proposta pelo imóvel incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">

<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
    <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></img></td>
</tr>
<tr>
    <td width="590" height="105" bgcolor="<%=escuro%>" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
          <td width="217" height="156" bgcolor="<%=escuro%>" ></td>    
          <td width="202" height="156" bgcolor="<%=escuro%>" ></img> 
            <div align="center"><font color="#FFFFFF"><strong>Sua proposta foi 
              inclu&iacute;da com sucesso!!</strong></font></div></td>   
          <td width="217" height="156" bgcolor="<%=escuro%>" ></td>
</tr>

</table>



</td>
</tr>
<tr>
    <td width="590" height="117" bgcolor="<%=escuro%>" ></td>
</tr>


<tr>
    <td width="590" height="36" bgcolor="<%=escuro%>" ></img></td>

</tr>


</table>












 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


