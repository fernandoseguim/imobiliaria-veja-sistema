




<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCodImovel,objFSO
dim varCodImovel2


varCodImovel = request.QueryString("varCodimovel")

session("nome")

session("telefone")

session("email")

session("origem")

session("origem") = request.QueryString("origem")



if session("nome") = "" then

session("nome") = request.form("nome")

end if

if session("telefone") = "" then

session("telefone") = request.form("telefone")

end if


if session("email") = "" then

session("email") = request.form("email")

end if







if session("nome") = "" then

session("nome") = request.querystring("nome")

end if

if session("telefone") = "" then

session("telefone") = request.querystring("telefone")

end if


if session("email") = "" then

session("email") = request.querystring("email")

end if


'-----------------------------------------------------------



application("telefone") = session("telefone")



'------------------------------------------------------------------





if session("nome") = "" and session("telefone") = "" and session("email") = "" then

response.Redirect "sem_cadastro.asp?varCodImovel="&varCodImovel&""

end if






   
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.cliques_no_imovel,imoveis.rateio,imoveis.valor_iptu,imoveis.valor_outros FROM imoveis Where cod_imovel = "&varCodImovel
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
		
		
		
	if not(rs.eof) and not(rs.bof) or (rs.recordcount >= 6) then
	
	
		 dim EnderecoIP
	 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
	 
	 dim PropostaFeita
	 
	 PropostaFeita = request.querystring("PropostaFeita")
	 
	 if PropostaFeita = "" then
	 
	 dim strSQL003
	 dim rs003
	 
	 Set rs003 = Server.CreateObject("ADODB.RecordSet")
	 
	 strSQL003 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02  FROM compradores where telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"'"
	
	
	
	
	


   
RS003.CursorLocation = 3
RS003.CursorType = 3

dim vAtendimento02


        rs003.Open strSQL003, Conexao 
	  if not rs003.eof then
	  vAtendimento02 = rs003("atendimento")
	  
	  end if
	 
	 if session("origem") <> "" then
	 session("origem") = "Site"
	 else
	 session("origem") = "Email"
	 end if
	 
	Conexao.execute"Insert into imovel_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao,atendimento,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& rs("cod_imovel") &"','"& EnderecoIP &"','"& now() &"','"& rs("tipo") &"','"& rs("quartos") &"','"& rs("vagas") &"','"& rs("cidade") &"','"& rs("bairro") &"','"& rs("valor") &"','"& rs("negociacao") &"','"& vAtendimento02 &"','"& session("origem") &"')"
	
	end if
		
  
  
  '--------------------------verificar se internauta tem conta-----------
  dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'"&session("telefone")&"' or telefone02 like'"&session("telefone")&"' or telefone03 like'"&session("telefone")&"'" 
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao
	 
    if rs444VerificaConta2.eof then
	
	dim VerificaNegociacao
	
	VerificaNegociacao = rs("negociacao")
	
	if VerificaNegociacao = "venda" then
	VerificaNegociacao = "compra"
	end if
	
	
	
 '----------------------Verificar se está cadastrado em imóveis--------------------
 dim vAtendimento
 vAtendimento = "internet"
 
 
 dim rs444VerificaConta404,strSQL444VerificaConta404
   
    Set rs444VerificaConta404 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta404 = "SELECT imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.captacao,imoveis.cod_imovel FROM imoveis where (telefone like '%"&session("telefone")&"%' or telefone02 like '%"&session("telefone")&"%' or telefone03 like '%"&session("telefone")&"%')" 
	
	
	
	rs444VerificaConta404.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta404.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta404.ActiveConnection = Conexao
	
	
	
	
	
	 rs444VerificaConta404.Open strSQL444VerificaConta404, Conexao
	

if  not rs444VerificaConta404.eof then
 
 
 'vAtendimento = rs444VerificaConta02("captacao")
  
  vAtendimento = rs444VerificaConta404("captacao")
 
 
 end if
 '-----------------------------------------------------------------------------
 
 
	
	
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem,area_total,area_construida,condominio,condicoes_pagamento) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& rs("cidade") &"','"& rs("bairro") &"','"& rs("tipo") &"','"& rs("quartos") &"','"& VerificaNegociacao &"','"& rs("valor") &"','"& now() &"','"& "não informado" &"','"& vAtendimento &"','"& now() &"','"& rs("vila") &"','"& rs("vagas") &"','"& "não informado" &"','"& "comprador a contatar" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "não informado" &"')"
	
	
	end if 
	
	
	
	
	
	
	
	'----------------------------adicionar acesso--------------------------------



if session("telefone") <> "" and session("acessos") = "" then



dim rs444VerificaConta022,strSQL444VerificaConta022
   
    Set rs444VerificaConta022 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta022 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'"&session("telefone")&"' or telefone02 like'"&session("telefone")&"' or telefone03 like'"&session("telefone")&"'" 
	 
	 
	 
	

	rs444VerificaConta022.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta022.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta022.ActiveConnection = Conexao
	
	
	
	 
	 
	 
	 
	 
	 rs444VerificaConta022.Open strSQL444VerificaConta022, Conexao
	


  dim vNumero_acessos
  
  vNumero_acessos = rs444VerificaConta022("acessos")
  
  
  if vNumero_acessos = "" then
  vNumero_acessos = "1"
  end if
  
  
  vNumero_acessos = int(vNumero_acessos) + 1


if  not rs444VerificaConta022.eof then




	 Conexao.execute"update compradores set data_ultimo_acesso='"&now()&"',acessos='"&vNumero_acessos&"'  where cod_compradores="&rs444VerificaConta022("cod_compradores")
	end if 
      
	  
	  session("acessos") = "acessado"
	  
	  

end if


'-----------------------------------fim no cadastro de número de acessos-------------
	
	
	
	'--------------------------Número de cliques no imóvel---------------
	
	
	dim vNumero_cliques
  
  vNumero_cliques = rs("cliques_no_imovel")
  
  
  if vNumero_cliques = "" then
  vNumero_cliques = "0"
  end if
  
  
  vNumero_cliques = int(vNumero_cliques) + 1




    if rs("telefone") <> session("telefone")  and rs("telefone02") <> session("telefone") and rs("telefone03") <> session("telefone") then


	 Conexao.execute"update imoveis set cliques_no_imovel='"&vNumero_cliques&"'  where cod_imovel="&varCodImovel
	
      
	  end if
	  

	  
	  

	
	
	
	
	
	
	
	
	
	'---------------------------------------------------------------------
	
	
	     
 %>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow11(abrejanela11) {
   openWindow11 = window.open(abrejanela11,'openWin11','width=330,height=473,resizable=yes,left=100,scrollbars=yes')
   openWindow11.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow44(abrejanela44) {
   openWindow44 = window.open(abrejanela44,'openWin44','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow44.focus( )
   }

</SCRIPT>


</head>
<!--#include file="style_imoveis.asp"-->
<body bgcolor="#FFFFFF" topmargin="0" bottommargin="50" rightmargin="0" leftmargin="0" marginheight="0" marginwidth="0">


<form name="doublecombo"  method="post" action="incluir_querpagar.asp?varCodImovel=<%=varCodImovel%>">
<table bgcolor="#f7ecbf" width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></td>
  </tr>
  <tr><td height="10"></td></tr>
  <tr>
    <td width="590" height="334"><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td width="580" bgcolor="<%=claro%>" height="334" style="border:1px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                    <div align="center"><img src="<%=rs("foto_grande")%>" name="photoslider" width="580" height="334"></img></div>
                      <% else %>
                      <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div>
                    <% end if %></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="590" height="18"><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td width="580"><table width="580" border="0" cellspacing="0" cellpadding="0">
              <tr>
			  <script language="JavaScript">
                         var photos=new Array()
                         var which=0
                         
photos[0]="<%=rs("foto_grande1")%>"
photos[1]="<%=rs("foto_grande2")%>"
photos[2]="<%=rs("foto_grande3")%>"
photos[3]="<%=rs("foto_grande4")%>"
photos[4]="<%=rs("foto_grande5")%>"
photos[5]="<%=rs("foto_grande6")%>"
photos[6]="<%=rs("foto_grande7")%>"
photos[7]="<%=rs("foto_grande8")%>"
photos[8]="<%=rs("foto_grande9")%>"
photos[9]="<%=rs("foto_grande10")%>"


 var tam = 0;
<% if rs("foto_grande1")<>"imovel00000.jpg"  then%>
                         var tam = 0;
						<%end if%>

<% if rs("foto_grande2")<>"imovel00000.jpg"  then %>
                         var tam = 1;
						<%end if%>
						
<% if rs("foto_grande3")<>"imovel00000.jpg"  then%>
                         var tam = 2;
						<%end if%>					 
                       
 <% if rs("foto_grande4")<>"imovel00000.jpg"  then %>
                         var tam = 3;
						<%end if%>
						
<% if rs("foto_grande5")<>"imovel00000.jpg"  then %>
                      var tam = 4;
						<%end if%>
						
<% if rs("foto_grande6")<>"imovel00000.jpg"  then %>
                      var tam = 5;
						<%end if%>
						
<% if rs("foto_grande7")<>"imovel00000.jpg"  then %>
                      var tam = 6;
						<%end if%>
												
<% if rs("foto_grande8")<>"imovel00000.jpg"  then %>
                      var tam = 7;
						<%end if%>
						
						
<% if rs("foto_grande9")<>"imovel00000.jpg"  then %>
                      var tam = 8;
						<%end if%>
																							
<% if rs("foto_grande10")<>"imovel00000.jpg"  then %>
                      var tam = 9;
						<%end if%>					   
					     function anterior(){
                           if (which>0){
                             which--
                           }else{
                             which=tam;
                           }
                           document.images.photoslider.src=photos[which]
                         }
                         function proxima(){
                           if (which<tam){
                             which++
                           }else{
                             which=0;
                           }
                           document.images.photoslider.src=photos[which]
                         }
                      </script>
                <td width="290">&nbsp;</td>
                <td width="290" height="18"><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:anterior()" class="link" onmouseover="window.status='Anterior'; return true" onmouseout="window.status=''"><img src="bt_anterior002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:proxima()" class="link" onmouseover="window.status='Próxima'; return true" onmouseout="window.status=''"><img src="bt_proxima002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                      </tr>
                    </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
 
<% end if%>


 <%
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   Set objFSO = Nothing
		   set conexao = nothing
		   
		 
		   
		   set rs444Verificaconta2 = nothing
           %>
		  
  <% response.flush%>
  <%response.clear%>
  

</body>
</html>
