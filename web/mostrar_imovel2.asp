




<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<% response.buffer=True%>
<%


if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if

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
	 
	Conexao.execute"Insert into imovel_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao,atendimento,origem,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& rs("cod_imovel") &"','"& EnderecoIP &"','"& now() &"','"& rs("tipo") &"','"& rs("quartos") &"','"& rs("vagas") &"','"& rs("cidade") &"','"& rs("bairro") &"','"& rs("valor") &"','"& rs("negociacao") &"','"& vAtendimento02 &"','"& session("origem") &"','"& session("vOrigem_Franquia") &"')"
	
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
 
 
	
	
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem,area_total,area_construida,condominio,condicoes_pagamento,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& rs("cidade") &"','"& rs("bairro") &"','"& rs("tipo") &"','"& rs("quartos") &"','"& VerificaNegociacao &"','"& rs("valor") &"','"& now() &"','"& "não informado" &"','"& vAtendimento &"','"& now() &"','"& rs("vila") &"','"& rs("vagas") &"','"& "não informado" &"','"& "comprador a contatar" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "Busca por referência" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "não informado" &"','"& session("vOrigem_Franquia") &"')"
	
	
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
	
if  not rs444VerificaConta022.eof then

  dim vNumero_acessos
  
  vNumero_acessos = rs444VerificaConta022("acessos")
  
  dim varCodCompradores
  
  varCodCompradores = rs444VerificaConta022("cod_compradores")
  
  if vNumero_acessos = "" then
  vNumero_acessos = "1"
  end if
  
  
  vNumero_acessos = int(vNumero_acessos) + 1



 'Conexao.execute"update compradores set nome='"&"Nico"&"' where cod_compradores="&varCodCompradores
	 



	 Conexao.execute "update compradores set acessos ='"&vNumero_acessos&"',data_ultimo_acesso='"&now()&"'  where cod_compradores="&varCodCompradores
	
	else
	
	 Conexao.execute "update compradores set acessos like'"&vNumero_acessos&"'  where cod_compradores like "&"3385"&""
	
	
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
  <% if PropostaFeita <> "" then %>
  <tr>
      <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Obrigado 
          por fazer uma proposta para esse imóvel, vamos estudá-la e retornaremos 
          em breve.</strong></font></div></td>
  </tr>
  <%end if%>
  
  
  <tr>
 
    <td height="80" bgcolor="#f7ecbf" ><table width="580" border="0" align="center">
          <tr>
            <td width="340"><table width="320" height="60" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Gostou 
                    desse im&oacute;vel? quanto voc&ecirc; pagaria por ele?</strong></font></td>
                </tr>
                <tr>
                  <td><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Qual 
                    a forma de pagamento?</strong></font></td>
                </tr>
              </table></td>
          <td width="130"><table width="130" border="0">
              <tr>
                <td height="18" style="border:1px solid #FFFFFF;" bgcolor="<%=claro%>"><input name="txt_valor" value="0,00" type="text" id="txt_valor" size="38" maxlength="21" align="left" class="inputBox" style="color:#9d9249;HEIGHT: 18px; WIDTH: 130px; background:<%=claro%>"></td>
              </tr>
              <tr>
                <td height="18" style="border:1px solid #FFFFFF;" bgcolor="<%=claro%>"><select name="txt_pagamento" size="1" class="inputBox" id="txt_pagamento" style="color:#9d9249;HEIGHT: 22px; WIDTH: 130px; background:<%=claro%>">
                      <option value="à vista" selected>à vista</option>
					<option value="finaciamento/FGTS">financiamento/FGTS</option>
                    <option value="parcelado">parcelado</option>
					<option value="por permuta">por permuta</option>
					<option value="à combinar">à combinar</option>
                  </select></td>
              </tr>
            </table></td>
          <td width="80"><input name="image" type="image"  src="bt_enviar004.jpg" width="80" height="48" border="0"></td>
        </tr>
      </table></td>
  </tr>
 <tr>
      <td height="150"> 
        <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></div>
        <div align="center"><br>
          <font color="#9d9249" size="5" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow11('simulador02.asp?varCodImovel=<%=varCodImovel%>&varValor=<%=rs("valor")%>')" style="color:#9d9249;text-decoration:none;">Clique 
          aqui e simule o financiamento.</a></strong></font></div>
        <div align="center"><br>
          <%
		  '----------------------------------pegar o telefone do corretor------------------------------------
		
		dim SqlAtendentePessoal
		dim rsAtendentePessoal
		
		SqlAtendentePessoal = "SELECT * From Compradores where (telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"') and atendimento not like 'internet' ORDER BY cod_compradores ASC" 

Set rsAtendentePessoal = Server.CreateObject("ADODB.RecordSet")

	rsAtendentePessoal.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsAtendentePessoal.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsAtendentePessoal.ActiveConnection = Conexao
	
	
	rsAtendentePessoal.Open SqlAtendentePessoal, Conexao

dim vAtendentePessoal

if rsAtendentePessoal.eof then
 vAtendentePessoal = ""
 else
 
 vAtendentePessoal = rsAtendentePessoal("atendimento")
 end if




		
		
		
		  
		  'Criando conexão com o banco de dados! 
dim Sql_Telefone_Corretor
dim rs_Telefone_Corretor

'Abrindo a tabela MARCAS!

if   not rsAtendentePessoal.eof then

Sql_Telefone_Corretor = "SELECT * FROM senha where List_Name like '"&vAtendentePessoal&"' ORDER BY ID ASC" 

else
Sql_Telefone_Corretor = "SELECT * FROM senha where List_Name like '"&rs("captacao")&"' ORDER BY ID ASC" 


end if





Set rs_Telefone_Corretor = Server.CreateObject("ADODB.RecordSet")

	rs_Telefone_Corretor.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs_Telefone_Corretor.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs_Telefone_Corretor.ActiveConnection = Conexao
	
	
	rs_Telefone_Corretor.Open Sql_Telefone_Corretor, Conexao






		
		
		%>
          <font color="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>O 
          respons&aacute;vel pelo seu atendimento é o sr(a) <font color="red" size="5" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs_Telefone_Corretor("Admin_id")%> 
          </strong></font><font color="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>, 
          Telefone <%=rs_Telefone_Corretor("telefone")%> </font><br>
          <br>
	
	
	<%	  
		  
	rs_Telefone_Corretor.close

set rs_Telefone_Corretor = nothing	  
		  
		  
	rsAtendentePessoal.close

set rsAtendentePessoal = nothing	  
		  
		  
		  '------------------------------------------------------------
		  
		  
		  
		  
		  
		  
		 dim rsComprador01
		 dim strSQLComprador01
		 Set rsComprador01 = Server.CreateObject("ADODB.RecordSet")
    
	strSQLComprador01 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone='"& rs("telefone")&"' order by cod_compradores DESC" 
	 
   
   
RSComprador01.CursorLocation = 3
RSComprador01.CursorType = 3

        rsComprador01.Open strSQLComprador01, Conexao 
		 
		 
		 if not rsComprador01.eof then
		 %>
          <%
		  else
		  
		  
		  end if
		  
		  
		  
		  
		  %>
          <%
		  
		  dim SqlTelefone_acesso
		  dim rsTelefone_acesso
		  
		  SqlTelefone_acesso = "SELECT * FROM telefone_acesso where telefone_acesso like '"&session("telefone")&"' ORDER BY cod_telefone_acesso ASC" 

Set rsTelefone_acesso = Server.CreateObject("ADODB.RecordSet")

	rsTelefone_acesso.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsTelefone_acesso.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsTelefone_acesso.ActiveConnection = Conexao
	
	
	rsTelefone_acesso.Open sqlTelefone_acesso, Conexao
		  
		  
		  
		  
		  if not rsTelefone_acesso.eof  then
		  
		  response.write "O telefone desse imóvel é "&rs("telefone")&" e o endereço "&rs("endereco")
		   end if
		  %>
          </strong></font></div></td>
 </tr>
  
  <tr>
    <td><table bgcolor="<%="#f7ecbf"%>" width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td width="579"><table bgcolor="<%="#f7ecbf"%>" width="579" border="0" cellspacing="0" cellpadding="0" >
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("cidade")%></strong></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("bairro")%></strong></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("vila")%></strong></font></div></td>
                                </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("tipo")%></strong></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
                            Total / Terreno</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("area_total")%> m&sup2;</strong>
                          </font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
                            Construida / &Uacute;til</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("area_construida")%> m&sup2;</strong>
                          </font></div></td>
                                </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("quartos")%></strong></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Banheiros</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("banheiros")%></strong></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("vagas")%></strong></font></div></td>
                                </tr>
                  </table></td>
              </tr>
              <tr>
                  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("negociacao")%></strong></font></div></td>
                                </tr>
                  </table></td>
                  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=FormatNumber(rs("valor"),2)%></strong></font></div></td>
                                </tr>
                  </table></td>
                  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Condom&iacute;nio</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("condominio") <> "" then response.write FormatNumber(rs("condominio"),2) else response.write "0,00" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
              </tr>
			  
			  <tr>
                  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Saldo 
                            devedor </strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("saldo_devedor") <> "" then response.write rs("saldo_devedor") else response.write "não informado" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
                  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Saldo 
                            devedor j&aacute; pago</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("ja_pago_devedor") <> "" then response.write FormatNumber(rs("ja_pago_devedor"),2) else response.write "0,00" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
                  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Saldo 
                            devedor a pagar</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("devendo_devedor") <> "" then response.write FormatNumber(rs("devendo_devedor"),2) else response.write "0,00" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
              </tr>
			  <tr>
			  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor 
                            do IPTU</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("valor_iptu") <> "" then response.write FormatNumber(rs("valor_iptu"),2) else response.write "0,00" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
				  
				  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Rateio</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("rateio") <> "" then response.write FormatNumber(rs("rateio"),2) else response.write "0,00" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
				  
				  <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Outros</strong></font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("valor_outros") <> "" then response.write FormatNumber(rs("valor_outros"),2) else response.write "0,00" end if %></strong></font></div></td>
                                </tr>
                  </table></td>
			  
			  </tr>
			  
			  
			  
            </table></td>
          <td width="10">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
            <td width="580" height="140" bgcolor="<%=claro%>"> 
              <center><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("obs_imovel")%></strong> <br>
              <br>
                <b><strong>Código de referência <%=rs("cod_imovel")%></strong></b></font>
</center></td>
          <td width="5" bgcolor="#f7ecbf">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  
  <%
 
 
 dim strSQL004
	 dim rs004
	 
	 Set rs004 = Server.CreateObject("ADODB.RecordSet")
	 
	 strSQL004 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02  FROM compradores where telefone like '"&rs("telefone")&"' or telefone02 like '"&rs("telefone")&"' or telefone03 like '"&rs("telefone")&"'"
	
	
	
	
	


   
RS004.CursorLocation = 3
RS004.CursorType = 3




        rs004.Open strSQL004, Conexao 
 
 
  if not rs004.eof then
  
  
    %>
 
  <tr><td><table width="590" height="160" border="0" cellpadding="0" cellspacing="0">
          <tr>
          <td width="5">&nbsp;</td>
            <td width="580" height="140" ><table width="599" height="140" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    
                  <td width="599" height="40"><div align="center"><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Aten&ccedil;&atilde;o 
                      este propriet&aacute;rio vende este im&oacute;vel para comprar 
                      outro, existe a possibilidade de <font size="4">permuta</font>, 
                      veja abaixo o que ele quer comprar.</strong></font></div></td>
                  </tr>
                  <tr>
                    
                  <td height="120" bgcolor="<%=claro%>">
<div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><strong><%=rs004("descricao")%></strong></b></font></div></td>
                  </tr>
                </table> 
              </td>
          <td width="366" bgcolor="#f7ecbf">&nbsp;</td>
        </tr>
      </table></td></tr>
  <tr>
  
  <%Else %>
  
  <% end if%>
  
  
  <tr>
  
  <tr><td height="50"><div align="center"><strong><a href="proposta_oficial01.asp?varCod_Imovel=<%=varCodImovel%>" target="_blank" style="color:#9d9249;text-decoration:none;"><font color="#FF0000">Clique 
          aqui e fa&ccedil;a uma proposta oficial por esse im&oacute;vel.</font></a></strong></div></td></tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
              <tr>
                  <td><a href="imprimir_imovel33.asp?varCod_Imovel=<%=varCodImovel%>" target="_blank"><img src="bt_imprimir33.jpg" width="289" height="18" border="0"></a></td>
                <td width="290"><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td width="289"><a href="imovel_proposta.asp?varCodImovel=<%=varCodImovel%>"><img src="bt_agenda03.jpg" width="289" height="18" border="0"></a></td>
                      <td width="1" height="18"></td>
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

<% else %>


<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">   Não foi encontrado o imóvel pedido!!</font>

<% end if %>



<%
'-----------------------------atualização de acesso-----------------
	
	
dim rs444VerificaConta02,strSQL444VerificaConta02
   
    Set rs444VerificaConta02 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta02 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"'" 
	 
	 
	 
	

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
	
	
	
	
	
	
	
	
	'--------------------------------------------------------------------
	


%>



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
