<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_conta.asp"-->
<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "S�o Bernardo"
end if

'Criando conex�o com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn

dim objFSO

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")



dim varCod_imovel


varCod_imovel = request.QueryString("varCod_imovel")

dim strSQL

strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento,imoveis.origem_franquia  FROM imoveis where  cod_imovel="&varCod_imovel
	
dim rs

Set rs = Server.CreateObject("ADODB.RecordSet")	
	

rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs.ActiveConnection = Conexao



rs.Open strSQL, Conexao	
	
	




'-----------------------pegar c�digo de permuta--------------------
 
 dim SQL444Permuta202
 
 SQL444Permuta202 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone='"& rs("telefone")&"' order by cod_permuta DESC" 
	

 dim rs444Permuta202

Set rs444Permuta202 = Server.CreateObject("ADODB.RecordSet")

	rs444Permuta202.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs444Permuta202.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs444Permuta202.ActiveConnection = Conexao
	
	
	rs444Permuta202.Open SQL444Permuta202, Conexao

dim vCod444Permuta2022

if not rs444Permuta202.eof then
vCod444Permuta2022 = rs444Permuta202("cod_permuta")
else
vCod444Permuta2022 = "0"

end if


rs444Permuta202.close

set rs444Permuta202 = nothing



'------------------Pegar c�digo de comprador-----------------------------

 
 dim SQL444Comprador202
 
 SQL444Comprador202 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.suites,compradores.piscina,compradores.portaria,compradores.area_total,compradores.quintal,compradores.quadras,compradores.condominio,compradores.edicula,compradores.clique   FROM compradores where telefone='"& rs("telefone")&"' order by cod_compradores DESC" 
	

 dim rs444Comprador202

Set rs444Comprador202 = Server.CreateObject("ADODB.RecordSet")

	rs444Comprador202.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs444Comprador202.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs444Comprador202.ActiveConnection = Conexao
	
	
	rs444Comprador202.Open SQL444Comprador202, Conexao

dim vCod444Comprador202

if not rs444Comprador202.eof then
vCod444Comprador202 = rs444Comprador202("cod_compradores")
else
vCod444Comprador202 = "0"

end if


rs444Comprador202.close

set rs444Comprador202 = nothing








'-----------------------------------------------------------------------	
	
	
	
	
'----------------------------Abrindo listagem de Cidades--------------------

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 


Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs3.ActiveConnection = Conexao
	
	
	rs3.Open Sql3, Conexao






'-----------------------------------------------------------------------------


'------------------Abrindo combo1---------------------------------------------
dim rs666
dim strSQL666


Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL666 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 where nome_combo1 ='"&rs("cidade")&"'  ORDER BY nome_combo1" 
	
	
	
	rs666.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs666.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs666.ActiveConnection = Conexao
	
	
	rs666.Open strSQL666, Conexao
	




'-------------------------------------------------------------------------------



'----------------------Selecionar os tipos de im�veis---------------------------

dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 
	 rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs444Tipo22.ActiveConnection = Conexao
	 
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao








'--------------------------------------------------------------------------------


'-----------------------Acrescentar acessos------------------------------------

'------------------Verifica se o internauta j� tem conta---------------------------
  
  dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&rs("telefone")&"%' or telefone02 like '%"&rs("telefone")&"%' or telefone03 like '%"&rs("telefone")&"%'" 
	
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao



 While NOT rs444VerificaConta2.EoF 
                      
              
		if rs444VerificaConta2("acessos") <> "" then
		
		 
	 Conexao.execute"update compradores set acessos='"&int(rs444VerificaConta2("acessos"))+1&"' where cod_compradores="&rs444VerificaConta2("cod_compradores")
	 
			else
			
			 	 
	 Conexao.execute"update compradores set acessos='"&"1"&"' where cod_compradores="&rs444VerificaConta2("cod_compradores")
	 
		end if
	
			   
                    
                      rs444VerificaConta2.MoveNext 
                      Wend 





'---------------------------------------------------------------------------------




%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow3.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2121(abrejanela2121) {
   openWindow2121 = window.open(abrejanela2121,'openWin2121','width=650,height=600,resizable=yes,scrollbars=yes,Left=0,Top=0')
   openWindow2121.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2323(abrejanela2323) {
   openWindow2323 = window.open(abrejanela2323,'openWin2323','width=400,height=600,resizable=yes,scrollbars=yes,Left=0,Top=0')
   openWindow2323.focus( )
   }

</SCRIPT>



<title>Conta de im�vel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#e6dca9">

<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_conta_imovel01.asp?varCod_imovel=<%=varCod_imovel%>&vPerguntaPermuta=<%=vPerguntaPermuta%>&vPerguntaCompradores=<%=vPerguntaCompradores%>">
<table width="794" height="430" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" height="430" valign="top"><table width="190" height="430" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="190" height="262"  style="border:1px solid #FFFFFF;"><table width="180" height="252" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="180" height="252" bgcolor="#e0a94e"> 
                  <table width="170" height="242" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="170" height="242"><table width="170" height="242" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="170" height="137"><table width="170" height="137" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td width="170" height="102" align="center">
                                      <% if  objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                                      <div align="center"><img src="<%=rs("foto_grande")%>" width="170" height="102"> 
                                        <%else%>
                                        <font color="#FFFFFF"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                                        n�o dispon�vel</font></strong></font></div>
                                      <%end if%></td>
                                </tr>
                                <tr>
                                  <td width="170" height="35" bgcolor="#e6dca9"><div align="center"><font size="2" face="Perpetua Titling MT"><strong>Im&oacute;vel</strong></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                          <tr>
                            <td width="170" height="5"></td>
                          </tr>
                          <tr>
                            <td width="170" height="30" bgcolor="#f1da9f"><input name="txt_nome" class="inputBox" type="text"  id="txt_nome" style="color:#000000;HEIGHT: 20px; WIDTH: 170px; background: #f1da9f ;border-color : #f1da9f;" value="<%=rs("proprietario")%>" size="38" maxlength="33" align="left"></td>
                          </tr>
                          <tr>
                            <td width="170" height="5"></td>
                          </tr>
                          <tr>
                            <td width="170" height="30" bgcolor="#f1da9f"><input name="txt_nome2" class="inputBox" type="text"  id="txt_nome2" style="color:#000000;HEIGHT: 20px; WIDTH: 170px; background: #f1da9f ;border-color : #f1da9f;" value="<%=rs("telefone")%>" size="38" maxlength="33" align="left"></td>
                          </tr>
                          <tr>
                            <td width="170" height="5"></td>
                          </tr>
                          <tr>
                            <td width="170" height="30" bgcolor="#f1da9f"><input name="txt_nome3" class="inputBox" type="text"  id="txt_nome3" style="color:#000000;HEIGHT: 20px; WIDTH: 170px; background: #f1da9f ;border-color : #f1da9f;" value="<%=rs("email")%>" size="38" maxlength="33" align="left"></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="168">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" height="430">&nbsp;</td>
    <td width="594" height="430" style="border:1px solid #FFFFFF;"><table width="584" height="420" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="584" height="420" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="584" height="274"><table width="584" height="274" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></td>
                    </tr>
                    <tr>
                      <td width="188" height="124"  valign="top" ><table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="118" valign="top" bgcolor="#e0a94e"  style="border:1px solid #f9edda;"><select name="combo3" class="inputBox" id="combo3" style="color:#FFFFFF;HEIGHT: 18px; WIDTH: 188px; background:<%=escuro%>" onChange="javascript:atualizacarros2(this.form);">
                                  <option value="<% if rs("cidade") = "n�o informado" or rs666.eof then response.write "cqualquer" else response.write rs666("id_combo1") end if  %>" select><%=rs("cidade")%></option>
                           
                            <% if not rs3.eof then %>
                            <% While NOT Rs3.EoF %>
                            <option value="<% = Rs3("id_combo1") %>" > 
                            <% = Rs3("nome_combo1") %>
                            </option>
                            <% Rs3.MoveNext %>
                            <% Wend %>
                            <%else%>
                            <option value=""></option>
                            <%end if%>
                            <option value="cqualquer">qualquer um</option>
                          </select></td>
  </tr>
</table></td>
                      <td width="10" height="124">&nbsp;</td>
                      <td width="188" height="124"  valign="top" ><select name="combo4" size="8" multiple class="inputBox" id="combo4"  style="color:#FFFFFF;HEIGHT: 124px; WIDTH: 188px; background:<%=escuro%>">
                            <%				 
	  '-----------------------pegar v�rios bairros-----------
  
  
  
dim Variavel
dim Retorno
dim i
Variavel = rs("bairro")
Retorno = Split(Variavel,", ")

i=0

Set rs4 = Server.CreateObject("ADODB.RecordSet")


for i=0 to UBound(Retorno)



strSQL4 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where nome_combo2 like '"& Retorno(i) &"' and cidade_combo2 ='"&rs("cidade")&"' "

 rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs4.ActiveConnection = Conexao
 
 
 
 

rs4.open strSQL4,Conexao,2,1

while not rs4.eof

%>
                            <option value="<%=rs4("id_combo2")%>" selected><%=rs4("nome_combo2")%></option>
                          <%
rs4.MoveNext
Wend

rs4.close




%>
                          <%
next



%>
                        </select></td>
                      <td width="10" height="124">&nbsp;</td>
                      <td width="188" height="124"  valign="top"  ><table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr>
                              <td height="118" valign="top"  ><select name="txt_tipo_vend" multiple size="8" id="txt_tipo_vend" class="inputBox" style="color:#FFFFFF;HEIGHT: 124px; WIDTH: 188px; background: <%=escuro%>">
                                  <%				 '-----------------------pegar v�rios tipos-----------
  
  
  
dim VariavelTipo
dim RetornoTipo
dim iTipo
VariavelTipo = rs("tipo")
RetornoTipo = Split(VariavelTipo,", ")

iTipo=0

Set rs04Tipo = Server.CreateObject("ADODB.RecordSet")


for iTipo=0 to UBound(RetornoTipo)



strSQL04Tipo = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo where tipo like '"& RetornoTipo(iTipo) &"'  ORDER BY tipo ASC"

 
 

rs04Tipo.open strSQL04Tipo,Conexao,2,1

while not (rs04Tipo.eof)

%>
                                  <option value="<%=rs04Tipo("tipo")%>" selected><%=rs04Tipo("tipo")%></option>
                <%
rs04Tipo.MoveNext
Wend

rs04Tipo.close




%>
                <%
next



%>
					 
					 
					 
					 
					 
					 
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
</table>
</td>
                    </tr>
                    <tr>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                        na garagem</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></td>
                    </tr>
                    <tr>
                        <td width="188" height="30"   ><select name="txt_quartos_vend" id="txt_quartos_vend" class="inputBox" style="color:#FFFFFF;HEIGHT: 18px; WIDTH: 188px; background:<%=escuro%>">
                            <option value="<%=rs("quartos")%>" selected><%=rs("quartos")%></option>
					<option value="n�o informado" >n�o informado</option>
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
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30"   ><select name="txt_vagas_vend" id="txt_vagas_vend" class="inputBox" style="color:#FFFFFF;HEIGHT: 20px; WIDTH: 188px; background:<%=escuro%>">
                            <option value="<%=rs("vagas")%>" selected><%=rs("vagas")%></option>
                            <option value="n�o informado" >n�o informado</option>
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
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30"  ><select name="txt_ocupacao_vend" id="txt_ocupacao_vend" class="inputBox" style="color:#FFFFFF;HEIGHT: 20px; WIDTH: 188px; background:<%=escuro%>">
                            <option value="<%=rs("ocupacao")%>" selected><%=rs("ocupacao")%></option>
                            <option value="n�o informado">n�o informado</option>
                <option value="vago">vago</option>
                <option value="alugado">Alugado</option>
                <option value="ocupado por terceiros">Ocupado por terceiros</option>
                <option value="ocupado pelo propriet�rio">Ocupado pelo propriet�rio</option>
                          </select></td>
                    </tr>
                    <tr>
                        <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Atendente</font></td>
                    </tr>
                    <tr>
                        <td width="188" height="30"   ><select name="txt_negociacao_vend" size="1" class="inputBox" id="txt_negociacao_vend"  style="color:#FFFFFF;HEIGHT: 20px; WIDTH: 188px; background:<%=escuro%>">
                            <option value="<%=rs("negociacao")%>" selected><%=rs("negociacao")%></option>
                            <option value="nqualquer" >Qualquer um</option>
                            <option  value="Aluguel">Aluguel</option>
                            <option value="Compra">Compra</option>
                          </select></td>
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30"   ><input name="txt_valor_vend" type="text" id="txt_valor2" size="12" maxlength="12" value="<%=formatnumber(rs("valor"),2)%>" class="inputBox" style="color:#FFFFFF;HEIGHT: 20px; WIDTH: 188px; background:<%=escuro%>"></td>
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30"   ><input type="text" id="stage22" size="12" maxlength="12" value="<%=rs("captacao")%>" class="inputBox" style="color:#FFFFFF;HEIGHT: 20px; WIDTH: 188px; background:<%=escuro%>"></td>
                    </tr>
                  </table></td>
              </tr>
			  <tr>
			      <td height="60"><table width="584" height="60" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Visualiza&ccedil;&otilde;es 
                          do seu im&oacute;vel</font></td>
                        <td width="10" height="30">&nbsp;</td>
                        <td width="188" height="30"><div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                        <td width="10" height="30">&nbsp;</td>
                        <td width="188" height="30"><div align="left"></div></td>
                      </tr>
                      <tr>
                        <td width="188" height="30" bgcolor="#e0a94e" style="border:1px solid #f9edda;"><table width="183" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr>
                              <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%if rs("cliques_no_imovel") <> "" then response.write rs("cliques_no_imovel") else response.write "0" end if%></font></td>
                            </tr>
                          </table></td>
                        <td width="10" height="30">&nbsp;</td>
                        <td width="188" height="30"></td>
                        <td width="10" height="30">&nbsp;</td>
                        <td width="188" height="30">
						<%
						dim sqlproposta01
						'Sqlproposta01 = "SELECT proposta.telefone_proposta,proposta.Nome_proposta,proposta.proposta_proposta,proposta.Cod_proposta,proposta.interesse_proposta,proposta.cod_imovel_proposta  FROM proposta where  cod_imovel_proposta ='"&rs("cod_imovel")&"'" 
	                    
						' <a href="javascript:newWindow2121('visualizar_foto02.asp?varCodimovel=<%=varCod_imovel   
	                   
					   
					   Sqlproposta01 = "Select proposta_oficial.cod_proposta_oficial,proposta_oficial.nome,proposta_oficial.telefone,proposta_oficial.email,proposta_oficial.cod_imovel,proposta_oficial.nacionalidade,proposta_oficial.estado_civil,proposta_oficial.profissao,proposta_oficial.rg,proposta_oficial.cpf,proposta_oficial.endereco,proposta_oficial.cidade,proposta_oficial.bairro,proposta_oficial.estado,proposta_oficial.valor,proposta_oficial.pagamento_vista,proposta_oficial.outro_valor01,proposta_oficial.outro_valor02,proposta_oficial.outro_valor03,proposta_oficial.outro_valor04,proposta_oficial.outro_valor05,proposta_oficial.outro_valor05,proposta_oficial.outro_valor06,proposta_oficial.outro_forma01,proposta_oficial.outro_forma02,proposta_oficial.outro_forma03,proposta_oficial.outro_forma04,proposta_oficial.outro_forma05,proposta_oficial.outro_forma06,proposta_oficial.obs_proposta_oficial,proposta_oficial.nome_contra,proposta_oficial.nacionalidade_contra,proposta_oficial.estado_civil_contra,proposta_oficial.profissao_contra,proposta_oficial.rg_contra,proposta_oficial.cpf_contra,proposta_oficial.endereco_contra,proposta_oficial.cidade_contra,proposta_oficial.bairro_contra,proposta_oficial.estado_contra,proposta_oficial.valor_contra,proposta_oficial.outro_valor_contra01,proposta_oficial.outro_valor_contra02,proposta_oficial.outro_valor_contra03,proposta_oficial.outro_valor_contra04,proposta_oficial.outro_valor_contra05,proposta_oficial.outro_valor_contra06,proposta_oficial.outro_forma_contra01,proposta_oficial.outro_forma_contra02,proposta_oficial.outro_forma_contra03,proposta_oficial.outro_forma_contra04,proposta_oficial.outro_forma_contra05,proposta_oficial.outro_forma_contra06,proposta_oficial.obs_proposta_oficial_contra,proposta_oficial.atendimento,proposta_oficial.data  from proposta_oficial  where  cod_imovel ='"&rs("cod_imovel")&"'"
	 
	 

Set rsproposta01 = Server.CreateObject("ADODB.RecordSet")

	rsproposta01.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsproposta01.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsproposta01.ActiveConnection = Conexao
	
	
	rsproposta01.Open sqlproposta01, Conexao
						
						
						if rsproposta01.recordcount > 0 then
						
						%>
						<font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_contas_proposta02.asp?varCodimovel=<%=varCod_imovel%>" target="_blank" style="color:#000000;text-decoration:none;">Sim , existem <%=rsproposta01.recordcount %>  propostas para o seu im�vel, clique aqui.</a></strong></font>
						<%else%>
						<font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>N�o existem propostas para o seu im�vel atualmente.</strong></font>
						<%end if%>
						<%
						rsproposta01.close

                          set rsproposta01 = nothing
				         %>		
						</td>
                      </tr>
					  
					   <tr>
                        <td width="188" height="30" ><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                          do im&oacute;vel</font></td>
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30" >&nbsp;</td>
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30" >&nbsp;</td>
                    </tr>
					 <tr>
                        <td width="188" height="30" ><input name="stage222" type="text" id="stage222" size="12" maxlength="12" value="<%=rs("cod_imovel")%>" class="inputBox" style="color:#FFFFFF;HEIGHT: 20px; WIDTH: 188px; background:<%=escuro%>"></td>
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30" >
						
						<%
						if vCod444Permuta202 <> "0" then
						%>
						  <div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="indicacao_permuta33.asp?varCodPermuta=<%=vCod444Permuta2022%>" target="_blank" style="color:#000000;text-decoration:none;">Veja 
                            as indica&ccedil;&otilde;es de permuta que temos para 
                            voc&ecirc;, clique aqui</a></strong></font></div>
							<%else%>
							
							<%end if%>
							
							</td>
                      <td width="10">&nbsp;</td>
                        <td width="188" height="30" ><%
						if vCod444Comprador202 <> "" then
						%>
						  <div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="indicacao_imoveis33.asp?varCod_imovel=<%=rs("cod_imovel")%>" target="_blank" style="color:#000000;text-decoration:none;">Veja 
                            as indica&ccedil;&otilde;es de compradores que temos 
                            para o seu im&oacute;vel</a>, clique aqui</strong></font></div>
							<%else%>
							
							<%end if%></td>
                    </tr>
                    </table></td>
			  </tr>
              <tr>
                <td width="584" height="146"><table width="584" height="146" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="584" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                        do seu im&oacute;vel</font></td>
                    </tr>
                    <tr>
                      <td width="585" height="94" ><div align="center">
                            <textarea name="txt_obs_imovel_vend" class="inputBox" id="txt_obs_imovel_vend" style="HEIGHT: 94px; WIDTH: 585px; background:<%=escuro%>; " onKeyPress="return limitfield(this, 800)"><%=rs("obs_imovel")%></textarea>
                          </div></td>
                    </tr>
                    <tr>
                      <td width="584" height="20"><div align="right"></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</form>

  <%

dim stringCidade2,stringBairro2,stringNegociacao2,stringQuartos2,stringValor2,stringTipo2
dim vNegocio
dim vValorMenor,vValorMaior
dim varCodIndicacao

dim varIndicacaoCidade
dim varIndicacaoBairro
dim varIndicacaoNegociacao
dim varIndicacaoQuartos
dim varIndicacaoVagas

dim varIndicacaoValor
dim varIndicacaoTipo


varIndicacaoCidade = rs("cidade")
varIndicacaoBairro = rs("bairro")
varIndicacaoNegociacao = rs("negociacao")
varIndicacaoQuartos = rs("quartos")
varIndicacaoVagas = rs("vagas")
varIndicacaoTipo = rs("tipo")
varIndicacaoValor = rs("valor")
varIndicacaoValor = int(varIndicacaoValor)
vValorMenor = int("0")
vValorMaior = int("0")

dim porcentual



%>
  <%





'dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

'dim negrito,negrito2,varCodComprador



	
 
 
'------------------------Cidade---------------------------

stringIndex2 = " where cod_compradores<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "n�o informado" and rs("cidade") <> "" then
stringCidade2 = " and (cidade='"&rs("cidade")&"' or cidade='"&"n�o informado"&"')"
else
stringCidade2 = ""
end if



 '--------------------------Bairro----------------------------

if rs("bairro") <> "qualquer um" and rs("bairro") <> "n�o informado" and rs("bairro") <> "" then
stringBairro2 = " and (Bairro like '%"&rs("bairro")&"%' or Bairro like'%"&"n�o informado"&"%')"
else
stringBairro2 = ""
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" and  rs("tipo") <> "n�o informado" and rs("tipo") <> "" then
stringTipo2 = " and Tipo like '%"&rs("Tipo")&"%'"
else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negocia��o---------------------------
'vNegocio = "Compra"
if rs("negociacao") = "venda" then
vNegocio = "compra"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if

if  rs("negociacao") <> "qualquer um" and rs("negociacao") <> "" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> 0 and rs("quartos") <> "" then
stringQuartos2 = " and quartos<="&rs("quartos")&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> 0 and rs("vagas") <> "" then
stringVagas2 = " and vagas <="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------




if rs("valor") <> "" and rs("valor") <> "0,00" and rs("valor") <> "0" then
'---------------------------------Valor-----------------------------------



 
   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)
  








stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

else

stringValor2 = ""

end if

'---------------------------------Condominio-----------------------------------



dim stringCondominio101


Porcentual02101 = int(rs("condominio"))*10/100
   


   vCondominioMenor101 = int(rs("condominio")) - int(Porcentual02101)
   vCondominioMaior101 = int(rs("condominio")) + int(Porcentual02101)




if  int(rs("condominio")) <> 0 and rs("condominio") <> ""  then
stringCondominio101 = " and Condominio >="& int(rs("condominio")) &" "
else
stringCondominio101 = ""
end if


'---------------------------------------------------------------------------


'---------------------------------�rea Total-----------------------------------



dim stringAreaTotal101


Porcentual03101 = int(rs("area_total"))*10/100
   


   vAreaTotalMenor101 = int(rs("area_total")) - int(Porcentual03101)
   vAreaTotalMaior101 = int(rs("area_total")) + int(Porcentual03101)



if  int(rs("area_total")) <> 0 and rs("area_total") <> "" then
stringAreaTotal101 = " and area_total >="& vAreaTotalMenor101 &" and area_total <="& vAreaTotalMaior101 &""
else
stringAreaTotal101 = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Su�tes-----------------------------------------


dim stringSuites101
 
if  rs("suites") <> "suiqualquer" and rs("suites") <> "n�o" and rs("suites") <> "0" and rs("suites") <>"00" and rs("suites") <>"" then
stringSuites101 = "  and suites <>'"&"n�o informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites101 = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina101
 
if  rs("piscina") <> "pisqualquer" and rs("piscina") <> "n�o" and rs("piscina") <> "0" and rs("piscina") <>"00" and rs("piscina") <>"" then
stringPiscina101 = "  and piscina <>'"&"n�o informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina101 = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria101
 
if  rs("portaria") <> "porqualquer" and rs("portaria") <> "n�o" and rs("portaria") <> "0" and rs("portaria") <>"00" and rs("portaria") <>"" then
stringPortaria101 = "  and portaria <>'"&"n�o informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria101 = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal101
 
if  rs("quintal") <> "quiqualquer" and rs("quintal") <> "n�o" and rs("quintal") <> "0" and rs("quintal") <>"00" and rs("quintal") <>"" then
stringQuintal101 = "  and quintal <>'"&"n�o informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal101 = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras101
 
if  rs("quadras") <> "quaqualquer" and rs("quadras") <> "n�o" and rs("quadras") <> "0" and rs("quadras") <>"00" and rs("quadras") <>"" then
stringQuadras101 = "  and quadras <>'"&"n�o informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras101 = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Ed�cula--------------------------------------
dim stringEdicula101
 
if  rs("edicula") <> "ediqualquer" and rs("edicula") <> "n�o" and rs("edicula") <> "0" and rs("edicula") <>"00" and rs("edicula") <>"" then
stringEdicula101 = "  and edicula <>'"&"n�o informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula101 = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Ocupa��o--------------------------------------
dim stringOcupacao101
 
if  rs("ocupacao") <> "oqualquer" and rs("ocupacao") <> "n�o informado"  then
stringOcupacao101 = "  and ocupacao ='"&rs("ocupacao")&"'  and ocupacao IS NOT NULL"




else

stringOcupacao101 = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

'stringStandby = " and standby like '"&"suspenso"&"' and standby like '"&"comprador OK"&"'"

stringStandby = " and (standby like 'comprador a contatar' or standby like 'comprador OK') "








'---------------------------------------------------------------------------



	'strSQL444 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.suites,compradores.piscina,compradores.portaria,compradores.area_total,compradores.quintal,compradores.quadras,compradores.condominio,compradores.edicula,compradores.clique   FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio101&stringAreaTotal101&stringSuites101&stringPiscina101&stringPortaria101&stringQuintal101&stringQuadras101&stringEdicula101&stringOcupacao101&stringStandby
	
	dim strSQL2
	strSQL2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.suites,compradores.piscina,compradores.portaria,compradores.area_total,compradores.quintal,compradores.quadras,compradores.condominio,compradores.edicula,compradores.clique,compradores.origem_franquia   FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	




'strSQL2 ="SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	


'----------------------------------------------------Fim da instru��o SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 
  
 
 
 '--------------incluir conta acessada-----------------
 
  dim JaComprador
	 
	 JaComprador = request.querystring("JaComprador")
	 
	 if JaComprador <> "" then
	'Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP2 &"','"& now() &"')"
	JaComprador = "JaExiste"
     else
	 
	 'JaComprador = "JaExiste"
	 Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data,atendimento,origem_franquia) values( '"& rs("proprietario") &"','"& rs("telefone") &"','"& rs("cod_imovel") &"','"& "Im�vel" &"','"& EnderecoIP2 &"','"& now() &"','"& rs("captacao") &"','"& rs("origem_franquia") &"')"
	
	JaComprador = "JaExiste"
	 end if
  
 
 
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 
dim rs2


Set RS2 = Server.CreateObject("ADODB.Recordset")
'um objeto recordset � inst�nciado.

Dim LinkTemp
'essa vari�vel vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
'as vari�veis acima s�o usadas para trocar a cor das tabelas que conter�o os valores
'dos recordsets.






dim intPage
'essa vari�vel vai receber um valor inicial "1" que mostra que estamos na primeira p�gina.

dim intPageCount
'Essa vari�vel vai receber o valor da quantidade de p�ginas do recordset.

dim intRecordCount
'Essa vari�vel vai receber o n�mero de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a vari�vel intPage recebe o valor "1" na primeira p�gina.
	
RS2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

RS2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

RS2.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conex�o o recordset utilizar�.
	
RS2.Open strSQL2, Conn, 1, 3
'o recordset � aberto
	
RS2.PageSize = 5
'Aqui configura-se o recordset para 20 registros por p�gina.

RS2.CacheSize = RS2.PageSize
'o Cache tamb�m conter� 20 registros por p�gina.

intPageCount = RS2.PageCount
'A vari�vel intPageCount recebe o valor do n�mero de p�gina do recordset retornado.

intRecordCount = RS2.RecordCount
'A vari�vel intRecordCount recebe o valor do n�mero de registros retornados no recordset.


dim vMostrar001
vMostrar001 = "n�o"



If NOT (RS2.BOF AND RS2.EOF) and vMostrar001 <> "n�o" Then
'verifica se existem registros retornados.
%>
 <table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"> <strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Veja as 
  indica��es de compradores para o seu im�vel.</font></strong> <br>
  <br></td>
  </tr>
</table>

  <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount 
'se intPage � maior que o n�mero de p�ginas ent�o intPage � igual ao n�mero de p�ginas.

	If CInt(intPage) <= 0 Then intPage = 1 
	'se intPage � menor ou igual a zero ent�o intPage igual a "1"
	'a vari�vel intPage sempre vai ser for�ada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados ent�o.
			 
			 RS2.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a p�gina exata que o registro atual
			'reside
			
			intStart = RS2.AbsolutePosition
			'a vari�vel intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posi��o exata do primeiro registro da p�gina correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage � igual ao n�mero de p�ginas no recordset , estamos na �ltima 
			'p�gina ent�o.
				intFinish = intRecordCount
				'a vari�vel intFinish recebe o valor do n�mero do �ltimo recordset.
				'intFinish corresponde ao valor do �ltimo registro da p�gina correspondente.
			Else
				intFinish = intStart + (RS2.PageSize - 1)
				'a vari�vel intFinish recebe o valor de intStart + o valor
				'do n�mero de registros na p�gina menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros ent�o
		For intRecord = 1 to RS2.PageSize
		'um contador inRecord � colocado at� o n�mero de registros na p�gina.
%>
  <% varCodComprador = rs2("cod_compradores") %>

<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="170"><table width="794" height="170" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="210">&nbsp;</td>
          <td width="584" height="170" style="border:1px solid #FFFFFF;"><table width="574" height="160" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="574" bgcolor="#e0a94e"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow3('visualizar_comprador01.asp?varCodCompradores=<%=varCodComprador%>')" style="color:#FFFFFF;text-decoration:none;" >Ol� 
                    , meu nome � <strong><%=rs2("nome")%></strong>, o sitema veja 
                    analizou os dados do seu im�vel e o que eu desejo comprar, 
                    e detectou a possibilidade de neg�cio entre n�s. Lique j� 
                    para <strong>4123-72-44</strong> e fale com o meu atendente 
                    o sr(a) <strong>
                    <%if rs2("atendimento") = "Spirity" or rs2("atendimento") = "internet" then response.write "Wanderlei" else response.write rs2("atendimento") end if%>
                    </strong>, para que o mesmo agende uma visita minha ao seu 
                    im�vel, <strong>clique aqui</strong> e saiba mais sobre meus 
                    interesses e condi��es de pagamento. Muito Obrigado. </a></font></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<br>

 <%
RS2.MoveNext


	  





 'acima � feito a troca de cores das tabelas e do texto dos recordsets.

If RS2.EOF Then Exit for

Next

	
%>

<%end if%>
<%end if%>

<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"><table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a p�gina atual for maior que "1" ent�o o link anteriro � colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&varCod_imovel=<%=varCod_imovel%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoValor=<%=varIndicacaoValor%>&JaComprador=<%=JaComprador%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" size="1" > 
            </font></div></td>
          
        <td> 
          <%If cInt(intPage) < cInt(intPageCount)  Then%>
          <!-- se intPage � menor que o n�mero de p�ginas ent�o colocar o bot�o pr�ximo -->
          <a href="?page=<%=intPage + 1%>&varCod_imovel=<%=varCod_imovel%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoValor=<%=varIndicacaoValor%>&JaComprador=<%=JaComprador%>"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pr�ximo</strong></font></a></td><% end if%>
        </tr>
      </table></td>
  </tr>
</table>



  

<%
Function EscreveFuncaoJavaScript ( Conexao )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 

Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsMarcas3.ActiveConnection = Conexao
	
	
	rsMarcas3.Open SqlMarcas3, Conexao



While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsCarros3.ActiveConnection = Conexao
	
	
	rsCarros3.Open SqlCarros3, Conexao



'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros3.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"
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
Function EscreveFuncaoJavaScript2 ( Conexao )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo3.options[doublecombo.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas33 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsMarcas33.ActiveConnection = Conexao
	
	
	rsMarcas33.Open SqlMarcas33, Conexao





While NOT (rsMarcas33.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas33("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros33 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 WHERE id_combo1 =" & rsMarcas33("id_combo1")&" order by nome_combo2"






Set rsCarros33 = Server.CreateObject("ADODB.RecordSet")

	rsCarros33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsCarros33.ActiveConnection = Conexao
	
	
	rsCarros33.Open SqlCarros33, Conexao





'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
 
While NOT (rsCarros33.EoF)

Response.Write "doublecombo.combo4.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend

Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d�vida da sua utiliza��o! 
Response.Write "break;" & vbcrlf

'Pr�xima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da fun��o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 





rsMarcas33.Close           
		   
           Set rsMarcas33 = Nothing
             
			rsCarros33.Close           
		   
           Set rsCarros33 = Nothing 





End Function
%> 



<%  EscreveFuncaoJavaScript2 ( Conexao ) %>
</body>
</html>
