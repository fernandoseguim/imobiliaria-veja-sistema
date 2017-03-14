<!--#include file="dsn.asp"-->
<!--#include file="style_conta.asp"-->
<!--#include file="cores02.asp"-->






<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if


'Criando conexão com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn



%>



<%

dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	dim strSQL
	dim rs
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais,permuta.origem_franquia  FROM permuta where cod_permuta="&varCodPermuta
	 rs.CursorLocation = 3
      rs.CursorType = 3
	 rs.Open strSQL, Conexao




'-----------------------Acrescentar acessos------------------------------------

'------------------Verifica se o internauta já tem conta---------------------------
  
  dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&rs("telefone")&"%' or telefone02 like '%"&rs("telefone")&"%' or telefone03 like '%"&rs("telefone")&"%'" 
	
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

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


'--------------------------------abrindo cidade atual------------------------------
dim Sql3
dim rs3

Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1"

Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao
	
	
	rs3.Open Sql3, Conexao




dim rs666,strSQL666
   
    Set rs666 = Server.CreateObject("ADODB.RecordSet")
	strSQL666 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 where nome_combo1 ='"&rs("cidade_vend")&"'  ORDER BY nome_combo1" 
	
	
	
	rs666.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs666.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs666.ActiveConnection = Conexao
	
	
	
	
	 rs666.Open strSQL666, Conexao

'-----------------------------------------------------------------------------------


'--------------Anrindo bairro atual-------------------------------------------

dim Sql888
dim rs888

Sql888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 



Set rs888 = Server.CreateObject("ADODB.RecordSet")

	rs888.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs888.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs888.ActiveConnection = Conexao
	
	
	rs888.Open Sql888, Conexao
	








'--------------------------------------------------------------------------------


'--------------abrindo cidade procurada------------------------------------------

'Abrindo a tabela MARCAS!
Sql4 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 






Set rs5 = Server.CreateObject("ADODB.RecordSet")

	rs5.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs5.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs5.ActiveConnection = Conexao
	
	
	rs5.Open Sql4, Conexao
	




dim rs777,strSQL777
   
    Set rs777 = Server.CreateObject("ADODB.RecordSet")
	strSQL777 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 where nome_combo1 ='"&rs("cidade_comp")&"'  ORDER BY nome_combo1" 
	 
	 
	 
	 
	rs777.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs777.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs777.ActiveConnection = Conexao
	
	 
	 
	 
	 
	 
	 rs777.Open strSQL777, Conexao





'------------------Selecionar os tipos do seu imóvel-------------------------------

'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo FROM tipo  ORDER BY tipo ASC" 
	 
	 
	 rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao
	 
	 
	 
	 
	 rs444Tipo22.Open strSQL444Tipo22, Conexao











'--------------------------------------------------------------------------------------



'--------------------------------abrir tipo pretendido--------------

 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo FROM tipo  ORDER BY tipo ASC" 
	
	rs444Tipo23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo23.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo23.ActiveConnection = Conexao
	
	
	
	
	
	
	 rs444Tipo23.Open strSQL444Tipo23, Conexao






'------------------------------------------------------------------



'-----------------Abrir bairro pretendido---------------------

dim rs8888,strSQL8888
   
    Set rs8888 = Server.CreateObject("ADODB.RecordSet")
	strSQL8888 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2   FROM combo2 where nome_combo2 ='"&rs("bairro_vend")&"' and cidade_combo2 ='"&rs("cidade_vend")&"'  ORDER BY nome_combo2" 
	
	
	
	
	rs8888.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs8888.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs8888.ActiveConnection = Conexao
	
	
	
	
	 rs8888.Open strSQL8888, Conexao

%>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow3.focus( )
   }

</SCRIPT>

</head>

<body bgcolor="#e6dca9">
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_conta_permuta01.asp?varCodPermuta=<%=varCodPermuta%>">

<table width="794" height="900" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" height="900" valign="top"><table width="190" height="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="190" height="262" style="border:1px solid #FFFFFF;"><table width="180" height="252" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="180" height="252" bgcolor="#e0a94e"> 
                  <table width="170" height="242" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="170" height="242"><table width="170" height="242" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="170" height="137"><img src="icone_conta02.jpg" width="170" height="137"></td>
                          </tr>
                          <tr>
                            <td width="170" height="5"></td>
                          </tr>
                          <tr>
                            <td width="170" height="30" bgcolor="#f1da9f"><input name="txt_nome" class="inputBox" type="text"  id="txt_nome" style="color:#000000;HEIGHT: 20px; WIDTH: 170px; background: #f1da9f ;border-color : #f1da9f;" value="<%=rs("nome")%>" size="38" maxlength="33" align="left"></td>
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
          <td width="190" height="638">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" height="900">&nbsp;</td>
    <td width="594" height="900"><table width="594" height="900" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="594" height="430" style="border:1px solid #FFFFFF;"><table width="584" height="420" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="584" height="420" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="584" height="274"><table width="584" height="274" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                              do seu im&oacute;vel</font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                              do seu im&oacute;vel </font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                              do seu im&oacute;vel </font></td>
                    </tr>
                    <tr>
                            <td width="188" height="124" bgcolor="#e0a94e" style="border:1px solid #f9edda;" valign="top"><table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="188" height="20" ><select name="combo1" class="inputBox" id="combo1" style="color:#FFFFFF;HEIGHT: 18px; WIDTH: 188px; background:<%=escuro%>" onChange="javascript:atualizacarros(this.form);">
                <option value="<% if rs("cidade_vend") = "não informado" or rs("cidade_vend") = "qualquer um" or   rs666.eof  then response.write "cqualquer" else response.write rs666("id_combo1") end if  %>" select>
                <% if rs("cidade_vend") <> "cqualquer" and rs("cidade_vend") <> "" then response.write rs("cidade_vend") else response.write "não informado" end if  %>
                </option>
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
              </select></td>
  </tr>
</table></td>
                      <td width="10" height="124">&nbsp;</td>
                      <td width="188" height="124" bgcolor="#e0a94e"  style="border:1px solid #f9edda;" valign="top"><table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="188" height="20" ><select name="combo2" onChange="javascript:atualizacarros888(this.form);" class="inputBox" id="combo2" style="HEIGHT: 18px; WIDTH: 190px; background:<%=escuro%>">
                <option value="<% if rs("bairro_vend") = "não informado" or rs("bairro_vend") = "qualquer um" or rs("bairro_vend") = "bqualquer" or  rs888.eof  then response.write "bqualquer" else response.write rs8888("id_combo2") end if  %>" select>
                <% if rs("bairro_vend") <> "bqualquer" and rs("bairro_vend") <> "" then response.write rs("bairro_vend") else response.write "não informado" end if  %>
                </option>
              </select></td>
  </tr>
</table></td>
                      <td width="10" height="124">&nbsp;</td>
                      <td width="188" height="124" bgcolor="#e0a94e"  style="border:1px solid #f9edda;" valign="top"><table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr>
                                    <td width="188" height="20" ><select name="txt_tipo" size="1" id="txt_tipo" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=escuro%>">
                <option value="<%if rs("tipo_vend") <> "tqualquer" then response.write rs("tipo_vend") else response.write "tqualquer" end if%>">
                <%if rs("tipo_vend") <> "tqualquer" then response.write rs("tipo_vend") else response.write "não informado" end if%>
                </option>
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
</table></td>
                    </tr>
                    <tr>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Quartos 
                              no seu im&oacute;vel</font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                              no seu im&oacute;vel</font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o 
                              do seu im&oacute;vel</font></td>
                    </tr>
                    <tr>
                            <td width="188" height="30" ><select name="txt_quartos_vend" size="1" id="txt_quartos_vend" class="inputBox" style="HEIGHT: 30px; WIDTH: 190px; background: <%=escuro%>">
                <option value="<%=rs("quartos_vend")%>" selected>
                <% if rs("quartos_vend") = "0" then response.write "não informado" else response.write rs("quartos_vend") end if%>
                </option>
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
              </select></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30" ><select name="txt_vagas_vend" size="1" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=escuro%>">
                <option value="<%=rs("vagas_vend")%>" selected>
                <% if rs("vagas_vend") = "0" then response.write "não informado" else response.write rs("vagas_vend") end if%>
                </option>
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
              </select></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30" ><table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%="Não informado"%></font></td>
                                  </tr>
                                </table></td>
                    </tr>
                    <tr>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                              que voc&ecirc; quer</font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                              do seu im&oacute;vel</font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Atendente</font></td>
                    </tr>
                    <tr>
                              <td width="188" height="30"  >
<table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%="Permuta"%></font></td>
                                  </tr>
                                </table></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30" ><input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 190px; background:<%=escuro%>" value="<%=FormatNumber(rs("valor_vend"),2)%>" size="12" maxlength="13"></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30"  >
<table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("atendimento")%></font></td>
                                  </tr>
                                </table></td>
                    </tr>
					<tr>
                              <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                                de permuta</font></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30">&nbsp;</td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30">&nbsp;</td>
                    </tr>
					<tr>
                              <td width="188" height="30"><table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("cod_permuta")%></font></td>
                                  </tr>
                                </table></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30">&nbsp;</td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30">&nbsp;</td>
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
                      <td width="585" height="94" >
                                <textarea name="txt_descricao_vend" class="inputBox" id="txt_descricao_vend" style="HEIGHT: 100px; WIDTH: 585px; background:<%=escuro%>" onKeyPress="return limitfield(this, 800)"><%=rs("descricao_vend")%></textarea></td>
                    </tr>
                    <tr>
                      <td width="584" height="20">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
        </tr>
        <tr>
          <td width="594" height="40">&nbsp;</td>
        </tr>
        <tr>
          <td width="594" height="430" style="border:1px solid #FFFFFF;"><table width="584" height="420" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="584" height="420" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="584" height="274"><table width="584" height="274" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                              pretendida </font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                              pretendido </font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                              pretendido </font></td>
                    </tr>
                    <tr>
                      <td width="188" height="124" bgcolor="#e0a94e" style="border:1px solid #f9edda;" valign="top"><table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="188" height="20" ><select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 190px; background:<%=escuro%>" onChange="javascript:atualizacarros2(this.form);">
              <option value="<% if rs("cidade_comp") = "não informado" or rs("cidade_comp") = "qualquer um" or   rs777.eof  then response.write "cqualquer" else response.write rs777("id_combo1") end if  %>" select>
              <% if rs("cidade_comp") <> "cqualquer" and rs("cidade_comp") <> "" then response.write rs("cidade_comp") else response.write "não informado" end if  %>
              </option>
              <% if not rs5.eof then %>
              <% While NOT Rs5.EoF %>
              <option value="<% = Rs5("id_combo1") %>" <% if rs5("nome_combo1") = rs("cidade_comp") then%>selected<%else%><%end if%>> 
              <% = Rs5("nome_combo1") %>
              </option>
              <% Rs5.MoveNext %>
              <% Wend %>
              <%else%>
              <option value=""></option>
              <%end if%>
              <option value="cqualquer">qualquer um</option>
            </select></td>
  </tr>
</table>
</td>
                      <td width="10" height="124">&nbsp;</td>
                              <td width="188" height="124" bgcolor="#e0a94e"  style="border:1px solid #f9edda;" valign="top"><select name="combo4"  class="inputBox" id="combo4" style="HEIGHT: 160px; WIDTH: 190px; background:<%=escuro%>" multiple size="8">
              <%				 
	  '-----------------------pegar vários bairros-----------
  
  
  
dim Variavel
dim Retorno
dim i
Variavel = rs("bairro_comp")
Retorno = Split(Variavel,", ")

i=0

Set rs4 = Server.CreateObject("ADODB.RecordSet")


for i=0 to UBound(Retorno)



strSQL4 = "select  combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where nome_combo2 like '"& Retorno(i) &"' and cidade_combo2 ='"&rs("cidade_comp")&"' "

 
 rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

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
                              <td width="188" height="124" bgcolor="#e0a94e"  style="border:1px solid #f9edda;" valign="top"><select name="txt_tipo2" multiple size="8" id="txt_tipo2" class="inputBox" style="HEIGHT: 160px; WIDTH: 190px; background: <%=escuro%>">
             
	 <%				 
	  '-----------------------pegar vários bairros-----------
  
  
  
dim VariavelTipo
dim RetornoTipo
dim iTipo
VariavelTipo = rs("tipo_comp")
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
            </select></td>
                    </tr>
                    <tr>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Quartos 
                              pretendidos </font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                              pretendidas</font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o 
                              pretendida </font></td>
                    </tr>
                    <tr>
                              <td width="188" height="30" ><select name="txt_quartos_comp" size="1" id="txt_quartos_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=escuro%>">
              <option value="<%=rs("quartos_comp")%>" selected>
              <% if rs("quartos_comp") = "0" then response.write "não informado" else response.write rs("quartos_comp") end if%>
              </option>
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
            </select></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30" ><select name="txt_vagas_comp" size="1" id="txt_vagas_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 190px; background: <%=escuro%>">
              <option value="<%=rs("vagas_comp")%>" selected>
              <% if rs("vagas_comp") = "0" then response.write "não informado" else response.write rs("vagas_comp") end if%>
              </option>
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
            </select></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30"  >
<table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%="Não informado"%></font></td>
                                  </tr>
                                </table></td>
                    </tr>
                    <tr>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                              pretendida </font></td>
                      <td width="10">&nbsp;</td>
                            <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                              pretendido </font></td>
                      <td width="10">&nbsp;</td>
                      <td width="188" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Atendente</font></td>
                    </tr>
                    <tr>
                              <td width="188" height="30"  >
<table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%="Permuta"%></font></td>
                                  </tr>
                                </table></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30" ><input name="txt_valor_comp" type="text" class="inputBox" id="txt_valor_comp" style="HEIGHT: 18px; WIDTH: 190px; background:<%=escuro%>" value="<%=FormatNumber(rs("valor_comp"),2)%>" size="12" maxlength="13"></td>
                      <td width="10">&nbsp;</td>
                              <td width="188" height="30"  >
<table width="188" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td bgcolor="#e0a94e" style="border:1px solid #f9edda;"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("atendimento")%></font></td>
                                  </tr>
                                </table></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="584" height="146"><table width="584" height="146" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="584" height="30"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                        do im&oacute;vel que voc&ecirc; quer</font></td>
                    </tr>
                    <tr>
                      <td width="585" height="94" bgcolor="#e0a94e"><textarea name="txt_descricao_comp" class="inputBox" id="txt_descricao_comp" style="HEIGHT: 100px; WIDTH: 585px; background:<%=escuro%>" onKeyPress="return limitfield(this, 800)"><%=rs("descricao_comp")%></textarea></td>
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
      </table></td>
  </tr>
</table>
</form>


<%







dim negrito,negrito2
dim vValor_vend,vValor_vend1,vValor_vend2
dim vValor_comp,vValor_comp1,vValor_comp2
dim vCidade_vend,vCidade_comp


dim varIndicacaoCidadeVend
dim varIndicacaoBairroVend
dim varIndicacaoVilaVend
dim varIndicacaoQuartosVend
dim varIndicacaoVagasVend
dim varIndicacaoValorVend
dim varIndicacaoTipoVend


dim varIndicacaoCidadeComp
dim varIndicacaoBairroComp
dim varIndicacaoVilaComp
dim varIndicacaoQuartosComp
dim varIndicacaoVagasComp
dim varIndicacaoValorComp
dim varIndicacaoTipoComp

dim varIndicacaoCodigo


 varIndicacaoCidadeVend = rs("cidade_vend")
 varIndicacaoBairroVend = rs("bairro_vend")
 varIndicacaoVilaVend = rs("vila_vend")
 varIndicacaoQuartosVend = rs("quartos_vend")
 varIndicacaoVagasVend = rs("vagas_vend")
 varIndicacaoValorVend = rs("valor_vend")
 varIndicacaoTipoVend = rs("tipo_vend")
 
 
 
 session("varIndicacaoCidadeVend") = varIndicacaoCidadeVend
 session("varIndicacaoBairroVend") = varIndicacaoBairroVend
 session("varIndicacaoVilaVend") = varIndicacaoVilaVend
 session("varIndicacaoQuartosVend") = varIndicacaoQuartosVend
 session("varIndicacaoVagasVend") = varIndicacaoVagasVend
 session("varIndicacaoValorVend") = varIndicacaoValorVend
 session("varIndicacaoTipoVend") = varIndicacaoTipoVend
 
 
 
 
 
 
 varIndicacaoCidadeComp = rs("cidade_comp")
 varIndicacaoBairroComp = rs("bairro_comp")
 varIndicacaoVilaComp = rs("vila_comp")
 varIndicacaoQuartosComp = rs("quartos_comp")
 varIndicacaoVagasComp = rs("vagas_comp")
 varIndicacaoValorComp = rs("valor_comp")
 varIndicacaoTipoComp = rs("tipo_comp")
 
 
 session("varIndicacaoCidadeComp") = varIndicacaoCidadeComp
 session("varIndicacaoBairroComp") = varIndicacaoBairroComp
 session("varIndicacaoVilaComp") = varIndicacaoVilaComp
 session("varIndicacaoQuartosComp") = varIndicacaoQuartosComp
 session("varIndicacaoVagasComp") = varIndicacaoVagasComp
 session("varIndicacaoValorComp") = varIndicacaoValorComp
 session("varIndicacaoTipoComp") = varIndicacaoTipoComp
 
 
 
 varIndicacaoCodigo=request.querystring("varIndicacaoCodigo")
 
session("varIndicacaoCodigo") = varIndicacaoCodigo


 
 
 
 
 
 
 
 
 
 
  '---------Selecionar permutante pelo telefone------------------------------------------------
		   
		     dim rs202,SQL444Permuta202
 Set rs202 = Server.CreateObject("ADODB.RecordSet")
 SQL444Permuta202 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone='"& rs("telefone")&"' order by cod_permuta DESC" 
	
	
	rs202.CursorLocation = 3
         rs202.CursorType = 3
           rs202.ActiveConnection = Conexao
	
	
	rs202.open SQL444Permuta202,Conexao,2,1  
	
			
	if  not rs202.eof then
		   
		   
		   
		   
		   
		   
'------------------------Sua Cidade--------------------------

stringIndex202 = " where cod_permuta<>"&"0"&""
 
 
 
  if   rs202("cidade_vend") = "não informado" or rs202("cidade_vend") = "" or rs202("cidade_vend") = "cqualquer" or  rs202("cidade_vend") = "qualquer um" then
	stringCidadeVend202 = ""
 else

stringCidadeVend202 = " and (Cidade_comp='"&rs202("cidade_vend")&"' or Cidade_comp='"&"não informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend202

 if   rs202("bairro_vend") = "não informado" or rs202("bairro_vend") = "" or rs202("bairro_vend") = "bqualquer" or  rs202("bairro_vend") = "qualquer um" then
	stringBairroVend202 = ""
 else
'stringBairroVend = ""
stringBairroVend202 = " and (Bairro_comp like'%"&rs202("bairro_vend")&"%' or Bairro_comp like'%"&"não informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend202

'" and (Vila_comp='"&rs("vila_vend")&"' or Vila_comp='"&"não informado"&"' or Vila_comp='"&"vlqualquer"&"' or Vila_comp='"&"qualquer um"&"' )"

 if   rs202("vila_vend") = "não informado" or rs202("vila_vend") = "" or rs202("vila_vend") = "vlqualquer" or rs202("vila_vend") = "qualquer um" then
	stringVilaVend202 =  ""
 else

stringVilaVend202 = ""

end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend202
 
 
 if rs202("tipo_vend") = "não informado" or rs202("tipo_vend") = "" or rs202("tipo_vend") = "tqualquer" or rs202("tipo_vend") = "qualquer um"  then

stringTipoVend202 = ""

else
stringTipoVend202 = " and Tipo_comp like '%"&rs202("tipo_vend")&"%'"
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend202
 
 
 

stringQuartosVend202 = " and Quartos_comp <="&int(rs202("quartos_vend"))&""

 


 '-----------------------Número de Vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend202
 
 
 



stringVagasVend202 = " and vagas_comp <="&int(rs202("vagas_vend"))&""

 




 
 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
dim PorcentualVend202

dim vValorMenorVend202
dim vValorMaiorVend202

PorcentualVend202 = int(rs202("valor_vend"))*20/100

   


   vValorMenorVend202 = int(rs202("valor_vend")) - int(PorcentualVend202)
   vValorMaiorVend202 = int(rs202("valor_vend")) + int(PorcentualVend202)

 
 
 
 
 
	 dim stringValorVend202
  
	
	
	
	stringValorVend202 = " and Valor_comp >="&  vValorMenorVend202 &""
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp202
  if rs202("cidade_comp")="não informado" or rs202("cidade_comp")="" or rs202("cidade_comp")="cqualquer" or rs202("cidade_comp") = "qualquer um" then
	stringCidadeComp202 = ""
	else
	
	stringCidadeComp202 = " and Cidade_vend ='"& rs202("cidade_comp") &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp202

	if rs202("bairro_comp") = "não informado" or  rs202("bairro_comp") = "" or  rs202("bairro_comp") = "bqualquer" or rs202("bairro_comp") = "qualquer um" then
	
	
	
	
	
	stringBairroComp202 = ""
	
	
	
	
	else
	
	
	
	'stringBairroComp = " and Bairro_vend ='"& rs("bairro_comp") &"'"
	
	
	
	
 
dim Numero_Indicacoes202
dim Numero_Indicacoes02202




Numero_Indicacoes202 = 0
Numero_Indicacoes02202 = 0


dim soma02202
dim soma202

soma202 = 0
soma02202 = 0

dim Variavel202
dim Retorno202
dim contar202
Variavel202 = rs202("bairro_comp")
Retorno202 = Split(rs202("bairro_comp"),", ")

contar202=0

dim stringBairro3202
dim stringBairro4202
dim stringBairro5202

for contar202=0 to UBound(Retorno202)

stringBairro3202 = "and ( "
stringBairro4202 = " Bairro_vend='"&Retorno202(contar202)&"'or  " &stringBairro4202

stringBairro5202 = " cod_permuta=0)"


stringBairroComp202 = stringBairro3202&stringBairro4202&stringBairro5202



next


stringBairro3202 = ""
stringBairro4202 = ""
stringBairro5202 = ""

	
	
	

	
	
	end if
	
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 'and Vila_vend ='"& rs("vila_comp") &"'
	 dim stringVilaComp202

	if rs202("vila_comp") <> "não informado" and rs202("vila_comp") <> "" and rs202("vila_comp") <> "vlqualquer" and rs202("vila_comp") <> "qualquer um" then
	stringVilaComp202 = ""
	else
	
	stringVilaComp202 = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp
  'if rs("tipo_comp")="não informado" or rs("tipo_comp")="" or rs("tipo_comp")="tqualquer" or rs("tipo_comp") = "qualquer um" then
	'stringTipoComp = ""
	'else
	
	
	'stringTipoComp = " and Tipo_vend ='"& rs("tipo_comp")&"'"
	'end if
	
	
	
	'--------------------------Tipo----------------------------

if rs202("tipo_comp") <> "qualquer um" and rs202("tipo_comp") <> "não informado" and rs202("tipo_comp") <> "" then




 
dim Numero_IndicacoesTipoComp202
dim Numero_Indicacoes02TipoComp202




Numero_IndicacoesTipoComp202 = 0
Numero_Indicacoes02TipoComp202 = 0


dim soma02TipoComp202
dim somaTipoComp202

somaTipoComp202 = 0
soma02TipoComp202 = 0

dim VariavelTipoComp202
dim RetornoTipoComp202
dim contarTipoComp202
VariavelTipoComp202 =  rs202("tipo_comp")
RetornoTipoComp202 = Split(rs202("tipo_comp"),", ")

contarTipoComp202=0

dim stringTipo3Comp202
dim stringTipo4Comp202
dim stringTipo5Comp202

for contarTipoComp202=0 to UBound(RetornoTipoComp202)

stringTipo3Comp202 = "and ( "
stringTipo4Comp202 = " tipo_vend='"&RetornoTipoComp202(contarTipoComp202)&"'or  " &stringTipo4Comp202

stringTipo5Comp202 = " cod_permuta=0)"


stringTipo2Comp202 = stringTipo3Comp202&stringTipo4Comp202&stringTipo5Comp202







next

stringTipo3Comp202 = ""
stringTipo4Comp202 = ""
stringTipo5Comp202 = ""


else
stringTipo2Comp202 = ""
end if

	
	
	
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp202
  
	
	stringQuartosComp202 = " and Quartos_vend >="& int(rs202("quartos_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 

	 dim stringVagasComp202
 
	
	stringVagasComp202 = " and vagas_vend >="& int(rs202("vagas_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------

dim PorcentualComp202

dim vValorMenorComp202
dim vValorMaiorComp202

PorcentualComp202 = int(rs202("valor_comp"))*20/100

   


   vValorMenorComp202 = int(rs202("valor_comp")) - int(PorcentualComp202)
   vValorMaiorComp202 = int(rs202("valor_comp")) + int(PorcentualComp202)


	 dim stringValorComp202
  
	
	
	'stringValorComp202 = " and Valor_vend >="& vValorMenorComp202 &" and Valor_vend <="& vValorMaiorComp202 &""
	
	stringValorComp202 = " and Valor_vend <="& int(vValorMaiorComp202) &""
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	'strSQL444 = "SELECT * FROM permuta"
	'&stringVilaVend
	'&stringVilaComp
	varIndicacaoCodigo202=rs202("cod_permuta")
	
	dim strSQL2
	
	strSQL2 = "SELECT permuta.cod_permuta   FROM permuta"&stringIndex202&stringCidadeVend202&stringBairroVend202&stringVilaVend202&stringTipoVend202&stringQuartosVend202&stringVagasVend202&stringValorVend202&stringCidadeComp202&stringBairroComp202&stringVilaComp202&stringTipo2Comp202&stringQuartosComp202&stringVagasComp202&stringValorComp202&" and standby <> 'incluido' and cod_permuta not like "&varIndicacaoCodigo202
	
 
	
'---------------------------------------------------------------	
	
	
	'strSQL2 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringVilaVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringVilaComp&stringTipo2Comp&stringQuartosComp&stringVagasComp&stringValorComp&" and cod_permuta not like "&varCodPermuta
	
	if vNome = "" then
	vNome = "não informado"
	end if
	
	if vTelefone = "" then
	vTelefone = "não informado"
	end if
	
	
	 dim vEnderecoIP , vData
  vData = now()
  
 
 vEnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	
	
  
  '--------------incluir conta acessada-----------------
 
  dim JaComprador
	 
	 JaComprador = request.querystring("JaComprador")
	 
	 if JaComprador <> "" then
	'Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_compradores") &"','"& "Compradores" &"','"& EnderecoIP2 &"','"& now() &"')"
	JaComprador = "JaExiste"
     else
	 
	 'JaComprador = "JaExiste"
	 Conexao.execute"Insert into contas_procuradas(nome,telefone,codigo_conta,tipo_conta,endereco_ip,data,atendimento,origem_franquia) values( '"& rs("nome") &"','"& rs("telefone") &"','"& rs("cod_permuta") &"','"& "Permuta" &"','"& EnderecoIP2 &"','"& now() &"','"& rs("atendimento") &"','"& rs("origem_franquia") &"')"
	
	JaComprador = "JaExiste"
	 end if
  
 
 
  
  
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 

dim rs2

Set RS2 = Server.CreateObject("ADODB.Recordset")
'um objeto recordset é instânciado.

Dim LinkTemp
'essa variável vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
'as variáveis acima são usadas para trocar a cor das tabelas que conterão os valores
'dos recordsets.






dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeira página.

dim intPageCount
'Essa variável vai receber o valor da quantidade de páginas do recordset.

dim intRecordCount
'Essa variável vai receber o número de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a variável intPage recebe o valor "1" na primeira página.
	
RS2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS2.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	
RS2.Open strSQL2, Conn, 1, 3
'o recordset é aberto
	
RS2.PageSize = 10
'Aqui configura-se o recordset para 20 registros por página.

RS2.CacheSize = RS2.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS2.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS2.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.



If NOT (RS2.BOF AND RS2.EOF) Then
'verifica se existem registros retornados.
%>


<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="210">&nbsp;</td>
    <td width="584" align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Veja 
      abaixo , as indicações de permuta para você.</font></strong> <br>
	<br></td>
  </tr>
</table>




<%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage é menor ou igual a zero então intPage igual a "1"
	'a variável intPage sempre vai ser forçada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados então.
			 
			 RS2.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a página exata que o registro atual
			'reside
			
			intStart = RS2.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posição exata do primeiro registro da página correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual ao número de páginas no recordset , estamos na última 
			'página então.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor do número do último recordset.
				'intFinish corresponde ao valor do último registro da página correspondente.
			Else
				intFinish = intStart + (RS2.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'do número de registros na página menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros então
		For intRecord = 1 to RS2.PageSize
		'um contador inRecord é colocado até o número de registros na página.

%>
<br>
<% varCodPermuta =RS2("cod_permuta") %>
<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="170"><table width="794" height="170" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="210">&nbsp;</td>
          <td width="584" height="170" style="border:1px solid #FFFFFF;"><table width="574" height="160" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><div align="center"><font face="Verdana, arial" size="1" color="FFFFFF"><a href="javascript:newWindow3('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')" style="color:#FFFFFF;text-decoration:none;" ><%=RS("descricao_vend")%></a></font></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>


 <%
RS2.MoveNext


	  




If RS2.EOF Then Exit for
Next	
%>


<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><font face="Verdana, arial" size="1"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
               <a href="?page=<%=intPage - 1%>&varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" style="color:#000000">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" >
              
			  <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
			  
             
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
             <a href="?page=<%=intPage + 1%>&varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend%>&varIndicacaoValorVend=<%=varIndicacaoValorVend%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp%>&varIndicacaoValorComp=<%=varIndicacaoValorComp%>&varIndicacaoCodigo=<%=varIndicacaoCodigo%>" style="color:#000000"><b>Próximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table>

<%End If


Else

%>




<%end if%>

<% end if%>


<%
Function EscreveFuncaoJavaScript ( Conexao )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas3.ActiveConnection = Conexao
	
	
	rsMarcas3.Open SqlMarcas3, Conexao
	

While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2 FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"




Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros3.ActiveConnection = Conexao
	
	
	rsCarros3.Open SqlCarros3, Conexao
	
	





'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros3.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



rsMarcas3.close

set rsMarcas3 = nothing


rsCarros3.close

set rsCarros3 = nothing





End Function



%> 



<%
Function EscreveFuncaoJavaScript2 ( Conexao )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo3.options[doublecombo.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas4 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas4 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas4.ActiveConnection = Conexao
	
	
	rsMarcas4.Open SqlMarcas4, Conexao




While NOT rsMarcas4.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas4("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas4("id_combo1")&" order by nome_combo2"





Set rsCarros4 = Server.CreateObject("ADODB.RecordSet")

	rsCarros4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros4.ActiveConnection = Conexao
	
	
	rsCarros4.Open SqlCarros4, Conexao
	
	




'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros4.EoF

Response.Write "doublecombo.combo4.options[" & i & "] = new Option('" & rsCarros4("nome_combo2") & "','" & rsCarros4("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros4.MoveNext
Wend
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas4.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



rsMarcas4.close

set rsMarcas4 = nothing

rsCarros4.close

set rsCarros4 = nothing



End Function
%> 




  <%  EscreveFuncaoJavaScript ( Conexao ) %>
  <%  EscreveFuncaoJavaScript2 ( Conexao ) %>
</body>
</html>
