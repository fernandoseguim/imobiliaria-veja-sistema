<!--#include file="dsn.asp"-->





















<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!

%> 






























<%
'Criando conexão com o banco de dados! 



%> 














<!--#include file="cores.asp"-->
<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
  
		
	
	
	
	 dim Conexao9,rs9

	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	
	dim strSQL9
	
	dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	 strSQL9 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais FROM permuta where cod_permuta="&varCodPermuta
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	 rs9.Open strSQL9, Conexao3
	
	
	
	
	
	
	'----------------------------verifica imóvel---------------------------
	
	
	dim rsVerifica2
	dim strSQLVerifica2
	
 Set rsVerifica2 = Server.CreateObject("ADODB.RecordSet")
    
	
	
	
	
	
	strSQLVerifica2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou FROM imoveis where telefone='"&rs9("telefone")&"'"
	 
   
   
rsVerifica2.CursorLocation = 3
rsVerifica2.CursorType = 3

        rsVerifica2.Open strSQLVerifica2, Conexao3 	
		
		
	
	
	
	
	
	
	'-----------------------------------------------------------------	
		
		
%>		






<html>

<title>Indicações</title>
<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=claro%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=605,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>


</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="imprimir_display_permuta22.asp?varCodPermuta=<%=varCodPermuta%>" style="color:#FFFFFF">Visualizar 
        impressão</a></strong></font></div></td>
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
                        <td width="290" height="170"> <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div></td>
                      </tr>
                      <tr>
                        <td width="290" height="20" bgcolor="<%=claro%>" >
						
						
						
						 <%if not rsVerifica2.eof then %>
        <div align="center"><a href="javascript:newWindow22('mostrar_imovel2.asp?varCod_imovel=<%=rsVerifica2("cod_imovel")%>')" style="color:#FFFFFF"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Clique 
          aqui e veja o meu imóvel</strong></font></a> 
          <% else %>
          <%end if%>
        </div>
						
						
						</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
            <td width="5" height="190">&nbsp;</td>
          </tr>
        </table></td>
  </tr>
  
  
  
  <tr>
      <td height="18">
<div align="center"> 
          <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi incluido com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      da permuta</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("cod_permuta")%></font></td>
              </tr>
			 
			    <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo atendimento</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("atendimento")%></font></td>
              </tr>
			 
			 
			 
			    <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do permutante</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("nome")%></font></td>
              </tr>
			   <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      do permutante</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6" then %><% if  UCase(rs9("atendimento")) <> UCase(Session("Admin_ID")) then response.write "Não informado" else response.write rs9("telefone") end if %><%else%><%response.write rs9("telefone") end if %></font></td>
              </tr>
              
			 
              
			  
			  
             
			  
			  
			  
                
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      do im&oacute;vel atual</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("cidade_vend") = "cqualquer" then response.write "não informado" else response.write rs9("cidade_vend") end if %>
                    </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      do im&oacute;vel atual</font></div></td>

                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("bairro_vend") = "bqualquer" then response.write "não informado" else response.write rs9("bairro_vend") end if %>
                    </font></td>
              </tr>
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      do im&oacute;vel atual</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("vila_vend") = "vlqualquer" then response.write "não informado" else response.write rs9("vila_vend") end if %>
                    </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel atual</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("tipo_vend") = "tqualquer" then response.write "não informado" else response.write rs9("tipo_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dormit&oacute;rios 
                      do im&oacute;vel atual</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("quartos_vend") = "qqualquer" then response.write "não informado" else response.write rs9("quartos_vend") end if %>
                    </font></td>
              </tr>
			  
			  <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      do im&oacute;vel atual</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("vagas_vend") = "vgqualquer" then response.write "não informado" else response.write rs9("vagas_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do im&oacute;vel atual</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("valor_vend") = "vqualquer" then response.write "não informado" else response.write FormatNumber(rs9("valor_vend"),2) end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                      do im&oacute;vel atual</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="textarea" class="inputBox" id="textarea"  style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 400)"><%=rs9("descricao_vend")%></textarea></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      desejada </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("cidade_comp") = "cqualquer" then response.write "não informado" else  response.write rs9("cidade_comp") end if %>
                    </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      desejado </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("bairro_comp") = "bqualquer" then response.write "não informado" else  response.write rs9("bairro_comp") end if %>
                    </font></td>
              </tr>
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      desejada </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("vila_comp") = "vlqualquer" then response.write "não informado" else  response.write rs9("vila_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel pretendido</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp; 
                    </font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("tipo_comp") = "tqualquer" then response.write "não informado" else  response.write rs9("tipo_comp") end if %>
                    </font></td>
              </tr>
                
				
				 
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                      de quartos do im&oacute;vel pretendido</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("quartos_comp") = "qqualquer" then response.write "não informado" else  response.write rs9("quartos_comp") end if %>
                    </font></td>
              </tr>
			  
			  <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                      de vagas do im&oacute;vel pretendido</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("vagas_comp") = "vgqualquer" then response.write "não informado" else  response.write rs9("vagas_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do im&oacute;vel pretendido</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("valor_comp") = "vqualquer" then response.write "não informado" else  response.write FormatNumber(rs9("valor_comp"),2) end if %>
                    </font></td>
              </tr>
				
				<tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Standby</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("standby") <> "" then response.write rs9("standby") else  response.write "excluído" end if %>
                    </font></td>
              </tr>
				
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                            do im&oacute;vel pretendido</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao2" class="inputBox" id="txt_descricao2"  style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; " onKeyPress="return limitfield(this, 400)"><%=rs9("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="145"><a href="indicacao_permuta22.asp?varIndicacaoCidadeVend=<%=session("varIndicacaoCidadeVend")%>&varIndicacaoBairroVend=<%=session("varIndicacaoBairroVend")%>&varIndicacaoVilaVend=<%=session("varIndicacaoVilaVend")%>&varIndicacaoTipoVend=<%=session("varIndicacaoTipoVend")%>&varIndicacaoQuartosVend=<%=session("varIndicacaoQuartosVend")%>&varIndicacaoVagasVend=<%=session("varIndicacaoVagasVend")%>&varIndicacaoValorVend=<%=session("varIndicacaoValorVend")%>&varIndicacaoCidadeComp=<%=session("varIndicacaoCidadeComp")%>&varIndicacaoBairroComp=<%=session("varIndicacaoBairroComp")%>&varIndicacaoVilaComp=<%=session("varIndicacaoVilaComp")%>&varIndicacaoTipoComp=<%=session("varIndicacaoTipoComp")%>&varIndicacaoQuartosComp=<%=session("varIndicacaoQuartosComp")%>&varIndicacaoVagasComp=<%=session("varIndicacaoVagasComp")%>&varIndicacaoValorComp=<%=session("varIndicacaoValorComp")%>&varIndicacaoCodigo=<%=session("varIndicacaoCodigo")%>"><img  src="bt_voltar001.jpg" width="148" height="18" border="0"></a></td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>

</center>
<%
           rs9.Close
           'fecha a conexão
           
		   
           Set rs9 = Nothing
		   
		   
		   
		   rsVerifica2.close
		   
		   set rsVerifica2 = nothing
		   
		   
		   conexao3.close
		   
		   set conexao3 = nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>
