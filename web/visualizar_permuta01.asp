<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->




















<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if
%> 













































<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
  
	
	
	 dim Conexao9,rs9
 Set Conexao9 = Server.CreateObject("ADODB.Connection")
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	Conexao9.Open dsn
	dim strSQL9
	
	dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	 strSQL9 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais FROM permuta where cod_permuta="&varCodPermuta
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	 rs9.Open strSQL9, Conexao9
	
	
	
	
	 dim Conexao2,rs7

	Set rs7 = Server.CreateObject("ADODB.RecordSet")

	dim strSQL7
	
	if rs9("cod_imovel") <> "não informado" then
	 strSQL7 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where cod_imovel="&rs9("cod_imovel")
	 rs7.CursorLocation = 3
      rs7.CursorType = 3
	 rs7.Open strSQL7, Conexao9
   if not rs7.eof then
   vimagem = rs7("foto_grande")
   else
   vimagem = "imovel00000.jpg"
  end if
	
	else
	vimagem = "imovel00000.jpg"
	end if
	
	
	
	
	 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	
'----------------------------verifica imóvel---------------------------
	
	
	dim rsVerifica2
	dim strSQLVerifica2
	
 Set rsVerifica2 = Server.CreateObject("ADODB.RecordSet")
    
	
	
	
	
	
	strSQLVerifica2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where telefone='"&rs9("telefone")&"'"
	 
   
   
rsVerifica2.CursorLocation = 3
rsVerifica2.CursorType = 3

        rsVerifica2.Open strSQLVerifica2, Conexao9 	
		
		
	
	
	
	
	
	
	'-----------------------------------------------------------------	
		
		
%>		






<html>

<title>Visualizar permutante</title>
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

<!--#include file="style_imoveis02.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="#f7ecbf" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo" ENCTYPE="multipart/form-data" onSubmit="return isValidDigitNumber(this);" method="post" action="outFile001.asp">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></td>
  </tr>
  
  <tr>
      <td width="590" height="190"><table width="590" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="190">&nbsp;</td>
            <td width="580" height="190" style="border:1px solid #FFFFFF;"><table width="580" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="290" height="190" bgcolor="<%=claro%>" >&nbsp;</td>
                  <td width="290" height="190" ><table width="290" height="190" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="290" height="170"><%If objFSO.FileExists(Server.MapPath(vimagem)) = True Then%><img src="<%=vimagem%>" width="290" height="170"></img><% else %><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div><% end if %></td>
                      </tr>
                      <tr>
                        <td width="290" height="20" bgcolor="<%=claro%>" >
						
						 <%if not rsVerifica2.eof then %>
        <div align="center"><a href="javascript:newWindow22('mostrar_imovel2.asp?varCodimovel=<%=rsVerifica2("cod_imovel")%>')" style="color:#FFFFFF"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Clique 
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
	  <br>
	  <br>
<div align="center"> <font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          <%
		  if session("telefone") = "43621135" then
		  
		  response.write "O telefone desse permutante é "&rs9("telefone")
		   end if
		  %>
          </strong></font> </div>
        <br><br></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
               <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Última atualização</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("data_atualizacao")%></font></td>
              </tr>
			   
			   
			   
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      da permuta</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("cod_permuta")%></font></td>
              </tr>
			 
			 
			 
			 
			    <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      nome &eacute;:</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;">
<input name="txt_nome" value="<%=rs9("nome")%>" type="text" id="txt_nome" size="38" maxlength="200" align="left" class="inputBox" style="border-color:#f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background: #f7ecbf; "></td>
              </tr>
              
			 
              
			  
			  
             
			  
			  
			  
                
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atualmente 
                      tenho um im&oacute;vel na cidade de:</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("cidade_vend") = "cqualquer" then response.write "não informado" else response.write rs9("cidade_vend") end if %>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      bairro: </font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("bairro_vend") = "bqualquer" then response.write "não informado" else response.write rs9("bairro_vend") end if %>
                    </font></td>
              </tr>
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Na 
                      vila: </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("vila_vend") = "vlqualquer" then response.write "não informado" else response.write rs9("vila_vend") end if %>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">do 
                      tipo: </font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("tipo_vend") = "tqualquer" then response.write "não informado" else response.write rs9("tipo_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("quartos_vend") = "qqualquer" then response.write "não informado" else response.write rs9("quartos_vend") end if %>
                    </font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de vagas na garagem</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("vagas_vend") = "vgqualquer" then response.write "não informado" else response.write rs9("vagas_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      valor de</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("valor_vend") = "vqualquer" then response.write "não informado" else response.write FormatNumber(rs9("valor_vend"),2) end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      im&oacute;vel tem a seguinte descri&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> <textarea name="textarea" class="inputBox" id="textarea"  style="border-color:#f7ecbf;color:#9d9249;HEIGHT: 100px; WIDTH: 290px; background:#f7ecbf; " onKeyPress="return limitfield(this, 800)"><%=rs9("descricao_vend")%></textarea></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Pretendo 
                      morar na cidade de:</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("cidade_comp") = "cqualquer" then response.write "não informado" else  response.write rs9("cidade_comp") end if %>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      bairro: </font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("bairro_comp") = "bqualquer" then response.write "não informado" else  response.write rs9("bairro_comp") end if %>
                    </font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Na 
                     vila: </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("vila_comp") = "vlqualquer" then response.write "não informado" else  response.write rs9("vila_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
                <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quero 
                      trocar por um im&oacute;vel do tipo:</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp; 
                    </font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("tipo_comp") = "tqualquer" then response.write "não informado" else  response.write rs9("tipo_comp") end if %>
                    </font></td>
              </tr>
                
				
				 
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("quartos_comp") = "qqualquer" then response.write "não informado" else  response.write rs9("quartos_comp") end if %>
                    </font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de vagas na garagem</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("vagas_comp") = "vgqualquer" then response.write "não informado" else  response.write rs9("vagas_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      valor de</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#9d9249">&nbsp;</font><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("valor_comp") = "vqualquer" then response.write "não informado" else  response.write FormatNumber(rs9("valor_comp"),2) end if %>
                    </font></td>
              </tr>
				
				
				
              <tr>
                  <td width="290" bgcolor="#f7ecbf" height="100" style="border:1px solid #FFFFFF;" >
<div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Que 
                      tenha a seguinte descri&ccedil;&atilde;o</font></div></td>
                  <td width="290" height="100" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao2" class="inputBox" id="txt_descricao2"  style="border-color:#f7ecbf;color:#9d9249;HEIGHT: 100px; WIDTH: 290px; background:#f7ecbf; " onKeyPress="return limitfield(this, 800)"><%=rs9("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>&nbsp;</td>
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
</form>

<%
           rs9.Close
           'fecha a conexão
       
           Set rs9 = Nothing
		   
		   
		   rs7.close
		   
		   set rs7 = nothing
		   
		   rsVerifica2.close
		   
		   set rsVerifica2 = nothing
		   
		   set objfso = nothing
		   
		   conexao9.close
		   
		   set conexao9 = nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>
