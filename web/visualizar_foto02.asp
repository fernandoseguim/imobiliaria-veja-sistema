<!--#include file="dsn.asp"-->

<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn




%> 










<!--#include file="cores02.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel,objFSO
Dim rs2,strSQL2,varCodImovel


dim varNumFoto


dim varNomeFoto
varCodImovel = request.QueryString("varCodimovel")
'varCodImovel = "3358"



varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
   Set rs2 = Server.CreateObject("ADODB.RecordSet")
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
	 
	 strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.data_foto_colocada,imoveis.quem_foto_colocada  FROM imoveis where cod_imovel="&varCodImovel
	 
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

RS2.CursorLocation = 3
RS2.CursorType = 3

        rs.Open strSQL, Conexao3 
		
		
	
		
%>		





<html>
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

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>

</head>
<!--#include file="style_imoveis.asp"-->





<title>Visualizar fotos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<!--#include file="style_imoveis.asp"-->
<body bgcolor="#f7ecbf" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><a href="visualizar_foto02.asp?varCodimovel=<%=varCodImovel%>"><img src="top_resultado02.jpg" width="590" height="48" border="0"></a></td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td height="30"> 
      <% if rs("data_foto_colocada") <> "" then %>
      <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
    <%else%>
	
	<%end if%>
 
 
  </tr>
  
  
   
  <tr>
    <td height="150"><table width="590" height="150" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" height="150" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                          1</strong></font></div></td>
                    </tr>
                    <tr>
                      <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rs("foto_grande1"))) = True Then%><img src="<%=rs("foto_grande1")%>" width="187" height="107" border="0"></img><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=1')" style="color:#9d9249"><strong>Foto não disponível</strong></a></font></div><%end if%></td>
                    </tr>
                    <tr>
                      <td width="187"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=1')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                          2</strong></font></div></td>
                    </tr>
                    <tr>
                      <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <% If objFSO.FileExists(Server.MapPath(rs("foto_grande2"))) = True Then%>
                        <img src="<%=rs("foto_grande2")%>" width="187" height="107" border="0"></img> 
                        <%else%>
                        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=2')" style="color:#9d9249"><strong>Foto 
                          não disponível</strong></a></font></div>
                        <%end if%>
                      </td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=2')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                          3</strong></font></div></td>
                    </tr>
                    <tr>
                      <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <% If objFSO.FileExists(Server.MapPath(rs("foto_grande3"))) = True Then%>
                        <img src="<%=rs("foto_grande3")%>" width="187" height="107" border="0"></img> 
                        <%else%>
                        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=3')" style="color:#9d9249"><strong>Foto 
                          não disponível</strong></a></font></div>
                        <%end if%>
                      </td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=3')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="150"><table width="590" height="150" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" height="150" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                          4</strong></font></div></td>
                    </tr>
                    <tr>
                      <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <% If objFSO.FileExists(Server.MapPath(rs("foto_grande4"))) = True Then%>
                        <img src="<%=rs("foto_grande4")%>" width="187" height="107" border="0"></img> 
                        <%else%>
                        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=4')" style="color:#9d9249"><strong>Foto 
                          não disponível</strong></a></font></div>
                        <%end if%>
                      </td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=4')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                          5</strong></font></div></td>
                    </tr>
                    <tr>
                      <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <% If objFSO.FileExists(Server.MapPath(rs("foto_grande5"))) = True Then%>
                        <img src="<%=rs("foto_grande5")%>" width="187" height="107" border="0"></img> 
                        <%else%>
                        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=5')" style="color:#9d9249"><strong>Foto 
                          não disponível</strong></a></font></div>
                        <%end if%>
                      </td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=5')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                          6</strong></font></div></td>
                    </tr>
                    <tr>
                      <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                        <% If objFSO.FileExists(Server.MapPath(rs("foto_grande6"))) = True Then%>
                        <img src="<%=rs("foto_grande6")%>" width="187" height="107" border="0"></img> 
                        <%else%>
                        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=6')" style="color:#9d9249"><strong>Foto 
                          não disponível</strong></a></font></div>
                        <%end if%>
                      </td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=6')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="150"><table width="580" height="150" border="0" cellpadding="0" cellspacing="0">
              <tr>
			   <td width="5">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
              <tr>
                      
                <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                    7</strong></font></div></td>
                    </tr>
                    <tr>
                      
                <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <% If objFSO.FileExists(Server.MapPath(rs("foto_grande7"))) = True Then%>
                  <img src="<%=rs("foto_grande7")%>" width="187" height="107" border="0"></img> 
                  <%else%>
                  <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=7')" style="color:#9d9249"><strong>Foto 
                    não disponível</strong></a></font></div>
                  <%end if%>
                </td>
                    </tr>
                    <tr>
                      
                <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=7')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
              <tr>
                      
                <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                    8</strong></font></div></td>
                    </tr>
                    <tr>
                      
                <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <% If objFSO.FileExists(Server.MapPath(rs("foto_grande8"))) = True Then%>
                  <img src="<%=rs("foto_grande8")%>" width="187" height="107" border="0"></img> 
                  <%else%>
                  <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=8')" style="color:#9d9249"><strong>Foto 
                    não disponível</strong></a></font></div>
                  <%end if%></td>
                    </tr>
                    <tr>
                      
                <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=8')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187"><table width="187" border="0" cellspacing="0" cellpadding="0">
              <tr>
                      
                <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                    9</strong></font></div></td>
                    </tr>
                    <tr>
                      
                <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <% If objFSO.FileExists(Server.MapPath(rs("foto_grande9"))) = True Then%>
                  <img src="<%=rs("foto_grande9")%>" width="187" height="107" border="0"></img> 
                  <%else%>
                  <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=9')" style="color:#9d9249"><strong>Foto 
                    não disponível</strong></a></font></div>
                  <%end if%>
                </td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=9')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td><table width="580" height="150" border="0" cellpadding="0" cellspacing="0">
              <tr>
			   <td width="5">&nbsp;</td>
                <td width="187"><table width="187" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                      
                <td height="20" bgcolor="#9d9249" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                    10</strong></font></div></td>
                    </tr>
                    <tr>
                      
                <td height="110" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <% If objFSO.FileExists(Server.MapPath(rs("foto_grande10"))) = True Then%>
                  <img src="<%=rs("foto_grande10")%>" width="187" height="107" border="0"></img> 
                  <%else%>
                  <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=10')" style="color:#9d9249"><strong>Foto 
                    não disponível</strong></a></font></div>
                  <%end if%></td>
                    </tr>
                    <tr>
                      <td width="223"><div align="center"><a href="javascript:newWindow3('form_atualizar_foto02.asp?varCodimovel=<%=varCodimovel%>&varNumFoto=10')" style="color:#9d9249"><img src="bt_incluir004.jpg" width="187" height="20" border="0"></a></div></td>
                    </tr>
                  </table></td>
                <td width="10">&nbsp;</td>
                <td width="187">&nbsp;</td>
                <td width="10">&nbsp;</td>
                <td width="187">&nbsp;</td>
              </tr>
            </table></td>
  </tr>
  
  </tr>
</table>

 <%
           rs.Close
           set rs = nothing
		   
		  
		   set rs2 = nothing
           
		   Set objFSO = Nothing
		   
		   conexao.close
		   
		   set conexao = nothing
         
           %>
  <% response.flush%>
  <%response.clear%>




</body>



</html>

