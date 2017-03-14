





<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCodImovel,objFSO
varCodImovel = request.QueryString("varCodImovel")
   
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis Where cod_imovel = "&varCodImovel
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
		
		
		
	if not(rs.eof) and not(rs.bof) or (rs.recordcount >= 6) then
		
          
 %>





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

	<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: #DF700F;}
</STYLE>
<script language="javascript">
function funScroll()
{
window.scrollTo(0,128)

}		
</script>



</head>


<body onLoad="funScroll()" bgcolor="<%=escuro%>" topmargin="0" bottommargin="0" rightmargin="0" leftmargin="0" marginheight="0" marginwidth="0">
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="590" height="48">&nbsp;</td>
  </tr>
  <tr> 
    <td height="5"></td>
  </tr>
  <tr> 
    <td width="590" height="435"><table width="590" height="435" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="5" height="435">&nbsp;</td>
          <td width="580" height="435"><table width="580" height="435" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="175" height="435"><table width="175" height="435" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cidade")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
					
					<tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("bairro")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
					
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("vila")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("tipo") = "terreno" then response.write "Terreno/Área" else response.write rs("tipo") end if%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                                      Total </font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                      <%if rs("area_total") = "não informado" then%>
                                      <%=rs("area_total")%></font>
                                      <%else%>
                                      <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("area_total")%>m&sup2;</font>
                                      <% end if %>
                                    </div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                                      Constru&iacute;da</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                      <%if rs("area_construida") = "não informado" then%>
                                      <%=rs("area_construida")%></font>
                                      <%else%>
                                      <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("area_construida")%> 
                                      m&sup2;</font>
                                      <% end if %>
                                    </div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("quartos")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("banheiros")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                                      na Garagem</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("vagas")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("negociacao")%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><table width="175" height="40" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><table width="175" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormatNumber(rs("valor"),2)%></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
                <td width="405" height="435"><table width="405" height="435" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="3" height="435">&nbsp;</td>
                      <td width="402" height="435"><table width="402" height="435" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="402" height="235" valign="top"><table width="402" height="232" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
								<script language="JavaScript">
                         var photos=new Array()
                         var which=0
                         
photos[0]="<%=rs("foto_grande1")%>"
photos[1]="<%=rs("foto_grande2")%>"
photos[2]="<%=rs("foto_grande3")%>"
photos[3]="<%=rs("foto_grande4")%>"
photos[4]="<%=rs("foto_grande5")%>"
 var tam = 3;
<% if rs("foto_grande")<>"imovel00000.jpg" and rs("foto_grande1")<>"imovel00000.jpg" and rs("foto_grande2")<>"imovel00000.jpg" and rs("foto_grande3")<>"imovel00000.jpg" and rs("foto_grande4")<>"imovel00000.jpg" and rs("foto_grande5")<>"imovel00000.jpg" then%>
                         var tam = 4;
						<%end if%>

<% if rs("foto_grande")<>"imovel00000.jpg" and rs("foto_grande1")<>"imovel00000.jpg" and rs("foto_grande2")<>"imovel00000.jpg" and rs("foto_grande3")<>"imovel00000.jpg" and rs("foto_grande4")="imovel00000.jpg" and rs("foto_grande5")="imovel00000.jpg" then %>
                         var tam = 3;
						<%end if%>
						
						<% if rs("foto_grande")<>"imovel00000.jpg" and rs("foto_grande1")<>"imovel00000.jpg" and rs("foto_grande2")<>"imovel00000.jpg" and rs("foto_grande3")<>"imovel00000.jpg" and rs("foto_grande4")="imovel00000.jpg"  and rs("foto_grande5")="imovel00000.jpg"  then%>
                         var tam = 2;
						<%end if%>					 
                       
					   <% if rs("foto_grande")<>"imovel00000.jpg" and rs("foto_grande1")<>"imovel00000.jpg" and rs("foto_grande2")<>"imovel00000.jpg" and rs("foto_grande3")="imovel00000.jpg" and rs("foto_grande4")="imovel00000.jpg" and rs("foto_grande5")="imovel00000.jpg" then %>
                         var tam = 1;
						<%end if%>
						
						<% if rs("foto_grande")<>"imovel00000.jpg" and rs("foto_grande1")<>"imovel00000.jpg" and rs("foto_grande2")="imovel00000.jpg" and rs("foto_grande3")="imovel00000.jpg" and rs("foto_grande4")="imovel00000.jpg" and rs("foto_grande5")="imovel00000.jpg" then %>
                         var tam = 0;
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
                                  <td width="402" height="232" style="border:1px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                    <div align="center"><img src="<%=rs("foto_grande")%>" name="photoslider" width="400" height="230"></img></div>
                      <% else %>
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div>
                    <% end if %></td>
                                </tr>
                              </table></td>
                          </tr>
                          
                            <td width="402" height="200"><table width="401" height="200" border="0" cellpadding="0" cellspacing="0">
                                
                                  <td width="402" height="200" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><table width="402" border="0" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td width="402" height="18" bgcolor="<%=claro%>"><table width="402" border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:anterior()" class="link" onmouseover="window.status='Anterior'; return true" onmouseout="window.status=''"><img src="bt_anterior001.jpg" width="201" height="18" border="0"></a><%else%><%end if%></td>
                                              <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:proxima()" class="link" onmouseover="window.status='Próxima'; return true" onmouseout="window.status=''"><img src="bt_proxima001.jpg" width="201" height="18" border="0"></a><%else%><%end if%></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                      <tr>
                                        <td width="402" height="182" valign="middle"><center><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("obs_imovel")%> <br><br><b>Código de referência <%=rs("cod_imovel")%></b></font></center></td>
                                      </tr>
                                    </table> </td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="5" height="435">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="289">&nbsp;</td>
          <td width="148">&nbsp;</td>
          <td width="148"><input type="image" onClick="window.history.go(-1);"  src="bt_voltar001.jpg" width="148" height="18"></img></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>


<% else %>


<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">   Não foi encontrado o imóvel pedido!!</font>

<% end if %>

 <%
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   Set objFSO = Nothing
           %>
  <% response.flush%>
  <%response.clear%>
</center>
</body>
</html>
