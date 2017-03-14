




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
	
	
		 dim EnderecoIP
	 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
	 
	 dim PropostaFeita
	 
	 PropostaFeita = request.querystring("PropostaFeita")
	 
	 if PropostaFeita = "" then
	 
	 
	Conexao.execute"Insert into imovel_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao) values( '"& session("nome") &"','"& session("telefone") &"','"& rs("cod_imovel") &"','"& EnderecoIP &"','"& now() &"','"& rs("tipo") &"','"& rs("quartos") &"','"& rs("vagas") &"','"& rs("cidade") &"','"& rs("bairro") &"','"& rs("valor") &"','"& rs("negociacao") &"')"
	
	end if
		
          
 %>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<!--#include file="style_imoveis.asp"-->
<body bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo"  method="post" action="incluir_querpagar.asp?varCodImovel=<%=varCodImovel%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  
  <tr>
    <td width="590" height="334"><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td width="580" height="334" style="border:1px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                    <div align="center"><img src="<%=rs("foto_grande")%>" name="photoslider" width="580" height="334"></img></div>
                      <% else %>
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
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
  
  <tr>
 
    <td height="80"><table width="550" border="0" align="center">
        <tr>
          <td width="340"><table width="340" border="0">
              <tr>
                <td><div align="right"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Voc&ecirc; 
                    quer pagar quanto por esse im&oacute;vel?</strong></font></div></td>
              </tr>
              <tr>
                <td><div align="right"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Qual 
                    forma de pagamento voc&ecirc; prop&otilde;e?</strong></font></div></td>
              </tr>
            </table></td>
          <td width="130"><table width="130" border="0">
              <tr>
                <td height="18" style="border:1px solid #FFFFFF;" bgcolor="<%=claro%>"><input name="txt_valor" value="0,00" type="text" id="txt_valor" size="38" maxlength="21" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 130px; background:<%=claro%>"></td>
              </tr>
              <tr>
                <td height="18" style="border:1px solid #FFFFFF;" bgcolor="<%=claro%>"><select name="txt_pagamento" size="1" class="inputBox" id="txt_pagamento" style="HEIGHT: 22px; WIDTH: 130px; background:<%=claro%>">
                      <option value="à vista" selected>à vista</option>
					<option value="finaciamento">Financiamento</option>
                    <option value="parcelado">parcelado</option>
                  </select></td>
              </tr>
            </table></td>
          <td width="80"><input name="image" type="image"  src="bt_enviar004.jpg" width="80" height="48" border="0"></td>
        </tr>
      </table></td>
  </tr>
 
  
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td width="579"><table bgcolor="<%=medio%>" width="579" border="0" cellspacing="0" cellpadding="0" >
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cidade")%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("bairro")%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("vila")%></font></div></td>
                                </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("tipo")%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                          Total </font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("area_total")%> m&sup2;
                          </font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                          Construida </font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("area_construida")%> m&sup2;
                          </font></div></td>
                                </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("quartos")%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("banheiros")%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("vagas")%></font></div></td>
                                </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("negociacao")%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormatNumber(rs("valor"),2)%></font></div></td>
                                </tr>
                  </table></td>
                <td width="193" height="60"><table width="180" bgcolor="<%=claro%>" height="47" border="0" align="center" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
                    <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Condom&iacute;nio</font></div></td>
                                </tr>
                                <tr> 
                                  <td><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if rs("condominio") <> "" then response.write FormatNumber(rs("condominio"),2) else response.write "0,00" end if %></font></div></td>
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
          <td width="580" height="140" bgcolor="<%=medio%>">
<center><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("obs_imovel")%> <br>
              <br>
              <b>Código de referência <%=rs("cod_imovel")%></b></font></center></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
                <td width="290"><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="145"><a href="imovel_proposta.asp?varCodImovel=<%=varCodImovel%>"><img src="bt_proposta001.jpg" width="148" height="18" border="0"></a></td>
                      <td width="145" height="18"><img src="bt_indicar001.jpg" width="148" height="18"></img></td>
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
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   Set objFSO = Nothing
           %>
  <% response.flush%>
  <%response.clear%>
</body>
</html>
