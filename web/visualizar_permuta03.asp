<!--#include file="dsn.asp"-->
<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas3 = Conexao3.Execute ( SqlMarcas3 )

While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"
Set rsCarros3 = Conexao3.Execute ( SqlCarros3 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
While NOT rsCarros3.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend

'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Sql5 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs5 = Conexao3.Execute ( Sql5 )
Set Rs3 = Conexao3.Execute ( Sql3 ) 
%> 










<%
Function EscreveFuncaoJavaScript2 ( Conexao4 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo3.options[form.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas4 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas4 = Conexao4.Execute ( SqlMarcas4 )

While NOT rsMarcas4.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas4("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros4 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas4("id_combo1")&" order by nome_combo2"
Set rsCarros4 = Conexao4.Execute ( SqlCarros4 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
While NOT rsCarros4.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros4("nome_combo2") & "','" & rsCarros4("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros4.MoveNext
Wend

'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas4.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao4 = Server.CreateObject("ADODB.Connection")
Conexao4.Open dsn

'Abrindo a tabela MARCAS!
Sql4 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs5 = Conexao4.Execute ( Sql4 ) 
%> 














<!--#include file="cores.asp"-->
<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
   dim rs4,strSQL4,strSQL6,rs6
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	Set rs6 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	strSQL6 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		rs4.Open strSQL4, Conexao
		rs6.Open strSQL6, Conexao
		
	
	
	
	 dim Conexao9,rs9
 Set Conexao9 = Server.CreateObject("ADODB.Connection")
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	Conexao9.Open dsn
	dim strSQL9
	
	dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	 strSQL9 = "SELECT * FROM permuta where cod_permuta="&varCodPermuta
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	 rs9.Open strSQL9, Conexao9
	
	
	
	
	 dim Conexao2,rs7
 Set Conexao2 = Server.CreateObject("ADODB.Connection")
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	Conexao2.Open dsn
	dim strSQL7
	
	if rs9("cod_imovel") <> "não informado" then
	 strSQL7 = "SELECT * FROM imoveis where cod_imovel="&rs9("cod_imovel")
	 rs7.CursorLocation = 3
      rs7.CursorType = 3
	 rs7.Open strSQL7, Conexao2
   if not rs7.eof then
   vimagem = rs7("foto_grande")
   else
   vimagem = "imovel00000.jpg"
  end if
	
	else
	vimagem = "imovel00000.jpg"
	end if
	
	
	
	
	 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	
		
%>		






<html>


<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript2 ( Conexao3 ) %>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=claro%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<center>
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48">&nbsp;</td>
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
                        <td width="290" height="170"><%If objFSO.FileExists(Server.MapPath(vimagem)) = True Then%><img src="<%=vimagem%>" width="290" height="170"></img><% else %> <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div><% end if %></td>
                      </tr>
                      <tr>
                        <td width="290" height="20" bgcolor="<%=claro%>" >
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                            do meu im&oacute;vel</font></div></td>
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      da permuta</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("cod_permuta")%></font></td>
              </tr>
			 
			  <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo atendimento</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("atendimento")%></font></td>
              </tr>
			 
			 
			 
			    <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      nome &eacute;:</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("nome")%></font></td>
              </tr>
              
			 
              
			  
			  
             
			  
			  
			  
                
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atualmente 
                      tenho um im&oacute;vel na cidade de:</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("cidade_vend") = "cqualquer" then response.write "não informado" else response.write rs9("cidade_vend") end if %>
                    </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      bairro: </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("bairro_vend") = "bqualquer" then response.write "não informado" else response.write rs9("bairro_vend") end if %>
                    </font></td>
              </tr>
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Na 
                      vila: </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("vila_vend") = "vlqualquer" then response.write "não informado" else response.write rs9("vila_vend") end if %>
                    </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">do 
                      tipo: </font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("tipo_vend") = "tqualquer" then response.write "não informado" else response.write rs9("tipo_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("quartos_vend") = "qqualquer" then response.write "não informado" else response.write rs9("quartos_vend") end if %>
                    </font></td>
              </tr>
			  
			  <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de vagas na garagem</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("vagas_vend") = "vgqualquer" then response.write "não informado" else response.write rs9("vagas_vend") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      valor de</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("valor_vend") = "vqualquer" then response.write "não informado" else response.write FormatNumber(rs9("valor_vend"),2) end if %>
                    </font></td>
              </tr>
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      im&oacute;vel tem a seguinte descri&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="textarea" class="inputBox" id="textarea"  style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 400)"><%=rs9("descricao_vend")%></textarea></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Pretendo 
                      morar na cidade de:</font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("cidade_comp") = "cqualquer" then response.write "não informado" else  response.write rs9("cidade_comp") end if %>
                    </font></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      bairro: </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%if rs9("bairro_comp") = "bqualquer" then response.write "não informado" else  response.write rs9("bairro_comp") end if %>
                    </font></td>
              </tr>
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Na 
                     vila: </font></div></td>
                  <td style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("vila_comp") = "vlqualquer" then response.write "não informado" else  response.write rs9("vila_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
                <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quero 
                      trocar por um im&oacute;vel do tipo:</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp; 
                    </font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("tipo_comp") = "tqualquer" then response.write "não informado" else  response.write rs9("tipo_comp") end if %>
                    </font></td>
              </tr>
                
				
				 
                <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("quartos_comp") = "qqualquer" then response.write "não informado" else  response.write rs9("quartos_comp") end if %>
                    </font></td>
              </tr>
			  
			  <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Com 
                      o seguinte n&uacute;mero de vagas</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("vagas_comp") = "vgqualquer" then response.write "não informado" else  response.write rs9("vagas_comp") end if %>
                    </font></td>
              </tr>
			  
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">No 
                      valor de</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF">&nbsp;</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%if rs9("valor_comp") = "vqualquer" then response.write "não informado" else  response.write FormatNumber(rs9("valor_comp"),2) end if %>
                    </font></td>
              </tr>
				
				
				
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=medio%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Que 
                            tenha a seguinte descri&ccedil;&atilde;o</font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=claro%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="7B9AB9" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao2" class="inputBox" id="txt_descricao2"  style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 400)"><%=rs9("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="145"><input name="image" type="image" onClick="window.history.go(-1);"  src="bt_voltar001.jpg" width="148" height="18"></td>
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
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>
