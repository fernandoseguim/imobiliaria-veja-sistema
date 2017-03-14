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
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
While NOT rsCarros3.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "form.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
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
Conexao3.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 




%> 





<%
Function EscreveFuncaoJavaScript222 ( Conexao333 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros222 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo2.options[doublecombo.combo2.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas333 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas333 = Conexao333.Execute ( SqlMarcas333 )

While NOT rsMarcas333.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas333("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo5.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros333 = "SELECT * FROM combo3 where id_combo2 =" & rsMarcas333("id_combo2")&""

Set rsCarros333 = Conexao333.Execute ( SqlCarros333 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1  
While NOT rsCarros333.EoF

Response.Write "doublecombo.combo5.options[" & i & "] = new Option('" & rsCarros333("nome_combo3") & "','" & rsCarros333("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros333.MoveNext
Wend


Response.Write "doublecombo.combo5.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas333.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 


<%

'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'

Sql333 = "SELECT * FROM combo2 ORDER BY nome_combo2" 
Set Rs333 = Conexao333.Execute ( Sql333 ) 




%> 









<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->


<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel,objFSO
Dim rs2,strSQL2,varCodImovel

varCodImovel = request.QueryString("varCod_imovel")
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
   Set rs2 = Server.CreateObject("ADODB.RecordSet")
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL2 = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 strSQL = "SELECT * FROM imoveis where cod_imovel="&varCod_imovel
	 
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

RS2.CursorLocation = 3
RS2.CursorType = 3

        rs.Open strSQL, Conexao 
		rs2.Open strSQL2, Conexao
		
	
	dim Sql4,rs4
	  Set rs4 = Server.CreateObject("ADODB.RecordSet")
Sql4 = "SELECT * FROM combo2 where nome_combo2 like '"& rs("bairro") &"' and cidade_combo2 like '"& rs("cidade") &"' ORDER BY nome_combo2" 
Set Rs4 = Conexao.Execute ( Sql4 ) 	




dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")
	strSQL444 = "SELECT * FROM combo3 where nome_combo3 ='"&rs("vila")&"' and bairro_combo3 ='"&rs("bairro")&"' and cidade_combo3 ='"&rs("cidade")&"'   ORDER BY nome_combo3" 
	 rs444.Open strSQL444, Conexao		





dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT * FROM combo1 where nome_combo1 ='"&rs("cidade")&"'  ORDER BY nome_combo1" 
	 rs555.Open strSQL555, Conexao		






%>		





<html>
<head><%  EscreveFuncaoJavaScript ( Conexao3 ) %>
<%  EscreveFuncaoJavaScript222 ( Conexao333 ) %>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

<script>
function isValidDigitNumber (doublecombo)
{


	
	if (doublecombo.combo2.value == "") {
        alert("O formulário Bairro do Imóvel está vazio!");
        doublecombo.combo1.focus();
		
        return false;
    }
	
	if (doublecombo.combo1.value == "") {
        alert("O formulário Cidade do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }





var elem=doublecombo.elements;





for (nCount=0; nCount < elem.length; nCount++)
  
    
  
	
	if(elem[nCount].type.indexOf("text")==0)	{
	var strValidNumber12_1="'";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)!=-1) {



alert("Este campo  não pode conter aspas");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}
//-----------------------------------------------

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






<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<!--#include file="style_imoveis.asp"-->
<body bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_imovel.asp?varCod_imovel=<%=varCod_imovel%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="590" height="48"><a href="visualizar_imovel.asp?varCod_imovel=<%=varCodImovel%>"><img src="top_resultado.jpg" width="590" height="48" border="0"></a></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
            <td><table width="580" border="0" cellspacing="0" cellpadding="0" style="border:1px solid #FFFFFF;">
              <tr>
                <td width="580" height="334" bgcolor="<%=escuro%>"><% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                    <div align="center"><img src="<%=rs("foto_grande")%>" name="photoslider" width="580" height="334"></img></div>
                      <% else %>
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div>
                    <% end if %></td>
                
				
				
				
				
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  
  <% if rs("foto_grande2")<>"imovel00000.jpg" then%>
  <tr>
  <td height="18"><table width="590" border="0" cellspacing="0" cellpadding="0">
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
                  <td width="290">&nbsp;</td>
                  <td width="290"> <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:anterior()" class="link" onmouseover="window.status='Anterior'; return true" onmouseout="window.status=''"><img src="bt_anterior002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:proxima()" class="link" onmouseover="window.status='Próxima'; return true" onmouseout="window.status=''"><img src="bt_proxima002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                      </tr>
                    </table> </td>
                </tr>
              </table></td>
            <td width="5">&nbsp;</td>
          </tr>
        </table></td>
  </tr>
  <%else%><%end if%>
  
  <tr>
    <td height="18"><div align="right"><table width="590" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="570"><div align="center">
			  <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
        <%else%>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
        foi atualizado com sucesso.</font> 
        <% end if %>

			  
			  
			  
			  </div></td>
              
            <td> <a href="javascript:newWindow3('form_adicionar_foto.asp?varCodImovel=<%=varCodImovel%>')"><img src="bt_mais03.jpg" width="18" height="18" border="0"></a></td>
            </tr>
          </table></td>
  </tr>
  
  
  
  
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de inclus&atilde;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs("data")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      da &uacute;ltima atualiza&ccedil;&atilde;o</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs("data_atualizacao")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			  
			   <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Capta&ccedil;&atilde;o</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if session("permissao") = "4" or session("permissao") = "5" then %><input name="txt_captacao" value="<%=rs("captacao")%>" id="txt_captacao" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"><%else%><%=rs("captacao")%><input name="txt_captacao" type="hidden" value="<%=rs("captacao")%>" id="txt_captacao" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"><%end if%></font></td>
              </tr>
			  
			   
			   <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link 
                      para o im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" > 
                    <div align="left">
                      <input name="txt_telefone2" type="text" id="txt_telefone2" value="http://www.imobiliariaveja.com.br/mostrar_imovel2.asp?varCodimovel=<%=rs("Cod_imovel")%>" size="38" maxlength="200" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%> ">
                    </div></td>
              </tr>
			   
			   
			   
			   
			    <tr> 
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                      do im&oacute;vel</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                    <div align="left"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Cod_imovel")%></font> 
                    </div></td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Uacute;ltima 
                      foto inclu&iacute;da</font></div></td>
                <td style="border:1px solid #FFFFFF;" ><div align="left">
                      <input  type="text" name="ultimo" value="<%if not rs2.eof then%><%=rs2.movefirst%>
					 
					 
					 
					 <% if rs2("Foto_Grande5") <> "imovel00000.jpg" then %>
					 <%=rs2("Foto_Grande5")%>
					 <%end if%>
					 
					 <% if rs2("Foto_Grande5") = "imovel00000.jpg" and rs2("foto_grande4") <> "imovel00000.jpg" then %>
					 <%=rs2("Foto_Grande4")%>
					 <%end if%>
					 
					 <% if rs2("Foto_Grande4") = "imovel00000.jpg" and  rs2("foto_grande3") <> "imovel00000.jpg"  then %>
					 <%=rs2("Foto_Grande3")%>
					 <%end if%>
					 
					 <% if rs2("Foto_Grande3") = "imovel00000.jpg" and  rs2("foto_grande2") <> "imovel00000.jpg" then %>
					 <%=rs2("Foto_Grande2")%>
					 <%end if%>
					 
					 <% if rs2("Foto_Grande2") = "imovel00000.jpg" and  rs2("foto_grande1") <> "imovel00000.jpg" then %>
					 <%=rs2("Foto_Grande1")%>
					 <%end if%>
					 
					 <% if rs2("Foto_Grande1") = "imovel00000.jpg" and  rs2("foto_grande") <> "imovel00000.jpg" then %>
					 <%=rs2("Foto_Grande")%>
					 <%end if%>
					 
					 <%else%>Sem registro<%end if%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>;">
                  </div></td>
              </tr>
			  
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Propriet&aacute;rio 
                    do im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                      <input name="txt_proprietario" type="text" id="txt_proprietario" value="<%=rs("proprietario")%>" size="38" maxlength="35" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;">
                  </div></td>
              </tr>
              <tr> 
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      do propriet&aacute;rio do im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;" bgcolor="<%=medio%>">
				<div align="left">
				<input name="txt_telefone" type="<% if session("permissao") = "1"  then response.write "Hidden" else response.write "text" end if %>" id="txt_telefone" value="<%=rs("telefone")%>" size="38" maxlength="30" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;">
				</div>
				</td>
              
			  </tr>
              <tr> 
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                      do propriet&aacute;rio do im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;" bgcolor="<%=claro%>">
				<div align="left">
				<input name="txt_email" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" id="txt_email" value="<%=rs("email")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>;">
				</div>
				</td>
              </tr>
              <tr> 
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                      do Im&oacute;vel</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left"><input name="txt_endereco" type="text" id="txt_endereco" value="<%=rs("endereco")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=medio%>;"></div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                      Grande</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                      <input name="blob" type="text"  class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"  value="<%=rs("foto_grande")%>" size="25" maxlength="120" align="left" >
                  </div></td>
              </tr>
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto 
                      Pequena</font></div></td>
                <td style="border:1px solid #FFFFFF;"><div align="left">
                      <input name="blob" type="text"  class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"  value="<%=rs("foto_pequena")%>" size="25" maxlength="120" align="left" >
                  </div></td>
              </tr>
			  
			  
			    <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                      do An&uacute;ncio</font> </div></td>
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<input name="txt_titulo" value="<%=rs("titulo_anuncio")%>"  type="text" id="txt_titulo4" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    do An&uacute;ncio</font> </div></td>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_anuncio" value="<%=rs("texto_anuncio")%>" type="text" id="txt_anuncio" size="38" maxlength="120" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>;">
                  </td>
              </tr>
			  
			  
			  
			  
			   <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presen&ccedil;a 
                      na primeira p&aacute;gina</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_presenca_primeira" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="<%=rs("presenca_primeira")%>"selected><%=rs("presenca_primeira")%></option>
				    <option value="excluido">Excluído</option>
                    <option value="incluido">Incluído</option>
                  </select>
                  </td>
              </tr>
			  
			  
			  
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link 
                      de visualiza&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <%
					
					dim varFoto
					 if rs("link_foto") = "icon_foto.gif" then
					 varFoto = "Com Foto"
					 else
					 varFoto = "Sem Foto"
					 end if
					 
					  %>
                    <select name="txt_link_foto" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="<%=rs("link_foto")%>" selected><%=varFoto%></option>
					  <option value="icon_foto.gif">Com Foto</option>
                      <option value="icon_foto2.gif">Sem Foto</option>
                    </select>
                     
                    </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="<% if rs("cidade") = "não informado" or rs555.eof then response.write "cqualquer" else response.write rs555("id_combo1") end if  %>" select><%=rs("cidade")%></option>
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
				 
				  </select>
                      <a href="javascript:newWindow3('form_incluir_cidade.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a></div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF"> 
                      <select name="combo2" class="inputBox" onChange="javascript:atualizacarros222(this.form);" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                       
                        <option value="<%if rs("bairro") = "não informado" or rs4.eof then response.write "bqualquer" else response.write rs4("id_combo2") end if%>" ><%=rs("bairro")%></option>
                       
                        <option value=""></option>
                      
                      </select>
                      <a href="javascript:newWindow3('form_incluir_bairro.asp')"><img src="bt_mais02.jpg" width="18" height="18" border="0"></a> 
                      </font></div></td>
              </tr>
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF"> 
                      <select name="combo5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="<%if rs("vila") <> "não informado" and  rs("vila") <>"" and  not rs444.eof then response.write rs444("id_combo3") else response.write "vlqualquer" end if%>" selected><%if rs("vila") <> "não informado" and  rs("vila") <>"" then response.write rs("vila") else response.write "não informado" end if%></option>
                      </select>
                      <a href="javascript:newWindow3('form_incluir_vila.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a> 
                      </font></div></td>
              </tr>
			  
			  
			  
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF">
                    <select name="txt_tipo" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>;">
                      <option value="<%=rs("tipo")%>" selected><%=rs("tipo")%></option>
                       <option value="Apartamento">Apartamento </option>
				   <option value="Térrea/Sobrado">Térrea/Sobrado</option>
				   <option value="Chácara">Chácara</option>
                  <option value="Flat">Flat</option>
				  <option value="Fazenda">Fazenda</option>
                  <option value="Prédio Comercial">Prédio Comercial</option>
                  <option value="Galpões">Galpões</option>
                  <option value="Sala Comercial">Sala Comercial</option>
				  <option value="Salão Comercial">Salão Comercial</option>
                  <option value="Terreno/Área">Terreno/Área</option>
                  <option value="Ponto Comercial">Ponto Comercial</option>
				  <option value="Cobertura">Cobertura</option>
                    </select>
                      </font></div></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Total</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left"><font color="#FFFFFF"> 
                      <input name="txt_a_total" value="<%=rs("area_total")%>" type="text" id="txt_a_total" size="12" maxlength="20" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">m&sup2;</font> 
                      </font></div></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Constru&iacute;da </font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                      <input name="txt_a_constr" value="<%=rs("area_construida")%>" type="text" id="txt_a_constr" size="12" maxlength="20" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>;">
                      <font color="#FFFFFF"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">m&sup2;</font> 
                      </font></div></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_quartos" id="select4" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="<%=rs("quartos")%>" selected><% if rs("quartos") = "0" then response.write "não informado" else response.write rs("quartos") end if%></option>
                      <option value="não informado">não informado</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                    </select>
                    </div></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_banheiros" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="<%=rs("banheiros")%>" selected><%=rs("banheiros")%></option>
                      <option value="não informado">não informado</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                    </select>
                    </div></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na Garagem</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_vagas" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="<%=rs("vagas")%>" selected><% if rs("vagas") = "0" then response.write "não informado" else response.write rs("vagas") end if%></option>
                      <option value="não informado">não informado</option>
                      <option value="01">01</option>
                      <option value="02">02</option>
                      <option value="03">03</option>
                      <option value="04">04</option>
                      <option value="05">05</option>
                      <option value="06">06</option>
                      <option value="07">07</option>
                      <option value="08">08</option>
                      <option value="09">09</option>
                    </select>
                    </div></td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                    <select name="txt_negociacao" id="select7" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <% if rs("negociacao") = "aluguel" then %>
                      <option value="aluguel" selected>Aluguel</option>
                      <option value="venda">Venda</option>
                      <%end if%>
                      <% if rs("negociacao") = "venda" then %>
                      <option value="aluguel">Aluguel</option>
                      <option value="venda" selected>Venda</option>
                      <% end if %>
                    </select>
                    </div></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="left">
                      <input name="txt_valor" value="<%=FormatNumber(rs("valor"),2)%>" type="text" id="txt_valor2" size="12" maxlength="30" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    </div></td>
              </tr>
              <tr>
			  
			  
			    <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     <option value="<%=rs("standby")%>" selected><%=rs("standby")%></option>
					  <option value="excluido" >Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                     <option value="<%=rs("ocupacao")%>" selected><%=rs("ocupacao")%></option>
					<option value="não informado" >não informado</option>
                    <option value="vago">vago</option>
                    <option value="ocupado">ocupado</option>
                    
                  </select>
                  </td>
              </tr>
			  
			  <tr>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Qualidade 
                      do neg&oacute;cio</font></div></td>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_qualidade" id="txt_qualidade" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="<%if rs("qualidade") <> "" then response.write rs("qualidade") else  response.write "não informado" end if%>" selected><%if rs("qualidade") <> "" then response.write rs("qualidade") else  response.write "não informado" end if%></option>
					  <option value="bom negócio" >Bom Negócio</option>
                    <option value="negócio comum" >Negócio Comum</option>
                    
                  </select></td>
              </tr>
			  
			  
			   
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                          <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&atilde;o 
                              sobre o im&oacute;vel</font></div></td>
                      </tr>
                      <tr>
                          <td width="290" height="82" bgcolor="<%=medio%>">&nbsp;</td>
                      </tr>
                    </table></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="left">
                      <textarea name="obs_imovel" class="inputBox" id="obs_imovel" style="HEIGHT: 100px; WIDTH: 290px; background: <%=medio%> " onKeyPress="return limitfield(this, 600)"><%=rs("obs_imovel")%></textarea>
                  </div></td>
              </tr>
              <tr> 
                <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center">
                    <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                          <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&atilde;o 
                              sobre o propriet&aacute;rio</font></div></td>
                      </tr>
                      <tr>
                          <td width="290" height="82" bgcolor="<%=medio%>">&nbsp;</td>
                      </tr>
                    </table>
                  </div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="left">
                    <textarea name="obs_proprietario" class="inputBox" id="obs_proprietario" style="HEIGHT: 100px; WIDTH: 290px; background: <%=medio%> " onKeyPress="return limitfield(this, 800)"><%=rs("obs_proprietario")%></textarea>
                  </div></td>
              </tr>
              <tr> 
                <td><div align="center"></div></td>
                <td><div align="left">
                    <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                          <td><% if  session("permissao") = "3" or session("permissao") = "2" or session("permissao") = "4" or session("permissao") = "5" then %><input name="image" type="image"  src="bt_atualizar002.jpg" width="145" height="18" border="0"><%else%><img  src="bt_atualizar002.jpg" width="145" height="18" border="0"></img><%end if%></td>
                          <td><% if  session("permissao") = "3" or session("permissao") = "2" or session("permissao") = "4" or session("permissao") = "5" then %><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_restaurar001.jpg" width="145" height="18" border="0"></img></a><%else%><img src="bt_restaurar001.jpg" width="145" height="18" border="0"></img><%end if%></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
 <%
           rs.Close
           'fecha a conexão
           Conexao.Close
		   Set objFSO = Nothing
           Set rs = Nothing
           %>
  <% response.flush%>
  <%response.clear%>





</body>



</html>
