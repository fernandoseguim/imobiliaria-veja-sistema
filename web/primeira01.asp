<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_imoveis.asp"-->
<%response.Buffer = true %>



<%
'Criando conex�o com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1 ASC" 

Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open sql3, Conexao3




	
	
	
	
	
	
	
	'--------------------------------------------------------------------
	



%> 


<%
dim varNotFind

varNotFind = request.QueryString("varNotFind")

dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT * FROM combo2 where id_combo1 = 4  ORDER BY id_combo2" 
	
	
	Conexao.Open dsn
	
	
	
	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs4.ActiveConnection = Conexao
	
	
	rs4.Open strSQL4, Conexao

dim rs55
dim strSQL55

Set rs55 = Server.CreateObject("ADODB.RecordSet")
	strSQL55 = "SELECT * FROM imoveis ORDER BY cod_imovel DESC" 



rs55.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs55.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs55.ActiveConnection = Conexao

rs55.open strSQL55, Conexao


%>





<%


'

Sql333 = "SELECT * FROM combo2 ORDER BY id_combo2" 
Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs33.ActiveConnection = Conexao3
	
	
	rs33.Open sql333, Conexao3



dim rsFrontPage,SQLFrontPage,objFSO 

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Set rsFrontPage = Server.CreateObject("ADODB.RecordSet")

SQLFrontPage = "SELECT * FROM imoveis where presenca_primeira like '"&"incluido"&"' ORDER BY cod_imovel DESC"

rsFrontPage.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsFrontPage.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsFrontPage.ActiveConnection = Conexao


rsFrontPage.open SQLFrontPage,Conexao

dim intRecordCount 


intRecordCount = rsFrontPage.RecordCount




'------------------------------selecionar os tipos de im�vel para o formul�rio-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT * FROM tipo  ORDER BY tipo ASC" 
	
	
	rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs444Tipo22.ActiveConnection = Conexao
	
	
	
	 rs444Tipo22.Open strSQL444Tipo22, Conexao










'-------------------------------------------------------------------------------------------------









'----------------------------------------------------------------------------





%> 




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>P�gina inicial</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>


<script>

// Verifica se somente n�meros foram digitados no campo
function isValidDigitNumber (doublecombo) 
{

if (doublecombo.example2.value == "nqualquer") {
		alert("Por favor, escolha a negocia��o desejada.");
		doublecombo.example2.focus();
		
		return false;
}

var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar seu telefone, digite apenas n�meros!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}


var strValidNumber1_5="a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,x,z,w,y,1,2,3,4,5,6,7,8,9,0,@,.,_,-";
for (nCount=0; nCount < doublecombo.txt_email.value.length; nCount++) 
		{
strTempChar1_5=doublecombo.txt_email.value.substring(nCount,nCount+1);
if (strValidNumber1_5.indexOf(strTempChar1_5,0)==-1) 
{
alert("Ao colocar seu email,use somente min�sculas!");
doublecombo.txt_email.focus();
doublecombo.txt_email.select();
return false;
}
}


if (doublecombo.combo1.value == "cqualquer") {
		alert("Voc� precisa escolher uma cidade.");
		doublecombo.combo1.focus();
		
		return false;
}


if (doublecombo.stage22.value == "vqualquer") {
		alert("Voc� precisa escolher um valor.");
		doublecombo.stage22.focus();
		
		return false;
}


}


</script>

<script>

// Verifica se somente n�meros foram digitados no campo
function isValidDigitNumber2 (doublecombo2) 



{




{


if (doublecombo2.ref.value == "Busca por refer�ncia:") {
		alert("Por favor,digite um n�mero de refer�ncia , pois assim , voc� ter� um atendimento preferencial e exclusivo.");
		doublecombo2.ref.focus();
		
		return false;
}







var strValidNumber1_4="1234567890";
for (nCount=0; nCount < doublecombo2.ref.value.length; nCount++) 
		{
strTempChar1_4=doublecombo2.ref.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("Ao colocar sua refer�ncia, digite apenas n�meros!");
doublecombo2.ref.focus();
doublecombo2.ref.select();
return false;
}
}






}
}
</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>


<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<table bgcolor="#FFFFFF" width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="106"><img src="top01.jpg" width="794" height="106"></td>
  </tr>
  <tr>
    <td height="237"><table width="794" height="257" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="552" height="257">
            <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="552" height="257">
              <param name="movie" value="frontpage001.swf">
              <param name="quality" value="high">
              <embed src="frontpage001.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="552" height="257"></embed></object>
            </td>
          <td width="242" bgcolor="#e0a94e"><div align="center">
              <table width="232" height="247" border="0" cellpadding="0" cellspacing="0" bgcolor="#e6dca9">
                <tr>
                  <td bgcolor="#e6dca9">
<div align="center">
                      <table width="222" height="237" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e0a94e"><div align="center">
                              <table width="212" height="227" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td bgcolor="#e6dca9"><div align="center"> 
                                      <table width="202" height="217" border="0" cellpadding="0" cellspacing="0">
                                       
									    <tr> 
										
                                          <td><div align="center">
										  
											  <table width="202" border="0" cellspacing="0" cellpadding="0">
                                               <form name="doublecombo" target="_blank" onSubmit="return isValidDigitNumber(this);" method="post" action="listar_imoveis01.asp">
									   
											   
											    <tr> 
                                                  <td height="20"><input name="txt_nome" onFocus="doublecombo.txt_nome.value=''"  type="text" class="inputBox" id="txt_nome"  style="HEIGHT: 18px; WIDTH: 202px; ;  background: #b2802c; color:#FFFFFF; font-size:12px" value="<% if session("nome") <> "" then response.write session("nome") else response.write "Seu nome:" end if%>"></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><input name="txt_telefone" onFocus="doublecombo.txt_telefone.value=''"  type="text" class="inputBox" id="txt_nome2"  style="HEIGHT: 18px; WIDTH: 202px;  background: #b2802c; color:#FFFFFF;  font-size:12px" value="<% if session("telefone") <> "" then response.write session("telefone") else response.write "Seu telefone:" end if%>"></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><input name="txt_email" onFocus="doublecombo.txt_email.value=''"  type="text" class="inputBox" id="txt_email"  style="HEIGHT: 18px; WIDTH: 202px;  background: #b2802c; color:#FFFFFF;  font-size:12px" value="<% if session("email") <> "" then response.write session("email") else response.write "Seu email:" end if%>"></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="combo1" id="combo1" onChange="javascript:atualizacarros(this.form);" size="1"  class="inputBox" style="HEIGHT: 18px;   WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                     
													  <option value="cqualquer" selected>Cidade</option>
				  <% if not rs3.eof then %>
                  <% While NOT Rs3.EoF %>
                  <option value="<% = Rs3("id_combo1") %>" > 
                  <% = Rs3("nome_combo1") %>
                  </option>
                  <% Rs3.MoveNext %>
                  <% Wend %>
				  <option value="cqualquer">qualquer uma</option>
                  <%else%>
                  <option value=""></option>
                  <%end if%>
												   
												    </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="combo2" id="combo2"   size="1"  class="inputBox" style="HEIGHT: 18px;   WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                      <option value="bqualquer" selected>Bairro/Regi�o</option>
				  <% if not rs4.eof then%>
                  <% While NOT Rs4.EoF %>
                  <option value="<% = Rs4("id_combo2") %>"> 
                  <% = Rs4("nome_combo2") %>
                  </option>
                  <% Rs4.MoveNext %>
				  
                  <% Wend %>
				   <option value="bqualquer">qualquer um</option>
				  
                  <% else %>
                  <option value=""></option>
                  <% end if %>
                                                    </select></td>
                                                </tr>
												
												 <tr> 
                                                  <td height="20"><select name="txt_tipo" id="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 18px;  WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                        <option value="tqualquer" selected>Tipo</option>
                                                        <option value="tqualquer">Qualquer 
                                                        um</option>
                                                        <% if not rs444Tipo22.eof then%>
                                                        <% While NOT (rs444Tipo22.EoF) %>
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
												
												
												
												
                                                <tr> 
                                                  <td height="20"><select name="txt_quartos" id="txt_quartos" size="1"  class="inputBox" style="HEIGHT: 18px;   WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                        <option value="qqualquer" selected>Quartos</option>
                                                        <option value="qqualquer">Qualquer 
                                                        um</option>
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
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="txt_garagem" id="txt_garagem" size="1"  class="inputBox" style="HEIGHT: 18px;   WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                        <option value="gqualquer" selected>Vagas 
                                                        na Garagem</option>
                                                        <option value="gqualquer">Qualquer 
                                                        um</option>
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
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="example2" id="example2" onChange="redirect2(this.options.selectedIndex)" size="1"  class="inputBox" style="HEIGHT: 18px;   WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                        <option value="nqualquer" selected>Negocia��o 
                                                        </option>
                                                        <option value="nqualquer" >Qualquer 
                                                        um </option>
                                                        <option  value="Aluguel">Aluguel 
                                                        </option>
                                                        <option value="Venda">Venda 
                                                        </option>
                                                      </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><select name="stage22" id="stage22" size="1"  class="inputBox" style="HEIGHT: 18px;   WIDTH: 202px ;background: #b2802c; color:#FFFFFF;  font-size:12px">
                                                        <option value="vqualquer" selected>Valor</option>
                                                        <option value="vqualquer">Qualquer 
                                                        um</option>
                                                        <option value="0000000000 0000020000">At� 
                                                        20.000,00</option>
                                                        <option value="0000020001 0000050000">20.001,00 
                                                        at� 50.000,00</option>
                                                        <option value="0000050001 0000080000">50.001,00 
                                                        at� 80.000,00</option>
                                                        <option value="0000080001 0000110000">80.001,00 
                                                        at� 110.000,00</option>
                                                        <option value="0000110001 0000150000">110.001,00 
                                                        at� 150.000,00</option>
                                                        <option value="0000150001 0000200000">150.001,00 
                                                        at� 200.000,00</option>
                                                        <option value="0000200001 0000250000">200.001,00 
                                                        at� 250.000,00</option>
                                                        <option value="0000250001 0000300000">250.001,00 
                                                        at� 300.000,00</option>
                                                        <option value="0000300001 0000350000">300.001,00 
                                                        at� 350.000,00</option>
                                                        <option value="0000350001 0000400000">350.001,00 
                                                        at� 400.000,00</option>
														
														<option value="0000400001 0000600000">400.001,00 
                                                        at� 600.000,00</option>
														<option value="0000600001 0000800000">600.001,00 
                                                        at� 800.000,00</option>
														
														<option value="0000800001 0001000000">800.001,00 
                                                        at� 1000.000,00</option>
														
                                                        <option value="0001000001 1000000000">Acima 
                                                        de 1000.000,00</option>
														
                                                      </select></td>
                                                </tr>
                                                <tr> 
                                                  <td height="20"><input name="image" type="image" src="bt_procurar303.jpg" width="201" height="18"></td>
                                                </tr>
												</form>
                                              </table>
                                            </div></td>
                                        </tr>
                                      </table>
                                    </div></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="16"><img src="subtop01.jpg" width="794" height="16"></td>
  </tr>
   
 
  <tr>
  <tr>
    <td height="60" valign="middle"> 
      <table  height="22" border="0" align="left"  cellpadding="0" cellspacing="0">
        <form name="doublecombo2" target="_blank" onSubmit="return isValidDigitNumber2(this);" method="post" action="listar_imoveis02.asp">
  
	<tr>
            <td width="400"><div align="right"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Busca por 
                refer&ecirc;ncia: </font></div></td>   
            <td width="202"><input name="ref"   type="text" class="inputBox" id="ref"  style="border-top : 1px solid;border-bottom : 1px solid;border-left : 1px solid;border-right : 1px solid;border-color:#e9dca8;HEIGHT: 20px; WIDTH: 202px; ; font-size : 9px; background:#FFFFFF; color:#9d9249;" value=""></td>
    <td width="23"><input name="image2" type="image"  src="bt_lupa01.jpg" width="23" height="20" border="0"></td>
    
  </tr>
  </form>
</table>

    </td>
  </tr>
  
  <%' if rsFrontPage.count 
  
  if  (rsFrontPage.recordcount >= 6) then
  
   %>
    <td height="200"><table width="794" height="200" border="0" cellpadding="0" cellspacing="0">
        
		
		
		
		
		
		
		
		
		
		
		
		
		<tr>
          <td width="200"><div align="center">
              <table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td  style="border:1px solid #ddddc5;"><div align="center">
                      <table width="180" height="180" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e9dca8">
<div align="center">
                              <table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td><table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                          <td width="165" height="97" style="border:2px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="167" height="97" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n�o dispon�vel</strong></a></font></div><%end if%></td>
                                      </tr>
									  <tr><td height="5"></td></tr>
                                      <tr>
                                          <td width="165" height="62" bgcolor="#f7ecbf" ><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table>
            </div></td>
          <td width="200"><div align="center"><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
				<%=rsFrontPage.movenext%>
                  <td  style="border:1px solid #ddddc5;"><div align="center">
                      <table width="180" height="180" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e9dca8">
<div align="center">
                              <table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td><table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                          <td width="165" height="102" style="border:2px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="167" height="97" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n�o dispon�vel</strong></a></font></div><%end if%></td>
                                      </tr>
									  
									  <tr><td height="5"></td></tr>
									  
                                      <tr>
                                          <td width="165" height="62" bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></div></td>
          <td width="413"><div align="center">
              <table width="422" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="422" style="border:1px solid #ddddc5;"><div align="center">
                      <table width="374" height="180" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e9dca8"><div align="center">
                              <table width="364" height="170" border="0" cellpadding="0" cellspacing="0">
                               <%=rsFrontPage.movenext%>
							    <tr>
                                  <td width="154" height="170"><table width="149" height="170" border="0" align="left" cellpadding="0" cellspacing="0">
                                        <tr>
                                          <td bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                        </tr>
                                      </table></td>
                                  <td width="210" height="170"><table width="210" border="0" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td width="210" height="128" style="border:2px solid #FFFFFF;"><div align="center"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="206" height="124" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n�o dispon�vel</strong></a></font></div><%end if%></div></td>
                                      </tr>
                                      <tr>
                                        <td width="210" height="42"><table width="210" height="42" border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                              <td bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249">Destaque</a></strong></font></div></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
				
				
				
				
				
				
				
				
				
				
				
              </table>
            </div></td>
        </tr>
		
		
		<tr>
          <td height="20"></td>
        </tr>
		
		
		
		
		
		
		
      </table></td>
	  
	  </tr>
	  <tr><td height="200" >
	  <table width="794" height="200" border="0" cellpadding="0" cellspacing="0">
        
		
		
		
		
		
		
		
		
		
		
		
		
		<tr>
          <td width="200"><div align="center">
              <table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td  style="border:1px solid #ddddc5;"><div align="center">
                      <table width="180" height="180" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e9dca8">
<div align="center">
                              <table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td><table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                          <td width="165" height="97" style="border:2px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="167" height="97" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n�o dispon�vel</strong></a></font></div><%end if%></td>
                                      </tr>
									  <tr><td height="5"></td></tr>
                                      <tr>
                                          <td width="165" height="62" bgcolor="#f7ecbf" ><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table>
            </div></td>
          <td width="200"><div align="center"><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
				<%=rsFrontPage.movenext%>
                  <td  style="border:1px solid #ddddc5;"><div align="center">
                      <table width="180" height="180" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e9dca8">
<div align="center">
                              <table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td><table width="165" height="165" border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                          <td width="165" height="102" style="border:2px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="167" height="97" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n�o dispon�vel</strong></a></font></div><%end if%></td>
                                      </tr>
									  
									  <tr><td height="5"></td></tr>
									  
                                      <tr>
                                          <td width="165" height="62" bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></div></td>
          <td width="413"><div align="center">
              <table width="422" height="190" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="422" style="border:1px solid #ddddc5;"><div align="center">
                      <table width="374" height="180" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td bgcolor="#e9dca8"><div align="center">
                              <table width="364" height="170" border="0" cellpadding="0" cellspacing="0">
                               <%=rsFrontPage.movenext%>
							    <tr>
                                  <td width="154" height="170"><table width="149" height="170" border="0" align="left" cellpadding="0" cellspacing="0">
                                        <tr>
                                          <td bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                        </tr>
                                      </table></td>
                                  <td width="210" height="170"><table width="210" border="0" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td width="210" height="128" style="border:2px solid #FFFFFF;"><div align="center"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="206" height="124" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n�o dispon�vel</strong></a></font></div><%end if%></div></td>
                                      </tr>
                                      <tr>
                                        <td width="210" height="42"><table width="210" height="42" border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                              <td bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>')" style="text-decoration:none;color:#9d9249">Destaque</a></strong></font></div></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                    </table></td>
                                </tr>
                              </table>
                            </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
				
				
				
				
				
				
				
				
				
				
				
              </table>
            </div></td>
        </tr>
		
		
		<tr>
          <td height="20"></td>
        </tr>
		
		
		
		
		
		
		
      </table>
	  
	  
	  
	  
	  </td></tr>
	  
	<tr></tr>
	 
	  
	  <%else%>
	  
	  <%end if%>
	  
  </tr>
  
  <%
  '--------------------Separar por Franquia-------------
dim SqlFranquia
dim rsFranquia

SqlFranquia = "SELECT franquia.id_franquia,franquia.nome_franquia,franquia.data_franquia,franquia.endereco,franquia.telefone,franquia.email FROM franquia where nome_franquia ='"&session("vOrigem_Franquia")&"' ORDER BY id_franquia DESC"  

Set rsFranquia = Server.CreateObject("ADODB.RecordSet")

	rsFranquia.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsFranquia.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rsFranquia.ActiveConnection = Conexao3
	
	
	rsFranquia.Open sqlFranquia, Conexao3
  
  if not rsFranquia.eof then
  %>
  
  
  <tr>
    <td>&nbsp;</td>
  </tr>
  
  
  
  
  
  <tr>
    <td height="120" bgcolor="#e9dca8" ><table width="784" height="110" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="784" height="110" bgcolor="#f7ecbf" style="border:1px solid #f7ecbf;">
<div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('form_enviar_email.asp')" style="color:#9d9249;text-decoration:none;"><font size="2"><strong>Imobili&aacute;ria 
              Veja</strong></font> <br>
              <font size="2"><%=rsFranquia("endereco")%><br>
              CRECI: 11.676-J<br>
              Contato: <%=rsFranquia("telefone")%> || <%=rsFranquia("email")%><br>
              Copyright 2008 Imobili&aacute;ria Veja- Todos os direitos reservados</font></a></font> 
            </div></td>
        </tr>
      </table></td>
  </tr>
  <%else%>
  
  
  <% end if %>
  
  
</table>
</form>

<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui � criada uma vari�vel "groups" que receber� os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a vari�vel group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero at� o n�mero de elementos do array "groups" */

group2[i2]=new Array()
/* aqui � criado o array "group" que receber� valores conforme o n�mero de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receber� valores de op��es. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("201,00 at� 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 at� 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 at� 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 at� 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 at� 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 at� 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 at� 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 at� 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 at� 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("At�  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 at� 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 at� 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 at� 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 at� 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 at� 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 at� 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 at� 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 at� 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 at� 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("400.001,00 at� 600.000,00","0000400001 0000600000")
group2[3][13]=new Option("600.001,00 at� 800.000,00","0000600001 0000800000")
group2[3][14]=new Option("800.001,00 at� 1000.000,00","0000800001 0001000000")
group2[3][15]=new Option("Acima de 1000.000,00","0001000001 1000000000")









/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


var temp2=document.doublecombo.stage22
/* aqui a vari�vel "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui � criada a fun��o "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que d� um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que � escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" ser� o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a vari�vel "location" recebe os valores de "stage2" que corresponde ao endere�o de
link para o carregamento de p�gina. */


//-->
</script>
 <% response.flush%>
  <%response.clear%>

<%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receber� uma conexao aberta!
'Em funcoes, geralmente n�o criamos objetos do tipo conex�es!
'Opte por sempre deixar sua fun��o o mais compat�vel poss�vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa fun��o JavaScript recebe o form em que est�o os campos a serem atualizados!
'Veja na chamada da fun��o no m�todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op��o foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op��es de carro!
SqlMarcas3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utiliz�o

rsMarcas3.ActiveConnection = Conexao3

rsMarcas3.Open SqlMarcas3, Conexao3


While NOT rsMarcas3.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"


Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utiliz�o

rsCarros3.ActiveConnection = Conexao3

rsCarros3.Open SqlCarros3, Conexao3


'Fazemos um loop por todos os carros, criando uma nova op��o no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Regi�o" & "','" & "bqualquer" & "');"& vbcrlf
i = 1
While NOT rsCarros3.EoF

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
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



  rsMarcas3.Close           
		   
           Set rsMarcas3 = Nothing
		   
		   
		   rsCarros3.Close           
		   
           Set rsCarros3 = Nothing






End Function


        




%> 

<%  EscreveFuncaoJavaScript ( Conexao3 ) %>



</body>
</html>
