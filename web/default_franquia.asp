<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="style2_primeira.asp"-->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<
<%

dim Conexao3

Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

dim vAbrir

vAbrir = ""

dim vOrigem_Franquia

vOrigem_Franquia = request.form("vOrigem_Franquia")

session("vOrigem_Franquia") = vOrigem_Franquia

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "Sao Bernardo"
end if

'--------------------Separar por Franquia-------------
dim SqlFranquia
dim rsFranquia

SqlFranquia = "SELECT franquia.id_franquia,franquia.nome_franquia,franquia.data_franquia,franquia.endereco,franquia.telefone,franquia.email FROM franquia where nome_franquia ='"&session("vOrigem_Franquia")&"' ORDER BY id_franquia DESC"  

Set rsFranquia = Server.CreateObject("ADODB.RecordSet")

	rsFranquia.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rsFranquia.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rsFranquia.ActiveConnection = Conexao3
	
	
	rsFranquia.Open sqlFranquia, Conexao3








%>

<script language="JavaScript">
var today=new Date();
var todaysec=today.getSeconds();

function xpop(){
{
window.open('form_enviar_email.asp', 'popUP','width=605,height=530,resizable=yes,scrollbars=yes,Left=0,Top=0')

}   
}
</script>




<title>Imobiliária Veja</title>

</head>

<body onBlur="_blank"  >
<table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="200"> 
      <div align="center">
        <table width="800" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="140"><div align="center">
                <table width="800" height="140" border="0" cellpadding="0" cellspacing="0">
                  <tr bgcolor="#e0a94e">
                    <td width="135" height="140" bgcolor="#e0a94e"> 
                      <div align="right"><img src="default_img01.jpg" width="135" height="137"></div></td>
                    <td width="469" height="140"><table width="469" height="140" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="40"><div align="center"></div></td>
                        </tr>
                        <tr>
                          <td height="30"><div align="center"><font color="#996600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="veja_dicas_comprando01.html" target="_blank" style="color:#000000;text-decoration:none;">Veja 
                              dicas comprando</a></strong></font></div></td>
                        </tr>
                        <tr>
                          <td height="30"><div align="center"><font color="#996600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="veja_dicas_vendendo01.html" target="_blank" style="color:#000000;text-decoration:none;">Veja 
                              dicas vendendo</a></strong></font></div></td>
                        </tr>
						<tr>
                          <td height="30"><div align="center"><font color="#996600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="colaborador01.asp" target="_blank" style="color:#000000;text-decoration:none;">Seja 
                              um colaborador da veja</a></strong></font></div></td>
                        </tr>
						
						
						
                        <tr>
                          <td height="40">&nbsp;</td>
                        </tr>
                      </table></td>
				    <td width="195" bgcolor="#e0a94e"> 
                      <div align="right"><img src="logotipo_caixa02.jpg" width="195" height="140"></div></td>
                   
                  </tr>
                </table>
              </div></td>
          </tr>
          
		  <%
		  if not rsFranquia.eof then
		  %>
		  <tr>
		  		  
            <td height="30"><div align="center"><font color="#996600" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone: 
                <%=rsFranquia("telefone")%></strong></font></div></td>
          </tr>
          <tr>
            <td height="30"><div align="center"><font color="#996600" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_enviar_email.asp" target="_blank" style="color:#996600;text-decoration:none;">Email: 
                <%=rsFranquia("email")%></a></strong></font></div></td>
          </tr>
		  <% else %>
		  
		  
		  <% end if %>
		  
		  
		  
		  <tr>
		  <td height="50"><table  height="22" border="0" align="left"  cellpadding="0" cellspacing="0">
        <form name="doublecombo2" target="_blank" onSubmit="return isValidDigitNumber2(this);" method="post" action="listar_imoveis02.asp">
  
	<tr>
            <td width="400"><div align="right"><font color="#996600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Busca 
                        por referência:</strong> </font></div></td>   
            <td width="202"><input name="ref"   type="text" class="inputBox" id="ref"  style="border-top : 1px solid;border-bottom : 1px solid;border-left : 1px solid;border-right : 1px solid;border-color:#e9dca8;HEIGHT: 20px; WIDTH: 202px; ; font-size : 9px; background:#FFFFFF; color:#9d9249;" value=""></td>
    <td width="23"><input name="image2" type="image"  src="bt_lupa01.jpg" width="23" height="20" border="0"></td>
    
  </tr>
  </form>
</table>
		  </td>
		  </tr>
        </table>
      </div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="250" valign="top"><table width="800" height="250" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="160" height="250"><table width="150" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><table width="140" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><table width="140" height="240" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="140" height="113"><a href="procurar_avaliacao02.asp" target="_blank"><img src="icone_front05.jpg" width="140" height="113" border="0"></a></td>
                          </tr>
                          <tr>
                            <td width="140" height="5"></td>
                          </tr>
                          <tr>
                            <td width="140" height="122"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="procurar_avaliacao02.asp" target="_blank" style="color:#000000;text-decoration:none;"><font color="#FF0000">Veja:</font>Aqui 
                                você pode fazer a avaliação 
                                do seu imóvel</a></strong></font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="160" height="250"><table width="150" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><table width="140" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><table width="140" height="240" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="140" height="113"><a href="procurar_compradores001.asp" target="_blank"><img src="icone_front01.jpg" width="140" height="113" border="0"></a></td>
                          </tr>
                          <tr>
                            <td width="140" height="5"></td>
                          </tr>
                          <tr>
                            <td width="140" height="122"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="procurar_compradores001.asp" target="_blank" style="color:#000000;text-decoration:none;"><font color="#FF0000">Veja:</font> 
                                Aqui você encontra um comprador ou inquilino 
                                para o seu imóvel</a></strong></font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="160" height="250"><table width="150" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><table width="140" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><table width="140" height="240" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="140" height="113"><div align="center"><a href="procurar_permuta001.asp" target="_blank"><img src="icone_front02.jpg" width="140" height="113" border="0"></a></div></td>
                          </tr>
                          <tr>
                            <td width="140" height="5"></td>
                          </tr>
                          <tr>
                            <td width="140" height="122"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="procurar_permuta001.asp" target="_blank" style="color:#000000;text-decoration:none;"><font color="#FF0000">Veja:</font> 
                                Aqui você encontra quem pode trocar de imóvel 
                                com você</a></strong></font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="160" height="250"><table width="150" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><table width="140" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><table width="140" height="240" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="140" height="113"><a href="acesso03.asp" target="_blank"><img src="icone_front04.jpg" width="140" height="113" border="0"></a></td>
                          </tr>
                          <tr>
                            <td width="140" height="5"></td>
                          </tr>
                          <tr>
                            <td width="140" height="122"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="contas.asp" target="_blank" style="color:#000000;text-decoration:none;"><font color="#FF0000">Veja:</font> 
                                Se novos imóveis, compradores ou permutantes 
                                estão no sistema para negociar com você</a></strong></font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="160" height="250"><table width="150" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e0a94e"><table width="140" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><table width="140" height="240" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="140" height="113"><a href="default04.asp" target="_blank"><img src="icone_front03.jpg" width="140" height="113" border="0"></a></td>
                          </tr>
                          <tr>
                            <td width="140" height="5"></td>
                          </tr>
                          <tr>
                            <td width="140" height="122"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="default04.asp" target="_blank" style="color:#000000;text-decoration:none;"><font color="#FF0000">Veja:</font> 
                                Nossos imóveis cadastrados para venda e 
                                locação </a></strong></font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table> </td>
  </tr>
</table>


<%

'response.write rsFranquia("endereco")&rsFranquia("telefone")&rsFranquia("email")

rsFranquia.close

set rsFranquia = nothing


%>
</body>
</html>
