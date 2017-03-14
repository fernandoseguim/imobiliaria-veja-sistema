




<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCodImovel,objFSO
dim varCodImovel2
dim varCod_imovel


varCod_Imovel = request.QueryString("varCod_Imovel")




   
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou FROM imoveis Where cod_imovel = "&varCod_Imovel
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
		
		
		
	if not(rs.eof) then
	
	
		 dim EnderecoIP
	 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
	 
	
		
  
  
	
	
	
	
	
	
	
	
	
	     
 %>




<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<!--#include file="style_imoveis.asp"-->
<body bgcolor="<%=escuro%>" topmargin="5" bottommargin="0" rightmargin="0" leftmargin="0" marginheight="0" marginwidth="0">


<form name="doublecombo"  method="post" action="incluir_querpagar.asp?varCodImovel=<%=varCodImovel%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  
  
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
 
 
</table>
</form>

<% else %>


<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">   Não foi encontrado o imóvel pedido!!</font>

<% end if %>



<%
'-----------------------------atualização de acesso-----------------
%>



 <%
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   Set objFSO = Nothing
		   set conexao = nothing
		   
		 
		   
		   
           %>
  <% response.flush%>
  <%response.clear%>
</body>
</html>
