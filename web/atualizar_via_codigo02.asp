<% response.buffer=True%>
<!--#include file="dsn.asp"-->
<% Server.ScriptTimeout = 600 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>


<%

dim vAtendimento01
dim vAtendimento02
dim vNumero


vAtendimento01 = request.form("txt_atendimento01")
vAtendimento02 = request.form("txt_atendimento02")
vNumero = request.form("txt_numero")

dim rs444Indicacao

dim strSQL444Indicacao

dim vAssunto_ligar_urgente

vAssunto_ligar_urgente = " Um novo imóvel foi atualizado e ocorreu uma indicação, ligue imediatamente para esse comprador"
  
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
   Conexao.Open dsn
   
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444Indicacao = "SELECT TOP "&vNumero&" imoveis.condominio,imoveis.captacao,imoveis.cod_imovel from imoveis where captacao like '"&vAtendimento01&"'"


         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3

            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao


      if int(rs444Indicacao.recordCount) < int(vNumero)  then
   
   response.Redirect "form_via_codigo02.asp?varSucesso="&"Não foi possível fazer a transferência,os imóveis selecionados passaram do limite!"&""
   end if
	
	
	 if not rs444Indicacao.eof then 
				     While NOT rs444Indicacao.EoF 
                   
                    'Conexao.execute"update imoveis set imovel_em_negociacao='"&"não informado"&"' where imovel_em_negociacao IS NULL" 
                   
				    'Conexao.execute"update imoveis set condominio='"&"0"&"' where condominio like '%"&"0,00"&"%'" 
                    Conexao.execute"update  imoveis set captacao='"&vAtendimento02&"' where captacao like '"&vAtendimento01&"' and cod_imovel="&rs444Indicacao("cod_imovel")
                   
				   
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	'-------------------------------
   rs444Indicacao.close
  
  
  set rs444Indicacao = nothing
 '------------------------------
	
	
	
	
	

response.Redirect "form_via_codigo02.asp?varSucesso="&"Mudança realizada com sucesso!"&""

%>

<% response.flush%>
	   <%response.clear%>
</body>
</html>
