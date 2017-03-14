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
   
   
   '-----------------------------ver o número máximo para se fazer a transferência-------
   dim rs444Indicacao01
   dim strSQL444Indicacao01
    Set rs444Indicacao01 = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444Indicacao01 = "SELECT  compradores.area_construida,compradores.area_total,compradores.condominio,compradores.cod_compradores   from compradores where atendimento like '"&vAtendimento01&"'"

         
         rs444Indicacao01.CursorLocation = 3
        rs444Indicacao01.CursorType = 3
        
 
            rs444Indicacao01.ActiveConnection = Conexao




	rs444Indicacao01.Open strSQL444Indicacao01, Conexao
   
   if int(rs444Indicacao01.recordCount) < int(vNumero)  then
   
   response.Redirect "form_via_codigo01.asp?varSucesso="&"Não foi possível fazer a transferência,os clientes selecionados passaram do limite!"&""
   end if
   
   
   '-------------------------------
   rs444Indicacao01.close
  
  
  set rs444Indicacao01 = nothing
 '-----------------------------
   '------------------------------------------------------------------------------------
   
   
   
   
   
   
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444Indicacao = "SELECT TOP "&vNumero&" compradores.area_construida,compradores.area_total,compradores.condominio,compradores.cod_compradores   from compradores where atendimento like '"&vAtendimento01&"'"

         
         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3
        
 
            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	dim num
	num = 0
	
	 if not rs444Indicacao.eof  then 
				     While   num < int(vNumero)
                   
                    'Conexao.execute"update imoveis set imovel_em_negociacao='"&"não informado"&"' where imovel_em_negociacao IS NULL" 
                   
				    Conexao.execute"update  compradores set atendimento='"&vAtendimento02&"' where atendimento like '"&vAtendimento01&"' and cod_compradores="&rs444Indicacao("cod_compradores")
                   
				   num = num +1
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	'-------------------------------
   rs444Indicacao.close
  
  
  set rs444Indicacao = nothing
 '------------------------------
	
	
	
	'response.write num&"||"&vNumero
	

response.Redirect "form_via_codigo01.asp?varSucesso="&"Mudança realizada com sucesso!"&""

%>

<% response.flush%>
	   <%response.clear%>
</body>
</html>
