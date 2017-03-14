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



vAtendimento01 = request.form("txt_atendimento01")
vAtendimento02 = request.form("txt_atendimento02")


dim rs444Indicacao

dim strSQL444Indicacao

dim vAssunto_ligar_urgente

vAssunto_ligar_urgente = " Um novo imóvel foi atualizado e ocorreu uma indicação, ligue imediatamente para esse comprador"
  
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
   Conexao.Open dsn
   
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444Indicacao = "SELECT imoveis_procurados.cod_procurados,imoveis_procurados.cidade,imoveis_procurados.bairro,imoveis_procurados.tipo,imoveis_procurados.negociacao,imoveis_procurados.valor,imoveis_procurados.data,imoveis_procurados.enderecoIP,imoveis_procurados.quartos,imoveis_procurados.vagas,imoveis_procurados.nome,imoveis_procurados.telefone,imoveis_procurados.email,imoveis_procurados.origem_franquia FROM imoveis_procurados  ORDER BY Cod_procurados DESC" 


         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3

            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	 if not rs444Indicacao.eof then 
				     While NOT rs444Indicacao.EoF 
                   
                    'Conexao.execute"update imoveis set data_contato='"&now()&"' where cod_imovel ="&rs444Indicacao("cod_imovel") 
                   
				   'Conexao.execute"update compradores set condominio='"&"0"&"',area_total='"&"0"&"',area_construida='"&"0"&"'" 
                   
				   ' Conexao.execute"update compradores set standby='"&"comprador a contatar"&"' where (standby like 'excluido' or standby like 'não informado') and cod_compradores="&rs444Indicacao("cod_compradores") 
                   
				   Conexao.execute"update imoveis_procurados set origem_franquia='"&"Sao Bernardo"&"' where  cod_procurados="&rs444Indicacao("cod_procurados") 
                   
				   
				   
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	'-------------------------------
   rs444Indicacao.close
  
  
  set rs444Indicacao = nothing
 '------------------------------
	
	
	
	
	

'response.Redirect "form_via_codigo01.asp?varSucesso="&"Mudança realizada com sucesso!"&""
response.write "Atualização de franquia realizada com sucesso !"
%>

<% response.flush%>
	   <%response.clear%>
</body>
</html>
