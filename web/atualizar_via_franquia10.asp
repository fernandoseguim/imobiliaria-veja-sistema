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
	
	 

	strSQL444Indicacao = "SELECT imovel_clicado.cod_clicado,imovel_clicado.nome,imovel_clicado.telefone,imovel_clicado.codigo_clicado,imovel_clicado.endereco_ip,imovel_clicado.data,imovel_clicado.tipo,imovel_clicado.quartos,imovel_clicado.vagas,imovel_clicado.cidade,imovel_clicado.bairro,imovel_clicado.valor,imovel_clicado.negociacao,imovel_clicado.tarja02,imovel_clicado.origem,imovel_clicado.origem_franquia  FROM imovel_clicado   ORDER BY cod_clicado DESC"


         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3

            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	 if not rs444Indicacao.eof then 
				     While NOT rs444Indicacao.EoF 
                   
                    'Conexao.execute"update imoveis set data_contato='"&now()&"' where cod_imovel ="&rs444Indicacao("cod_imovel") 
                   
				   'Conexao.execute"update compradores set condominio='"&"0"&"',area_total='"&"0"&"',area_construida='"&"0"&"'" 
                   
				   ' Conexao.execute"update compradores set standby='"&"comprador a contatar"&"' where (standby like 'excluido' or standby like 'não informado') and cod_compradores="&rs444Indicacao("cod_compradores") 
                   
				   Conexao.execute"update imovel_clicado set origem_franquia='"&"Sao Bernardo"&"' where  cod_clicado="&rs444Indicacao("cod_clicado") 
                   
				   
				   
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
