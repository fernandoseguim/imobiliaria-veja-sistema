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

vAssunto_ligar_urgente = " Um novo im�vel foi atualizado e ocorreu uma indica��o, ligue imediatamente para esse comprador"
  
   
    Set Conexao = Server.CreateObject("ADODB.Connection")
   Conexao.Open dsn
   
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

	strSQL444Indicacao = "Select proposta.Cod_proposta,proposta.proposta_proposta,proposta.foto_proposta,proposta.foto_proposta,proposta.nome_proposta,proposta.telefone_proposta,proposta.email_proposta,proposta.data_proposta,proposta.horario_proposta,proposta.interesse_proposta,proposta.cod_imovel_proposta,proposta.clique,proposta.origem_franquia  from proposta ORDER BY cod_proposta DESC"


         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3

            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	 if not rs444Indicacao.eof then 
				     While NOT rs444Indicacao.EoF 
                   
                    'Conexao.execute"update imoveis set data_contato='"&now()&"' where cod_imovel ="&rs444Indicacao("cod_imovel") 
                   
				   'Conexao.execute"update compradores set condominio='"&"0"&"',area_total='"&"0"&"',area_construida='"&"0"&"'" 
                   
				   ' Conexao.execute"update compradores set standby='"&"comprador a contatar"&"' where (standby like 'excluido' or standby like 'n�o informado') and cod_compradores="&rs444Indicacao("cod_compradores") 
                   
				   Conexao.execute"update proposta set origem_franquia='"&"Sao Bernardo"&"' where  cod_proposta="&rs444Indicacao("cod_proposta") 
                   
				   
				   
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	'-------------------------------
   rs444Indicacao.close
  
  
  set rs444Indicacao = nothing
 '------------------------------
	
	
	
	
	

'response.Redirect "form_via_codigo01.asp?varSucesso="&"Mudan�a realizada com sucesso!"&""
response.write "Atualiza��o de franquia realizada com sucesso !"
%>

<% response.flush%>
	   <%response.clear%>
</body>
</html>
