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
	
	 

	strSQL444Indicacao = "Select proposta_oficial.cod_proposta_oficial,proposta_oficial.nome,proposta_oficial.telefone,proposta_oficial.email,proposta_oficial.cod_imovel,proposta_oficial.nacionalidade,proposta_oficial.estado_civil,proposta_oficial.profissao,proposta_oficial.rg,proposta_oficial.cpf,proposta_oficial.endereco,proposta_oficial.cidade,proposta_oficial.bairro,proposta_oficial.estado,proposta_oficial.valor,proposta_oficial.pagamento_vista,proposta_oficial.outro_valor01,proposta_oficial.outro_valor02,proposta_oficial.outro_valor03,proposta_oficial.outro_valor04,proposta_oficial.outro_valor05,proposta_oficial.outro_valor05,proposta_oficial.outro_valor06,proposta_oficial.outro_forma01,proposta_oficial.outro_forma02,proposta_oficial.outro_forma03,proposta_oficial.outro_forma04,proposta_oficial.outro_forma05,proposta_oficial.outro_forma06,proposta_oficial.obs_proposta_oficial,proposta_oficial.nome_contra,proposta_oficial.nacionalidade_contra,proposta_oficial.estado_civil_contra,proposta_oficial.profissao_contra,proposta_oficial.rg_contra,proposta_oficial.cpf_contra,proposta_oficial.endereco_contra,proposta_oficial.cidade_contra,proposta_oficial.bairro_contra,proposta_oficial.estado_contra,proposta_oficial.valor_contra,proposta_oficial.outro_valor_contra01,proposta_oficial.outro_valor_contra02,proposta_oficial.outro_valor_contra03,proposta_oficial.outro_valor_contra04,proposta_oficial.outro_valor_contra05,proposta_oficial.outro_valor_contra06,proposta_oficial.outro_forma_contra01,proposta_oficial.outro_forma_contra02,proposta_oficial.outro_forma_contra03,proposta_oficial.outro_forma_contra04,proposta_oficial.outro_forma_contra05,proposta_oficial.outro_forma_contra06,proposta_oficial.obs_proposta_oficial_contra,proposta_oficial.atendimento,proposta_oficial.data,proposta_oficial.origem_franquia  from proposta_oficial   ORDER BY cod_proposta_oficial DESC"


         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3

            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	
	 if not rs444Indicacao.eof then 
				     While NOT rs444Indicacao.EoF 
                   
                    'Conexao.execute"update imoveis set data_contato='"&now()&"' where cod_imovel ="&rs444Indicacao("cod_imovel") 
                   
				   'Conexao.execute"update compradores set condominio='"&"0"&"',area_total='"&"0"&"',area_construida='"&"0"&"'" 
                   
				   ' Conexao.execute"update compradores set standby='"&"comprador a contatar"&"' where (standby like 'excluido' or standby like 'não informado') and cod_compradores="&rs444Indicacao("cod_compradores") 
                   
				   Conexao.execute"update proposta_oficial set origem_franquia='"&"Sao Bernardo"&"' where  cod_proposta_oficial="&rs444Indicacao("cod_proposta_oficial") 
                   
				   
				   
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
