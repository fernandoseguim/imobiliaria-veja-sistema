<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if

Dim Conexao,strSQL,rs,vdata
 Dim vdata2

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
 
 dim vNome
 dim vNacionalidade
 dim vEstado_civil
 dim vProfissao
 dim vRg
 dim vCpf
 dim vEndereco
 dim vCidade
 dim vBairro
 dim vEstado
 dim vValor
 dim vPagamento_vista
 dim vOutro_valor01
 dim vOutro_valor02 
 dim vOutro_valor03
 dim vOutro_valor04                                               
dim vOutro_valor05														  
dim vOutro_valor06

dim vOutro_forma01
dim vOutro_forma02 
dim vOutro_forma03
dim vOutro_forma04                                               
dim vOutro_forma05														  
dim vOutro_forma06

dim vObs_proposta_oficial
dim vAtendimento
dim varCod_imovel

 vNome = request.Form("txt_nome")
  vNacionalidade = request.Form("txt_nacionalidade")
  vEstado_civil = request.Form("txt_estado_civil")
  vProfissao = request.Form("txt_profissao")
  vRg = request.Form("txt_rg")
  vCpf = request.Form("txt_cpf")
  vEndereco = request.Form("txt_endereco")
  vCidade = request.Form("txt_cidade")
  vBairro = request.Form("txt_bairro")
  vEstado = request.Form("txt_estado")
  vValor = request.Form("txt_valor")
  vPagamento_vista = request.Form("txt_pagamento_vista")
  vOutro_valor01 = request.Form("txt_outro_valor01")
  vOutro_valor02  = request.Form("txt_outro_valor02")
  vOutro_valor03 = request.Form("txt_outro_valor03")
  vOutro_valor04 = request.Form("txt_outro_valor04")                                               
 vOutro_valor05 = request.Form("txt_outro_valor05")														  
 vOutro_valor06 = request.Form("txt_outro_valor06")

 vOutro_forma01 = request.Form("txt_outro_forma01")
 vOutro_forma02  = request.Form("txt_outro_forma02")
 vOutro_forma03 = request.Form("txt_outro_forma03")
 vOutro_forma04 = request.Form("txt_outro_forma04")                                              
 vOutro_forma05 = request.Form("txt_outro_forma05")														  
 vOutro_forma06 = request.Form("txt_outro_forma06")

 vObs_proposta_oficial = request.Form("txt_obs_proposta_oficial")
vAtendimento = request.QueryString("vAtendimento")
varCod_imovel = request.QueryString("varCod_imovel")







    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	 dim SqlAtendimento01
	 dim rsAtendimento01
	 
	 SqlAtendimento01 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.historico_atual01,compradores.historico_atual02,compradores.historico_atual03,compradores.historico_atual04,compradores.historico_atual05,compradores.historico_atual06,compradores.historico_quem01,compradores.historico_quem02,compradores.historico_quem03,compradores.historico_quem04,compradores.historico_quem05,compradores.historico_quem06,compradores.ocupacao_hist,compradores.endereco_hist,compradores.valor_hist,compradores.quartos_hist,compradores.vagas_hist,compradores.suites_hist,compradores.piscina_hist,compradores.area_total_hist,compradores.area_construida_hist,compradores.edicula_hist,compradores.condominio_hist  FROM compradores where (telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"') order by cod_compradores DESC "
    

Set rsAtendimento01 = Server.CreateObject("ADODB.RecordSet")

	rsAtendimento01.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsAtendimento01.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsAtendimento01.ActiveConnection = Conexao
	
	
	rsAtendimento01.Open sqlAtendimento01, Conexao


vAtendimento = rsAtendimento01("atendimento")

rsAtendimento01.close

set rsAtendimento01 = nothing
	 
	 
	
	
	 
Conexao.execute "Insert into proposta_oficial(nome, telefone, email ,cod_imovel,atendimento,nacionalidade,estado_civil,profissao,rg,cpf,endereco,cidade,bairro,estado,valor,pagamento_vista,outro_valor01,outro_valor02,outro_valor03,outro_valor04,outro_valor05,outro_valor06,outro_forma01,outro_forma02,outro_forma03,outro_forma04,outro_forma05,outro_forma06,obs_proposta_oficial,data,nome_contra,nacionalidade_contra,estado_civil_contra,rg_contra,cpf_contra,profissao_contra,endereco_contra,cidade_contra,bairro_contra,estado_contra,valor_contra,outro_valor_contra01,outro_valor_contra02,outro_valor_contra03,outro_valor_contra04,outro_valor_contra05,outro_valor_contra06,outro_forma_contra01,outro_forma_contra02,outro_forma_contra03,outro_forma_contra04,outro_forma_contra05,outro_forma_contra06,obs_proposta_oficial_contra,origem_franquia) values( '"& vNome &"','"& session("telefone") &"','"& session("email") &"','"& varCod_imovel &"','"& vAtendimento &"','"& vNacionalidade &"','"& vEstado_civil &"','"& vProfissao &"','"& vRg &"','"& vCpf &"','"& vEndereco &"','"& vCidade &"','"& vBairro &"','"& vEstado &"','"& vValor &"','"& vPagamento_vista &"','"& vOutro_valor01 &"','"& vOutro_valor02 &"','"& vOutro_valor03 &"','"& vOutro_valor04 &"','"& vOutro_valor05 &"','"& vOutro_valor06 &"','"& vOutro_forma01 &"','"& vOutro_forma02 &"','"& vOutro_forma03 &"','"& vOutro_forma04 &"','"& vOutro_forma05 &"','"& vOutro_forma06 &"','"& vObs_proposta_oficial &"','"& now() &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& session("vOrigem_Franquia") &"')" 
	 
	 
	 
	
	
	response.Redirect "proposta_oficial01.asp?varSucesso="&"Mensagem enviada com sucesso"&"&varCod_imovel="&varCod_imovel&""
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>












 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
