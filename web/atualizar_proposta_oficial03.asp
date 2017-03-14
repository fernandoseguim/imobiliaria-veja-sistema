<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
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
 
 
 
 
 
 
 
 dim varCod_proposta_oficial
 
 varCod_proposta_oficial = request.QueryString("varCod_proposta_oficial")
 
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



dim vNome_contra
 dim vNacionalidade_contra
 dim vEstado_civil_contra
 dim vProfissao_contra
 dim vRg_contra
 dim vCpf_contra
 dim vEndereco_contra
 dim vCidade_contra
 dim vBairro_contra
 dim vEstado_contra
 dim vValor_contra
 dim vPagamento_vista_contra
 dim vOutro_valor_contra01
 dim vOutro_valor_contra02 
 dim vOutro_valor_contra03
 dim vOutro_valor_contra04                                               
dim vOutro_valor_contra05														  
dim vOutro_valor_contra06

dim vOutro_forma_contra01
dim vOutro_forma_contra02 
dim vOutro_forma_contra03
dim vOutro_forma_contra04                                               
dim vOutro_forma_contra05														  
dim vOutro_forma_contra06
dim vObs_proposta_oficial_contra




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




vNome_contra = request.Form("txt_nome_contra")
  vNacionalidade_contra = request.Form("txt_nacionalidade_contra")
  vEstado_civil_contra = request.Form("txt_estado_civil_contra")
  vProfissao_contra = request.Form("txt_profissao_contra")
  vRg_contra = request.Form("txt_rg_contra")
  vCpf_contra = request.Form("txt_cpf_contra")
  vEndereco_contra = request.Form("txt_endereco_contra")
  vCidade_contra = request.Form("txt_cidade_contra")
  vBairro_contra = request.Form("txt_bairro_contra")
  vEstado_contra = request.Form("txt_estado_contra")
  vValor_contra = request.Form("txt_valor_contra")
  vPagamento_vista_contra = request.Form("txt_pagamento_vista_contra")
  vOutro_valor_contra01 = request.Form("txt_outro_valor_contra01")
  vOutro_valor_contra02  = request.Form("txt_outro_valor_contra02")
  vOutro_valor_contra03 = request.Form("txt_outro_valor_contra03")
  vOutro_valor_contra04 = request.Form("txt_outro_valor_contra04")                                               
 vOutro_valor_contra05 = request.Form("txt_outro_valor_contra05")														  
 vOutro_valor_contra06 = request.Form("txt_outro_valor_contra06")

 vOutro_forma_contra01 = request.Form("txt_outro_forma_contra01")
 vOutro_forma_contra02  = request.Form("txt_outro_forma_contra02")
 vOutro_forma_contra03 = request.Form("txt_outro_forma_contra03")
 vOutro_forma_contra04 = request.Form("txt_outro_forma_contra04")                                              
 vOutro_forma_contra05 = request.Form("txt_outro_forma_contra05")														  
 vOutro_forma_contra06 = request.Form("txt_outro_forma_contra06")

 vObs_proposta_oficial_contra = request.Form("txt_obs_proposta_oficial_contra")







    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	 
	
	
	 
 
	 Conexao.execute"update proposta_oficial set nome_contra='"&vNome_contra&"',nacionalidade_contra='"&vNacionalidade_contra&"',estado_civil_contra='"&vEstado_civil_contra&"',profissao_contra='"&vProfissao_contra&"',rg_contra='"&vRg_contra&"',cpf_contra='"&vCpf_contra&"',endereco_contra='"&vEndereco_contra&"',cidade_contra='"&vCidade_contra&"',bairro_contra='"&vBairro_contra&"',estado_contra='"&vEstado_contra&"',outro_valor_contra01='"&vOutro_valor_contra01&"',outro_valor_contra02='"&vOutro_valor_contra02&"',outro_valor_contra03='"&vOutro_valor_contra03&"',outro_valor_contra04='"&vOutro_valor_contra04&"',outro_valor_contra05='"&vOutro_valor_contra05&"',outro_valor_contra06='"&vOutro_valor_contra06&"',outro_forma_contra01='"&vOutro_forma_contra01&"',outro_forma_contra02='"&vOutro_forma_contra02&"',outro_forma_contra03='"&vOutro_forma_contra03&"',outro_forma_contra04='"&vOutro_forma_contra04&"',outro_forma_contra05='"&vOutro_forma_contra05&"',outro_forma_contra06='"&vOutro_forma_contra06&"',pagamento_vista_contra='"&vPagamento_vista_contra&"',obs_proposta_oficial_contra='"&vObs_proposta_oficial_contra&"',valor_contra='"&vValor_contra&"' where cod_proposta_oficial="&varCod_proposta_oficial
	 
	 
	 
	
	
	response.Redirect "proposta_oficial03.asp?varSucesso="&"Mensagem enviada com sucesso"&"&varCod_imovel="&varCod_imovel&"&varCod_proposta_oficial="&varCod_proposta_oficial&""
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>












 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
