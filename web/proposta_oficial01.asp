<!--#include file="dsn.asp"-->
<%

dim Conexao
dim Sql
dim rs
dim varCod_imovel

varCod_imovel = request.QueryString("varCod_imovel")

Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn

'Abrindo a tabela MARCAS!
Sql = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento  FROM imoveis where cod_imovel="&varCod_imovel 

Set rs = Server.CreateObject("ADODB.RecordSet")

	rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao
	
	
	rs.Open sql, Conexao





%>





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

	<script>
function isValidDigitNumber (doublecombo)
{


if (doublecombo.txt_nome.value == "") {
        alert("Você precisa indicar o seu nome!");
        doublecombo.txt_nome.focus();
		
        return false;
    }
	

if (doublecombo.txt_nacionalidade.value == "") {
        alert("Você precisa indicar sua nacionalidade!");
        doublecombo.txt_nacionalidade.focus();
		
        return false;
    }

	
	if (doublecombo.txt_estado_civil.value == "") {
        alert("Você precisa indicar o seu estado civil!");
        doublecombo.txt_estado_civil.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_profissao.value == "") {
        alert("Você precisa indicar a sua profissao!");
        doublecombo.txt_profissao.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_rg.value == "") {
        alert("Você precisa indicar o seu RG!");
        doublecombo.txt_rg.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_endereco.value == "") {
        alert("Você precisa indicar o seu endereço!");
        doublecombo.txt_endereco.focus();
		
        return false;
    }
	
	if (doublecombo.txt_cidade.value == "") {
        alert("Você precisa indicar a sua cidade!");
        doublecombo.txt_cidade.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_bairro.value == "") {
        alert("Você precisa indicar o seu bairro!");
        doublecombo.txt_bairro.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_estado.value == "") {
        alert("Você precisa indicar o seu estado!");
        doublecombo.txt_estado.focus();
		
        return false;
    }
	
	
	if (doublecombo.txt_valor.value == "") {
        alert("Você precisa indicar o valor que você pretende pagar pelo imóvel!");
        doublecombo.txt_valor.focus();
		
        return false;
    }
	
	
	if (!doublecombo.txt_pagamento_vista[0].checked && !doublecombo.txt_pagamento_vista[1].checked) {
  alert("Por favor, informe se você pretende pagar a vista");
  return false;
  }

	
if (doublecombo.txt_pagamento_vista[1].checked && doublecombo.txt_outro_valor01.value == "") {
  alert("Por favor, informe suas condições de pagamento");
  doublecombo.txt_outro_valor01.focus();
  return false;
  }	
	
	
	
	
	

if (vercpf(document.doublecombo.txt_cpf.value)) 
{}else 
{errors="1";if (errors) alert('CPF NÃO VÁLIDO');

document.retorno = (errors == '');
doublecombo.txt_cpf.focus();
 return false;

}
{
{
function vercpf (txt_cpf) 
{if (txt_cpf.length != 11 || txt_cpf == "00000000000" || txt_cpf == "11111111111" || txt_cpf == "22222222222" || txt_cpf == "33333333333" || txt_cpf == "44444444444" || txt_cpf == "55555555555" || txt_cpf == "66666666666" || txt_cpf == "77777777777" || txt_cpf == "88888888888" || txt_cpf == "99999999999")
return false;
add = 0;
for (i=0; i < 9; i ++)
add += parseInt(txt_cpf.charAt(i)) * (10 - i);
rev = 11 - (add % 11);
if (rev == 10 || rev == 11)
rev = 0;
if (rev != parseInt(txt_cpf.charAt(9)))
return false;
add = 0;
for (i = 0; i < 10; i ++)
add += parseInt(txt_cpf.charAt(i)) * (11 - i);
rev = 11 - (add % 11);
if (rev == 10 || rev == 11)
rev = 0;
if (rev != parseInt(txt_cpf.charAt(10)))
return false;
 return true;
 }
	
	}
	
}

	
	
	
	
	
	
}
</script>	
	
	


</head>

<body>
<form name="doublecombo" onSubmit="return isValidDigitNumber(this);"   method="post" action="incluir_proposta_oficial01.asp?varCod_imovel=<%=rs("cod_imovel")%>&vAtendimento=<%=rs("captacao")%>">
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="800" height="1000"><table width="800" height="1000" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="100"><strong>Proposta para compra do im&oacute;vel com a 
            refer&ecirc;ncia de n&uacute;mero:<%=rs("cod_imovel")%> , no site da Veja Admin e vendas 
            de bens im&oacute;veis s/c ltda, creci 11.676-J, empresa esta que 
            contrata para intermedia&ccedil;&atilde;o na realiza&ccedil;&atilde;o 
            do neg&oacute;cio imobili&aacute;rio.</strong></td>
        </tr>
		
        <tr> 
          <td><table width="800" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="150" height="30"><div align="center">Nome completo:</div></td>
                <td width="650" height="30"><input name="txt_nome" type="text" id="txt_nome" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Nacionalidade:</div></td>
                <td width="650" height="30"><input name="txt_nacionalidade" type="text" id="txt_nacionalidade" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Estado Civil:</div></td>
                <td width="650" height="30"><input name="txt_estado_civil" type="text" id="txt_estado_civil" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Profiss&atilde;o:</div></td>
                <td width="650" height="30"><input name="txt_profissao" type="text" id="txt_profissao" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">RG:</div></td>
                <td width="650" height="30"><input name="txt_rg" type="text" id="txt_rg" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">CPF:</div></td>
                <td width="650" height="30"><input name="txt_cpf" type="text" id="txt_cpf" style="HEIGHT: 30px; WIDTH: 650px;"   ></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Estado:</div></td>
                <td width="650" height="30"><input name="txt_estado" type="text" id="txt_estado" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Cidade:</div></td>
                <td width="650" height="30"><input name="txt_cidade" type="text" id="txt_cidade" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Bairro:</div></td>
                <td width="650" height="30"><input name="txt_bairro" type="text" id="txt_bairro" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Endere&ccedil;o:</div></td>
                <td width="650" height="30"><input name="txt_endereco" type="text" id="txt_endereco" style="HEIGHT: 30px; WIDTH: 650px;"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="50"><div align="center">Por meio desta, me proponho a comprar 
              o im&oacute;vel com a descri&ccedil;&atilde;o abaixo:</div></td>
        </tr>
        <tr> 
          <td><table width="800" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="150" height="30"><div align="center">Cidade:</div></td>
                <td height="30"><input name="textfield102" type="text" style="HEIGHT: 30px; WIDTH: 650px;" value="<%=rs("cidade")%>"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Bairro:</div></td>
                <td height="30"><input name="textfield103" type="text" style="HEIGHT: 30px; WIDTH: 650px;" value="<%=rs("bairro")%>"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Tipo:</div></td>
                <td height="30"><input name="textfield104" type="text" style="HEIGHT: 30px; WIDTH: 650px;" value="<%=rs("tipo")%>"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Vagas:</div></td>
                <td height="30"><input name="textfield105" type="text" style="HEIGHT: 30px; WIDTH: 650px;" value="<%=rs("vagas")%>"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">Quartos:</div></td>
                <td height="30"><input name="textfield106" type="text" style="HEIGHT: 30px; WIDTH: 650px;" value="<%=rs("quartos")%>"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center">C&oacute;digo 
                    no site:</div></td>
                <td height="30"><input name="textfield107" type="text" style="HEIGHT: 30px; WIDTH: 650px;" value="<%=rs("cod_imovel")%>"></td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center"></div></td>
                <td height="30">&nbsp;</td>
              </tr>
              <tr> 
                <td width="150" height="30"><div align="center"></div></td>
                <td height="30">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="50"><div align="center">Se aceita for a forma de pagamento 
              e o valor a que me proponho a pagar, conforme descrito abaixo:</div></td>
        </tr>
        <tr> 
          <td height="30"><table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="175">Valor total a ser pago: R$</td>
                <td width="150"><div align="center">
                    <input name="txt_valor" type="text" id="txt_valor" style="HEIGHT: 30px; WIDTH: 150px;">
                  </div></td>
                <td width="175"><div align="center">Pagamento a vista:</div></td>
                <td width="150"> <input type="radio" name="txt_pagamento_vista" value="sim">
                  Sim</td>
                <td width="150"> <input type="radio" name="txt_pagamento_vista" value="não">
                  N&atilde;o</td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="50"><div align="center">N&atilde;o sendo a vista, descreve 
              abaixo a forma de pagamento que voc&ecirc; pretende usar para comprar 
              o im&oacute;vel:</div></td>
        </tr>
        <tr> 
          <td><table width="800" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="30" height="30"><div align="center"><strong>R$</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_valor01" type="text" id="txt_outro_valor01" style="HEIGHT: 30px; WIDTH: 300px;"></td>
                <td width="170" height="30">
<div align="center"><strong>Representado 
                    por;</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_forma01" type="text" id="txt_outro_forma01" style="HEIGHT: 30px; WIDTH: 300px;"></td>
              </tr>
              <tr> 
                <td width="30" height="30"><div align="center"><strong>R$</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_valor02" type="text" id="txt_outro_valor02" style="HEIGHT: 30px; WIDTH: 300px;"></td>
                <td width="170" height="30">
<div align="center"><strong>Representado 
                    por;</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_forma02" type="text" id="txt_outro_forma02" style="HEIGHT: 30px; WIDTH: 300px;"></td>
              </tr>
              <tr> 
                <td width="30" height="30"><div align="center"><strong>R$</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_valor03" type="text" id="txt_outro_valor03" style="HEIGHT: 30px; WIDTH: 300px;"></td>
                <td width="170" height="30">
<div align="center"><strong>Representado 
                    por;</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_forma03" type="text" id="txt_outro_forma03" style="HEIGHT: 30px; WIDTH: 300px;"></td>
              </tr>
              <tr> 
                <td width="30" height="30"><div align="center"><strong>R$</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_valor04" type="text" id="txt_outro_valor04" style="HEIGHT: 30px; WIDTH: 300px;"></td>
                <td width="170" height="30">
<div align="center"><strong>Representado 
                    por;</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_forma04" type="text" id="txt_outro_forma04" style="HEIGHT: 30px; WIDTH: 300px;"></td>
              </tr>
              <tr> 
                <td width="30" height="30"><div align="center"><strong>R$</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_valor05" type="text" id="txt_outro_valor05" style="HEIGHT: 30px; WIDTH: 300px;"></td>
                <td width="170" height="30">
<div align="center"><strong>Representado 
                    por;</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_forma05" type="text" id="txt_outro_forma05" style="HEIGHT: 30px; WIDTH: 300px;"></td>
              </tr>
              <tr> 
                <td width="30" height="30"><div align="center"><strong>R$</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_valor06" type="text" id="txt_outro_valor06" style="HEIGHT: 30px; WIDTH: 300px;"></td>
                <td width="170" height="30">
<div align="center"><strong>Representado 
                    por;</strong></div></td>
                <td width="300" height="30"><input name="txt_outro_forma06" type="text" id="txt_outro_forma06" style="HEIGHT: 30px; WIDTH: 300px;"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="50"><div align="center">Fa&ccedil;a obs: tais como, prazo 
              que quer as chaves, o que deseja que fique no im&oacute;vel, ou 
              algo da sua preocupa&ccedil;&atilde;o:</div></td>
        </tr>
        <tr> 
          <td height="100"><textarea name="txt_obs_proposta_oficial" id="txt_obs_proposta_oficial" style="HEIGHT: 100px; WIDTH: 800px;"></textarea></td>
        </tr>
        <tr> 
          <td height="50"><div align="center"><strong>Dispensado neste ato a necessidade 
              da minha assinatura para validar o compromisso aqui assumido</strong></div></td>
        </tr>
        <tr> 
          <td height="70"><div align="center">
                <input name="image2" type="image"  src="bt_enviar_proposta01.jpg" width="200" height="50" border="0"  >
              </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</form>

<%
rs.close

set rs = nothing

%>
</body>
</html>
