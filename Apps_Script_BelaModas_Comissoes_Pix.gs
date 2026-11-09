
const CONFIG = {
  ABA: "Produtos",
  PLANILHA: "PLANILHA_ID",
  PASTA: "PASTA_FOTOS_ID",
  SENHA_HASH: "CHAVE_ADMIN_HASH",
  COLUNAS: [
    "codigo", "nome", "categoria", "descricao", "preco",
    "estoque", "disponibilidade", "foto_principal",
    "fotos_adicionais", "destaque", "ativo", "atualizado_em",
    "tamanhos", "video_url", "preco_valor", "ordem",
    "preco_promocional", "promocao_inicio", "promocao_fim"
  ],
  SESSAO_SEGUNDOS: 21600,
  TENTATIVAS_MAX: 8,
  TEMPO_RESPOSTA_SEGUNDOS: 120
};

function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.PLANILHA);
  if (!id) throw new Error("Planilha não configurada.");
  return SpreadsheetApp.openById(id);
}

function aba_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(CONFIG.ABA);

  if (!sh) sh = ss.insertSheet(CONFIG.ABA);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CONFIG.COLUNAS.length).setValues([CONFIG.COLUNAS]);
    return sh;
  }

  // Confere os 15 cabeçalhos antigos sem reescrever ou deslocar dados existentes.
  const antigos = CONFIG.COLUNAS.slice(0, 15);
  const atuais = sh.getRange(1, 1, 1, 15).getValues()[0];

  if (antigos.some((valor, i) => atuais[i] !== valor)) {
    throw new Error("Os cabeçalhos iniciais da aba Produtos não correspondem. Nenhum dado foi alterado.");
  }

  // Acrescenta somente a nova coluna. Nunca substitui os cabeçalhos antigos.
  const cabecalhoOrdem = sh.getRange(1, 16).getValue();
  if (cabecalhoOrdem === "") {
    if (sh.getLastColumn() >= 16 && sh.getLastRow() >= 2) {
      const dadosColuna = sh.getRange(2, 16, sh.getLastRow() - 1, 1).getValues();
      if (dadosColuna.some(linha => linha[0] !== "" && linha[0] !== null)) {
        throw new Error("A coluna 16 já contém dados. Verifique a aba Produtos antes de adicionar a coluna ordem.");
      }
    }
    sh.getRange(1, 16).setValue("ordem");
  } else if (cabecalhoOrdem !== "ordem") {
    throw new Error("A coluna 16 já possui outro cabeçalho. A coluna ordem não foi criada para evitar sobrescrever dados.");
  }

  garantirColunasPromocao_(sh);
  return sh;
}

function garantirColunasPromocao_(sh) {
  const cabecalhos = ["preco_promocional", "promocao_inicio", "promocao_fim"];
  cabecalhos.forEach((cabecalho, indice) => {
    const coluna = 17 + indice;
    const atual = sh.getRange(1, coluna).getValue();
    if (atual === cabecalho) return;
    if (atual !== "") {
      throw new Error("A coluna " + coluna + " já possui outro cabeçalho. A coluna " + cabecalho + " não foi criada para evitar sobrescrever dados.");
    }
    if (sh.getLastRow() >= 2 && sh.getRange(2, coluna, sh.getLastRow() - 1, 1).getValues().some(linha => linha[0] !== "" && linha[0] !== null)) {
      throw new Error("A coluna " + coluna + " contém dados. Verifique a aba Produtos antes de adicionar " + cabecalho + ".");
    }
    sh.getRange(1, coluna).setValue(cabecalho);
  });
}

function garantirOrdens_(sh) {
  const ultima = sh.getLastRow();
  if (ultima < 2) return [];

  const intervalo = sh.getRange(2, 16, ultima - 1, 1);
  const valores = intervalo.getValues();
  let alterou = false;

  valores.forEach((linha, i) => {
    const atual = linha[0];
    if (atual === "" || atual === null || !Number.isFinite(Number(atual))) {
      // Produtos antigos recebem a ordem atual da planilha, sem mudar sua posição.
      linha[0] = i + 1;
      alterou = true;
    } else {
      linha[0] = Number(atual);
    }
  });

  if (alterou || valores.some((linha, i) => linha[0] !== intervalo.getCell(i + 1, 1).getValue())) {
    intervalo.setValues(valores);
  }

  return valores.map(linha => Number(linha[0]));
}

function jsonp_(dados, callback) {
  const json = JSON.stringify(dados).replace(/</g, "\\u003c");

  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]{0,100}$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};

  if (p.acao === "entregas") {
    try {
      return jsonp_({ sucesso: true, config: obterConfiguracaoEntregas_() }, p.callback);
    } catch (err) {
      return jsonp_({ sucesso: false, erro: "Não foi possível carregar as configurações de entrega." }, p.callback);
    }
  }

  // Consulta pública da configuração do vídeo do banner.
  if (p.acao === "banner") {
    try {
      return jsonp_({
        sucesso: true,
        banner: obterBanner_()
      }, p.callback);
    } catch (err) {
      return jsonp_({
        sucesso: false,
        erro: "Não foi possível carregar o vídeo do banner."
      }, p.callback);
    }
  }

  if (p.acao === "resultado") {
    return consultarResultado_(p);
  }

  try {
    return jsonp_({
      sucesso: true,
      produtos: listar_()
    }, p.callback);
  } catch (err) {
    return jsonp_({
      sucesso: false,
      erro: "Não foi possível carregar o catálogo."
    }, p.callback);
  }
}

function consultarResultado_(p) {
  const nonce = String(p.nonce || "");
  const callback = p.callback || "";

  if (!/^[a-f0-9]{48}$/.test(nonce)) {
    return jsonp_({ sucesso: false, erro: "Consulta inválida." }, callback);
  }

  const cache = CacheService.getScriptCache();
  const chave = "resposta_" + nonce;
  const salvo = cache.get(chave);

  if (!salvo) {
    return jsonp_({ pendente: true }, callback);
  }

  cache.remove(chave);

  try {
    return jsonp_(JSON.parse(salvo), callback);
  } catch (err) {
    return jsonp_({
      sucesso: false,
      erro: "Não foi possível ler a resposta."
    }, callback);
  }
}

function listar_() {
  const sh = aba_();
  const ultima = sh.getLastRow();
  if (ultima < 2) return [];

  garantirOrdens_(sh);

  const dados = sh.getRange(2, 1, ultima - 1, CONFIG.COLUNAS.length).getValues();

  return dados
    .map((row, indice) => ({ row: row, indice: indice }))
    .filter(item => item.row[0] && item.row[10] !== false)
    .sort((a, b) => {
      const diferenca = Number(a.row[15]) - Number(b.row[15]);
      return diferenca || a.indice - b.indice;
    })
    .map(item => {
      const row = item.row;
      const produto = {};

      CONFIG.COLUNAS.forEach((col, i) => {
        const value = row[i];
        produto[col] = value instanceof Date ? value.toISOString() : value;
      });

      produto.ordem = Number(row[15]);
      produto.miniaturas = [];
      return produto;
    });
}

function doPost(e) {
  const p = e && e.parameter ? e.parameter : {};
  const nonce = String(p.nonce || "");

  if (!/^[a-f0-9]{48}$/.test(nonce)) {
    return ContentService.createTextOutput("Requisição inválida.");
  }

  let resposta;

  try {
    let resultado;

    if (p.acao === "login") {
      resultado = criarSessao_(p.senha || "");
    } else if (p.acao === "cadastrarVendedor") {
      resultado = cadastrarVendedor_(JSON.parse(p.vendedor || "{}"));
    } else if (p.acao === "loginVendedor") {
      resultado = criarSessaoVendedor_(p.email || "", p.senha || "");
    } else if (p.acao === "registrarPedidoVendedor") {
      resultado = registrarPedidoVendedor_(JSON.parse(p.pedido || "{}"));
    } else if (p.acao === "calcularEntrega") {
      resultado = calcularEntrega_(p.destino);
    } else if (p.acao === "painelVendedor" || p.acao === "trocarSenhaVendedor" || p.acao === "sairVendedor" || p.acao === "salvarChavePixVendedor") {
      const vendedorId = validarSessaoVendedor_(p.tokenVendedor);
      const dadosVendedorSessao = localizarVendedorPorId_(vendedorId).dados;
      if (dadosVendedorSessao[12] === true && p.acao === "painelVendedor") {
        throw new Error("Antes de acessar sua área, crie uma nova senha.");
      }
      if (p.acao === "painelVendedor") resultado = obterPainelVendedor_(vendedorId);
      else if (p.acao === "trocarSenhaVendedor") resultado = trocarSenhaVendedor_(vendedorId, p.novaSenha || "");
      else if (p.acao === "salvarChavePixVendedor") resultado = salvarChavePixVendedor_(vendedorId,p.chavePix||"");
      else { CacheService.getScriptCache().remove("sessao_vendedor_" + p.tokenVendedor); resultado = { saiu: true }; }
    } else {
      validarSessao_(p.token);

      if (p.acao === "excluirVendedor") {
        resultado = excluirVendedor_(p.id);
      } else if (p.acao === "salvar") {
        resultado = salvar_(JSON.parse(p.produto || "{}"));
      } else if (p.acao === "excluir") {
        resultado = excluir_(p.codigo);
      } else if (p.acao === "midia") {
        resultado = midia_(p);
      } else if (p.acao === "ordenar") {
        resultado = ordenar_(p.codigos);
      } else if (p.acao === "salvarBanner") {
        resultado = salvarBanner_(p.video_url);
      } else if (p.acao === "listarVendedores") {
        resultado = listarVendedores_();
      } else if (p.acao === "atualizarVendedor") {
        resultado = atualizarVendedor_(p.id, p.status, p.comissao);
      } else if (p.acao === "resetarSenhaVendedor") {
        resultado = resetarSenhaVendedor_(p.id);
      } else if (p.acao === "listarPedidosVendedores") {
        resultado = listarPedidosVendedores_();
      } else if (p.acao === "atualizarPedidoVendedor") {
        resultado = atualizarPedidoVendedor_(p.id, p.status);
      } else if (p.acao === "listarPagamentosVendedores") {
        resultado = listarPagamentosVendedores_();
      } else if (p.acao === "registrarPagamentoVendedor") {
        resultado = registrarPagamentoVendedor_(JSON.parse(p.pagamento||"{}"));
      } else if (p.acao === "salvarEntregas") {
        resultado = salvarConfiguracaoEntregas_(JSON.parse(p.config || "{}"));
      } else {
        throw new Error("Operação desconhecida.");
      }
    }

    resposta = {
      sucesso: true,
      resultado: resultado,
      nonce: nonce
    };

  } catch (err) {
    resposta = {
      sucesso: false,
      erro: err && err.message ? err.message : "Falha na operação.",
      nonce: nonce
    };
  }

  CacheService.getScriptCache().put(
    "resposta_" + nonce,
    JSON.stringify(resposta),
    CONFIG.TEMPO_RESPOSTA_SEGUNDOS
  );

  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}

function criarSessao_(senha) {
  const props = PropertiesService.getScriptProperties();
  const hashCorreto = props.getProperty(CONFIG.SENHA_HASH);

  if (!hashCorreto) {
    throw new Error("Senha administrativa não configurada.");
  }

  const cache = CacheService.getScriptCache();
  const chave = "tentativas_login_admin";
  const tentativas = Number(cache.get(chave) || 0);

  if (tentativas >= CONFIG.TENTATIVAS_MAX) {
    throw new Error("Muitas tentativas. Aguarde alguns minutos.");
  }

  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    senha,
    Utilities.Charset.UTF_8
  ).map(b => ("0" + ((b + 256) % 256).toString(16)).slice(-2)).join("");

  if (!senha || hash !== hashCorreto) {
    cache.put(chave, String(tentativas + 1), 600);
    throw new Error("Senha administrativa incorreta.");
  }

  cache.remove(chave);

  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put("sessao_" + token, "administrador", CONFIG.SESSAO_SEGUNDOS);

  return {
    token: token,
    validadeSegundos: CONFIG.SESSAO_SEGUNDOS
  };
}

function validarSessao_(token) {
  if (!token || typeof token !== "string" || token.length > 100) {
    throw new Error("Sessão inválida. Entre novamente.");
  }

  if (CacheService.getScriptCache().get("sessao_" + token) !== "administrador") {
    throw new Error("Sessão expirada ou inválida. Entre novamente.");
  }
}

function salvar_(p) {
  if (!p.codigo || !p.nome) {
    throw new Error("Código e nome são obrigatórios.");
  }

  const trava = LockService.getScriptLock();
  trava.waitLock(20000);

  try {
    const sh = aba_();
    const ultima = sh.getLastRow();
    const ordens = garantirOrdens_(sh);

    const codigos = ultima > 1
      ? sh.getRange(2, 1, ultima - 1, 1).getValues().flat()
      : [];

    const codigo = String(p.codigo).trim();
    const indice = codigos.findIndex(c =>
      String(c).trim().toLowerCase() === codigo.toLowerCase()
    );

    const preco = Number(p.preco_valor ?? p.preco ?? 0);
    const anteriores = indice >= 0
      ? sh.getRange(indice + 2, 17, 1, 3).getValues()[0]
      : ["", "", ""];
    const precoPromocional = Object.prototype.hasOwnProperty.call(p, "preco_promocional")
      ? (p.preco_promocional === "" || p.preco_promocional === null ? "" : Number(p.preco_promocional))
      : anteriores[0];
    const promocaoInicio = Object.prototype.hasOwnProperty.call(p, "promocao_inicio")
      ? String(p.promocao_inicio || "")
      : anteriores[1];
    const promocaoFim = Object.prototype.hasOwnProperty.call(p, "promocao_fim")
      ? String(p.promocao_fim || "")
      : anteriores[2];
    if (precoPromocional !== "" && (!Number.isFinite(Number(precoPromocional)) || Number(precoPromocional) <= 0 || Number(precoPromocional) >= preco)) {
      throw new Error("O preço promocional deve ser maior que zero e menor que o preço normal.");
    }
    if (promocaoInicio && promocaoFim && new Date(promocaoFim).getTime() <= new Date(promocaoInicio).getTime()) {
      throw new Error("O encerramento deve ser posterior ao início da promoção.");
    }
    let ordem;

    if (indice >= 0) {
      // Editar produto não altera sua posição atual.
      ordem = ordens[indice] || indice + 1;
    } else {
      // Produto novo recebe uma ordem anterior a todos os demais.
      const validas = ordens.filter(n => Number.isFinite(n));
      ordem = validas.length ? Math.min.apply(null, validas) - 1 : 1;
    }

    const linha = [
      codigo,
      String(p.nome).trim(),
      String(p.categoria || ""),
      String(p.descricao || ""),
      preco,
      Number(p.estoque || 0),
      String(p.disponibilidade || "Consulte disponibilidade"),
      String(p.foto_principal || ""),
      String(p.fotos_adicionais || ""),
      Boolean(p.destaque),
      p.ativo !== false,
      new Date(),
      String(p.tamanhos || ""),
      String(p.video_url || ""),
      preco,
      ordem,
      precoPromocional,
      promocaoInicio,
      promocaoFim
    ];

    if (indice >= 0) {
      sh.getRange(indice + 2, 1, 1, linha.length).setValues([linha]);
    } else {
      sh.appendRow(linha);
    }

    return { codigo: codigo, ordem: ordem };
  } finally {
    trava.releaseLock();
  }
}

function ordenar_(codigosRecebidos) {
  let codigos;

  if (Array.isArray(codigosRecebidos)) {
    codigos = codigosRecebidos;
  } else {
    try {
      codigos = JSON.parse(String(codigosRecebidos || "[]"));
    } catch (err) {
      throw new Error("A lista de produtos para ordenar é inválida.");
    }
  }

  if (!Array.isArray(codigos) || codigos.length === 0) {
    throw new Error("Nenhum produto foi informado para ordenar.");
  }

  codigos = codigos.map(c => String(c).trim());
  if (codigos.some(c => !c)) throw new Error("Há um código de produto vazio na ordenação.");

  const normalizados = codigos.map(c => c.toLowerCase());
  if (new Set(normalizados).size !== normalizados.length) {
    throw new Error("A lista contém produtos repetidos.");
  }

  const trava = LockService.getScriptLock();
  trava.waitLock(20000);

  try {
    const sh = aba_();
    const ultima = sh.getLastRow();
    if (ultima < 2) throw new Error("Não há produtos cadastrados.");

    garantirOrdens_(sh);

    const dados = sh.getRange(2, 1, ultima - 1, 16).getValues();
    const mapa = new Map();

    dados.forEach((linha, indice) => {
      const chave = String(linha[0] || "").trim().toLowerCase();
      if (chave) mapa.set(chave, indice);
    });

    codigos.forEach(codigo => {
      if (!mapa.has(codigo.toLowerCase())) {
        throw new Error("Um dos produtos enviados para ordenação não foi encontrado: " + codigo);
      }
    });

    // Os códigos enviados vêm na ordem desejada. Produtos não enviados
    // (por exemplo, inativos) permanecem depois, mantendo a ordem anterior.
    const escolhidos = new Set(normalizados);
    const restantes = dados
      .map((linha, indice) => ({ linha: linha, indice: indice }))
      .filter(item => !escolhidos.has(String(item.linha[0] || "").trim().toLowerCase()))
      .sort((a, b) => Number(a.linha[15]) - Number(b.linha[15]) || a.indice - b.indice)
      .map(item => String(item.linha[0]).trim());

    const ordemFinal = codigos.concat(restantes);
    ordemFinal.forEach((codigo, indice) => {
      const linha = mapa.get(codigo.toLowerCase());
      dados[linha][15] = indice + 1;
    });

    // Grava a coluna de ordem em uma única operação, na ordem física das linhas.
    sh.getRange(2, 16, dados.length, 1).setValues(dados.map(linha => [Number(linha[15])]));

    return { ordenados: codigos.length, total: dados.length };
  } finally {
    trava.releaseLock();
  }
}

function excluir_(codigo) {
  if (!codigo) throw new Error("Código não informado.");

  const trava = LockService.getScriptLock();
  trava.waitLock(20000);

  try {
    const sh = aba_();
    const ultima = sh.getLastRow();

    if (ultima < 2) throw new Error("Produto não encontrado.");

    const codigos = sh.getRange(2, 1, ultima - 1, 1).getValues().flat();

    const indice = codigos.findIndex(c =>
      String(c).trim().toLowerCase() === String(codigo).trim().toLowerCase()
    );

    if (indice < 0) throw new Error("Produto não encontrado.");

    sh.deleteRow(indice + 2);

    return { codigo: codigo, excluido: true };
  } finally {
    trava.releaseLock();
  }
}


// ======================================================
// VENDEDORES PARCEIROS
// ======================================================

const COLUNAS_VENDEDORES = [
  "id","nome","telefone","instagram","status","codigo","comissao","criado_em","atualizado_em",
  "email","senha_hash","senha_salt","trocar_senha","senha_atualizada_em","chave_pix"
];
const COLUNAS_PEDIDOS_VENDEDORES = [
  "id","vendedor_id","codigo_vendedor","nome_vendedor","cliente","itens","total_produtos","frete",
  "total_pedido","comissao_percentual","comissao_valor","status","data_criacao","atualizado_em","observacoes","pagamento_id"
];
const COLUNAS_PAGAMENTOS_VENDEDORES=["id","vendedor_id","vendedor","periodo","data_pagamento","pedidos_ids","quantidade_vendas","valor_total","chave_pix","comprovante_url","arquivo_nome","criado_em"];

function abaVendedores_() {
  const ss=planilha_(); let sh=ss.getSheetByName("Vendedores");
  if(!sh){sh=ss.insertSheet("Vendedores");sh.getRange(1,1,1,COLUNAS_VENDEDORES.length).setValues([COLUNAS_VENDEDORES]);return sh;}
  if(sh.getLastRow()===0){sh.getRange(1,1,1,COLUNAS_VENDEDORES.length).setValues([COLUNAS_VENDEDORES]);return sh;}
  const antigos=COLUNAS_VENDEDORES.slice(0,9), atuais=sh.getRange(1,1,1,9).getValues()[0];
  if(antigos.some((v,i)=>atuais[i]!==v)) throw new Error("Os cabeçalhos iniciais da aba Vendedores não correspondem. Nenhum dado foi alterado.");
  for(let i=9;i<COLUNAS_VENDEDORES.length;i++){
    const col=i+1, atual=sh.getRange(1,col).getValue();
    if(atual===COLUNAS_VENDEDORES[i]) continue;
    if(atual!=="") throw new Error("A coluna "+col+" da aba Vendedores já possui outro cabeçalho. Nenhum dado foi alterado.");
    if(sh.getLastRow()>=2 && sh.getRange(2,col,sh.getLastRow()-1,1).getValues().some(l=>l[0]!==""&&l[0]!==null))
      throw new Error("A coluna "+col+" da aba Vendedores contém dados. Verifique antes de continuar.");
    sh.getRange(1,col).setValue(COLUNAS_VENDEDORES[i]);
  }
  return sh;
}
function abaPedidosVendedores_(){
  const ss=planilha_();let sh=ss.getSheetByName("PedidosVendedores");
  if(!sh){sh=ss.insertSheet("PedidosVendedores");sh.getRange(1,1,1,COLUNAS_PEDIDOS_VENDEDORES.length).setValues([COLUNAS_PEDIDOS_VENDEDORES]);return sh;}
  if(sh.getLastRow()===0){sh.getRange(1,1,1,COLUNAS_PEDIDOS_VENDEDORES.length).setValues([COLUNAS_PEDIDOS_VENDEDORES]);return sh;}
  const h=sh.getRange(1,1,1,15).getValues()[0];
  if(COLUNAS_PEDIDOS_VENDEDORES.slice(0,15).some((v,i)=>h[i]!==v)) throw new Error("Os cabeçalhos da aba PedidosVendedores não correspondem.");
  if(sh.getRange(1,16).getValue()==="") sh.getRange(1,16).setValue("pagamento_id");
  if(sh.getRange(1,16).getValue()!=="pagamento_id") throw new Error("A coluna 16 da aba PedidosVendedores já está ocupada.");
  return sh;
}
function cadastrarVendedor_(dados){
  const nome=String(dados.nome||"").trim(), telefone=String(dados.telefone||"").trim();
  const instagram=String(dados.instagram||"").trim(),email=String(dados.email||"").trim().toLowerCase();
  const senha=String(dados.senha||"");
  const digitos=telefone.replace(/\D/g,"");
  if(nome.length<3||nome.length>100)throw new Error("Informe um nome válido.");
  if(digitos.length<10||digitos.length>13)throw new Error("Informe um WhatsApp válido com DDD.");
  if(instagram.length>60)throw new Error("O Instagram informado é muito longo.");
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)||email.length>150)throw new Error("Informe um e-mail válido.");
  if(senha.length<8||senha.length>72)throw new Error("A senha deve ter entre 8 e 72 caracteres.");
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const sh=abaVendedores_(),last=sh.getLastRow();
    if(last>=2){
      const rows=sh.getRange(2,1,last-1,COLUNAS_VENDEDORES.length).getValues();
      if(rows.some(l=>String(l[2]||"").replace(/\D/g,"")===digitos))throw new Error("Este WhatsApp já possui um cadastro. Entre em contato com a Bela Modas.");
      if(rows.some(l=>String(l[9]||"").trim().toLowerCase()===email))throw new Error("Este e-mail já possui um cadastro.");
    }
    const id=Utilities.getUuid(),codigo="BM"+Utilities.getUuid().replace(/-/g,"").substring(0,8).toUpperCase(),agora=new Date();
    const salt=Utilities.getUuid()+Utilities.getUuid();
    const hash=hashSenhaVendedor_(senha,salt);
    sh.appendRow([id,nome,telefone,instagram,"pendente",codigo,0,agora,agora,email,hash,salt,false,agora]);
    return {id,codigo,status:"pendente"};
  }finally{lock.releaseLock();}
}
function listarVendedores_(){
  const sh=abaVendedores_(),last=sh.getLastRow();if(last<2)return [];
  return sh.getRange(2,1,last-1,COLUNAS_VENDEDORES.length).getValues().filter(l=>l[0] && String(l[4]||"").toLowerCase()!=="excluido").map(l=>({
    id:String(l[0]),nome:String(l[1]||""),telefone:String(l[2]||""),instagram:String(l[3]||""),
    status:String(l[4]||"pendente"),codigo:String(l[5]||""),comissao:Number(l[6]||0),
    criado_em:l[7] instanceof Date?l[7].toISOString():l[7],atualizado_em:l[8] instanceof Date?l[8].toISOString():l[8],
    email:String(l[9]||""),acessoConfigurado:Boolean(l[10]&&l[11]),trocarSenha:l[12]===true
  }));
}
function atualizarVendedor_(id,status,comissao){
  if(!id)throw new Error("Vendedor não informado.");
  const st=String(status||"").toLowerCase();
  if(st&&!["pendente","aprovado","reprovado","bloqueado"].includes(st))throw new Error("Status de vendedor inválido.");
  const tem=comissao!==undefined&&comissao!==null&&String(comissao)!=="",valor=tem?Number(comissao):null;
  if(tem&&(!Number.isFinite(valor)||valor<0||valor>100))throw new Error("A comissão deve estar entre 0 e 100%.");
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const sh=abaVendedores_(),last=sh.getLastRow();if(last<2)throw new Error("Vendedor não encontrado.");
    const ids=sh.getRange(2,1,last-1,1).getValues().flat(),i=ids.findIndex(v=>String(v)===String(id));
    if(i<0)throw new Error("Vendedor não encontrado.");
    const row=i+2;if(st)sh.getRange(row,5).setValue(st);if(tem)sh.getRange(row,7).setValue(valor);sh.getRange(row,9).setValue(new Date());
    return {id:String(id),status:st||String(sh.getRange(row,5).getValue()),comissao:Number(sh.getRange(row,7).getValue()||0)};
  }finally{lock.releaseLock();}
}
function excluirVendedor_(id){
  if(!id)throw new Error("Vendedor não informado.");
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const x=localizarVendedorPorId_(id);
    if(String(x.dados[4]||"").toLowerCase()==="excluido")throw new Error("Este vendedor já foi excluído.");
    // Exclusão lógica: preserva pedidos/comissões e impede novas sessões.
    x.sh.getRange(x.linha,5).setValue("excluido");
    x.sh.getRange(x.linha,9).setValue(new Date());
    x.sh.getRange(x.linha,11,1,4).setValues([["","",true,new Date()]]);
    return {id:String(id),excluido:true,historicoPreservado:true};
  }finally{lock.releaseLock();}
}

function hashSenhaVendedor_(senha,salt){
  let v=String(salt)+":"+String(senha);
  for(let i=0;i<5000;i++)v=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,v,Utilities.Charset.UTF_8)
    .map(b=>("0"+((b+256)%256).toString(16)).slice(-2)).join("");
  return v;
}
function localizarVendedorPorId_(id){
  const sh=abaVendedores_(),last=sh.getLastRow();if(last<2)throw new Error("Vendedor não encontrado.");
  const rows=sh.getRange(2,1,last-1,COLUNAS_VENDEDORES.length).getValues(),i=rows.findIndex(l=>String(l[0])===String(id));
  if(i<0)throw new Error("Vendedor não encontrado.");return {sh,linha:i+2,dados:rows[i]};
}
function criarSessaoVendedor_(emailRecebido,senha){
  const email=String(emailRecebido||"").trim().toLowerCase(),pass=String(senha||"");
  if(!email||!pass)throw new Error("Informe e-mail e senha.");
  const cache=CacheService.getScriptCache(),chave="tentativas_vendedor_"+Utilities.base64EncodeWebSafe(email).substring(0,80),tent=Number(cache.get(chave)||0);
  if(tent>=CONFIG.TENTATIVAS_MAX)throw new Error("Muitas tentativas. Aguarde alguns minutos.");
  const sh=abaVendedores_(),last=sh.getLastRow();if(last<2)throw new Error("E-mail ou senha incorretos.");
  const rows=sh.getRange(2,1,last-1,COLUNAS_VENDEDORES.length).getValues(),i=rows.findIndex(l=>String(l[9]||"").trim().toLowerCase()===email);
  if(i<0){cache.put(chave,String(tent+1),600);throw new Error("E-mail ou senha incorretos.");}
  const v=rows[i];
  if(String(v[4]||"").toLowerCase()!=="aprovado")throw new Error("Seu cadastro ainda não está aprovado ou está bloqueado.");
  if(!v[10]||!v[11])throw new Error("Seu acesso ainda não foi criado. Entre em contato com a Bela Modas para receber uma senha temporária.");
  if(hashSenhaVendedor_(pass,String(v[11]))!==String(v[10])){cache.put(chave,String(tent+1),600);throw new Error("E-mail ou senha incorretos.");}
  cache.remove(chave);const token=Utilities.getUuid()+Utilities.getUuid();cache.put("sessao_vendedor_"+token,String(v[0]),CONFIG.SESSAO_SEGUNDOS);
  return {token,nome:String(v[1]||""),trocarSenha:v[12]===true,validadeSegundos:CONFIG.SESSAO_SEGUNDOS};
}
function validarSessaoVendedor_(token){
  if(!token||typeof token!=="string"||token.length>100)throw new Error("Sessão do vendedor inválida. Entre novamente.");
  const cache=CacheService.getScriptCache(),id=cache.get("sessao_vendedor_"+token);if(!id)throw new Error("Sessão expirada. Entre novamente.");
  const v=localizarVendedorPorId_(id).dados;if(String(v[4]||"").toLowerCase()!=="aprovado"){cache.remove("sessao_vendedor_"+token);throw new Error("Seu acesso não está liberado.");}
  return String(id);
}
function trocarSenhaVendedor_(id,novaSenha){
  const senha=String(novaSenha||"");if(senha.length<8||senha.length>72)throw new Error("A nova senha deve ter entre 8 e 72 caracteres.");
  const x=localizarVendedorPorId_(id),salt=Utilities.getUuid()+Utilities.getUuid();
  x.sh.getRange(x.linha,11,1,4).setValues([[hashSenhaVendedor_(senha,salt),salt,false,new Date()]]);
  return {alterada:true};
}
function resetarSenhaVendedor_(id){
  if(!id)throw new Error("Vendedor não informado.");const x=localizarVendedorPorId_(id);
  if(String(x.dados[4]||"").toLowerCase()!=="aprovado")throw new Error("Somente vendedores aprovados podem receber acesso.");
  const temp=Utilities.getUuid().replace(/-/g,"").substring(0,12).toUpperCase(),salt=Utilities.getUuid()+Utilities.getUuid();
  x.sh.getRange(x.linha,11,1,4).setValues([[hashSenhaVendedor_(temp,salt),salt,true,new Date()]]);
  return {senhaTemporaria:temp,email:String(x.dados[9]||"")};
}
function normalizarStatusPedido_(status){const s=String(status||"aguardando").toLowerCase();return ({confirmada:"confirmando",concluida:"finalizado"})[s]||s;}
function salvarChavePixVendedor_(id,chave){const v=String(chave||"").trim();if(v.length<3||v.length>150)throw new Error("Informe uma chave Pix válida (até 150 caracteres).");const x=localizarVendedorPorId_(id);x.sh.getRange(x.linha,15).setValue(v);x.sh.getRange(x.linha,9).setValue(new Date());return {salva:true};}
function abaPagamentosVendedores_(){const ss=planilha_();let sh=ss.getSheetByName("PagamentosVendedores");if(!sh){sh=ss.insertSheet("PagamentosVendedores");sh.getRange(1,1,1,COLUNAS_PAGAMENTOS_VENDEDORES.length).setValues([COLUNAS_PAGAMENTOS_VENDEDORES]);}else if(sh.getLastRow()===0)sh.getRange(1,1,1,COLUNAS_PAGAMENTOS_VENDEDORES.length).setValues([COLUNAS_PAGAMENTOS_VENDEDORES]);else{const h=sh.getRange(1,1,1,COLUNAS_PAGAMENTOS_VENDEDORES.length).getValues()[0];if(COLUNAS_PAGAMENTOS_VENDEDORES.some((v,i)=>h[i]!==v))throw new Error("Cabeçalhos da aba PagamentosVendedores inválidos.");}return sh;}
function listarPagamentosVendedores_(){const p=abaPagamentosVendedores_(),lp=p.getLastRow(),pagamentos=lp<2?[]:p.getRange(2,1,lp-1,COLUNAS_PAGAMENTOS_VENDEDORES.length).getValues().map(l=>({id:String(l[0]),vendedorId:String(l[1]),vendedor:String(l[2]),periodo:String(l[3]),dataPagamento:l[4] instanceof Date?l[4].toISOString():l[4],quantidade:Number(l[6]||0),valor:Number(l[7]||0),chavePix:String(l[8]||""),comprovanteUrl:String(l[9]||""),arquivoNome:String(l[10]||"")}));const pedidos=listarPedidosVendedores_().filter(p=>p.status==="finalizado"&&!p.pagamentoId).map(p=>{p.chavePix=String(localizarVendedorPorId_(p.vendedorId).dados[14]||"");return p;});return {pedidos,pagamentos};}
function registrarPagamentoVendedor_(d){const lock=LockService.getScriptLock();lock.waitLock(20000);let arquivo=null;try{const vendedorId=String(d.vendedorId||""),ids=Array.isArray(d.pedidoIds)?d.pedidoIds.map(String):[],nome=String(d.nomeArquivo||"comprovante").slice(0,120),tipo=String(d.tipo||""),base64=String(d.base64||"");if(!vendedorId||!ids.length||new Set(ids).size!==ids.length)throw new Error("Selecione vendas válidas para pagar.");if(!["application/pdf","image/jpeg","image/png"].includes(tipo))throw new Error("Comprovante deve ser PDF, JPG ou PNG.");if(!base64||base64.length>4.2*1024*1024)throw new Error("Comprovante muito grande. Limite de 3 MB.");const x=localizarVendedorPorId_(vendedorId),sh=abaPedidosVendedores_(),last=sh.getLastRow();if(last<2)throw new Error("Não há vendas para pagamento.");const rows=sh.getRange(2,1,last-1,COLUNAS_PEDIDOS_VENDEDORES.length).getValues(),sel=rows.map((l,i)=>({l,i})).filter(o=>ids.includes(String(o.l[0])));if(sel.length!==ids.length||sel.some(o=>String(o.l[1])!==vendedorId||normalizarStatusPedido_(o.l[11])!=="finalizado"||o.l[15]))throw new Error("Uma ou mais vendas já foram pagas ou não pertencem a este vendedor.");const total=Math.round(sel.reduce((s,o)=>s+Number(o.l[10]||0),0)*100)/100;if(total<=0)throw new Error("Total de comissão inválido.");const agora=new Date(),dia=agora.getDate(),ini=dia<=15?1:16,fim=dia<=15?15:new Date(agora.getFullYear(),agora.getMonth()+1,0).getDate(),periodo=Utilities.formatDate(agora,Session.getScriptTimeZone(),"MM/yyyy")+" · "+ini+" a "+fim;const pai=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty(CONFIG.PASTA)),fs=pai.getFoldersByName("ComprovantesPagamentosVendedores"),pasta=fs.hasNext()?fs.next():pai.createFolder("ComprovantesPagamentosVendedores");arquivo=pasta.createFile(Utilities.newBlob(Utilities.base64Decode(base64),tipo,nome));arquivo.setName(Utilities.getUuid()+"_"+nome);try{arquivo.addViewer(String(x.dados[9]||""));}catch(err){arquivo.setTrashed(true);throw new Error("Não foi possível compartilhar o comprovante com o e-mail do vendedor. Confira se ele pode receber arquivos do Google Drive.");}const pid=Utilities.getUuid(),pedidos=sel.map(o=>String(o.l[0]));abaPagamentosVendedores_().appendRow([pid,vendedorId,String(x.dados[1]||""),periodo,agora,JSON.stringify(pedidos),pedidos.length,total,String(x.dados[14]||""),arquivo.getUrl(),nome,agora]);sel.forEach(o=>sh.getRange(o.i+2,16).setValue(pid));return {registrado:true,id:pid,total,quantidade:pedidos.length,comprovanteUrl:arquivo.getUrl(),periodo};}catch(err){if(arquivo)try{arquivo.setTrashed(true);}catch(_){}throw err;}finally{lock.releaseLock();}}
function obterPainelVendedor_(id){
  const x=localizarVendedorPorId_(id),v=x.dados,sh=abaPedidosVendedores_(),last=sh.getLastRow();
  const pedidos=last<2?[]:sh.getRange(2,1,last-1,COLUNAS_PEDIDOS_VENDEDORES.length).getValues().filter(l=>String(l[1])===String(id)).map(l=>({
    id:String(l[0]),cliente:String(l[4]||""),itens:String(l[5]||"[]"),totalProdutos:Number(l[6]||0),frete:Number(l[7]||0),
    total:Number(l[8]||0),comissaoPercentual:Number(l[9]||0),comissao:Number(l[10]||0),status:normalizarStatusPedido_(l[11]),pagamentoId:String(l[15]||""),statusComissao:normalizarStatusPedido_(l[11])!=="finalizado"?"Não liberada":(l[15]?"Paga":"Aguardando pagamento"),
    criadoEm:l[12] instanceof Date?l[12].toISOString():l[12],observacoes:String(l[14]||"")
  })).sort((a,b)=>new Date(b.criadoEm).getTime()-new Date(a.criadoEm).getTime());
  const concluidos=pedidos.filter(p=>p.status==="finalizado"),naoPagos=concluidos.filter(p=>!p.pagamentoId),pag=abaPagamentosVendedores_(),lp=pag.getLastRow();
  const historico=lp<2?[]:pag.getRange(2,1,lp-1,COLUNAS_PAGAMENTOS_VENDEDORES.length).getValues().filter(l=>String(l[1])===String(id)).map(l=>({id:String(l[0]),periodo:String(l[3]),dataPagamento:l[4] instanceof Date?l[4].toISOString():l[4],quantidade:Number(l[6]||0),valor:Number(l[7]||0),comprovanteUrl:String(l[9]||"")}));
  return {vendedor:{nome:String(v[1]||""),email:String(v[9]||""),codigo:String(v[5]||""),comissao:Number(v[6]||0),chavePix:String(v[14]||"")},pedidos,historico,resumo:{totalPedidos:pedidos.length,vendasConcluidas:concluidos.length,vendasCanceladas:pedidos.filter(p=>p.status==="cancelada").length,valorVendasConcluidas:concluidos.reduce((s,p)=>s+p.totalProdutos,0),comissaoLiberada:concluidos.reduce((s,p)=>s+p.comissao,0),comissaoPendente:naoPagos.reduce((s,p)=>s+p.comissao,0),comissaoPaga:concluidos.filter(p=>p.pagamentoId).reduce((s,p)=>s+p.comissao,0)}};
}
function registrarPedidoVendedor_(pedido){
  const codigo=String(pedido.codigoVendedor||"").trim().toUpperCase();if(!codigo)return {registrado:false,motivo:"sem_vendedor"};
  const recebidos=Array.isArray(pedido.itens)?pedido.itens:[];if(!recebidos.length||recebidos.length>50)throw new Error("Os itens do pedido são inválidos.");
  const cliente=String(pedido.cliente||"Cliente").trim().slice(0,100),obs=String(pedido.observacoes||"").trim().slice(0,500),frete=Number(pedido.frete||0);
  if(!Number.isFinite(frete)||frete<0||frete>5000)throw new Error("Valor de entrega inválido.");
  const sv=abaVendedores_(),lv=sv.getLastRow();if(lv<2)return {registrado:false,motivo:"vendedor_nao_encontrado"};
  const vendedor=sv.getRange(2,1,lv-1,COLUNAS_VENDEDORES.length).getValues().find(l=>String(l[5]||"").toUpperCase()===codigo);
  if(!vendedor||String(vendedor[4]||"").toLowerCase()!=="aprovado")return {registrado:false,motivo:"vendedor_nao_aprovado"};
  const sp=aba_(),lp=sp.getLastRow();if(lp<2)throw new Error("Não foi possível validar os produtos do pedido.");
  const produtos=sp.getRange(2,1,lp-1,CONFIG.COLUNAS.length).getValues(),mapa=new Map();produtos.forEach(l=>mapa.set(String(l[0]||"").trim().toLowerCase(),l));
  const agora=new Date(),itens=[];let totalProdutos=0;
  recebidos.forEach(item=>{
    const cod=String(item.codigo||"").trim().toLowerCase(),q=Number(item.quantidade||0);
    if(!cod||!Number.isInteger(q)||q<1||q>50)throw new Error("Quantidade ou produto inválido no pedido.");
    const p=mapa.get(cod);if(!p||p[10]===false)throw new Error("Um produto do pedido não está mais disponível: "+cod);
    let preco=Number(p[14]||p[4]||0),promo=Number(p[16]||0),ini=p[17]?new Date(p[17]):null,fim=p[18]?new Date(p[18]):null;
    if(promo>0&&promo<preco&&(!ini||agora>=ini)&&(!fim||agora<=fim))preco=promo;
    if(!Number.isFinite(preco)||preco<0)throw new Error("Preço inválido para o produto "+cod);
    const subtotal=Math.round(preco*q*100)/100;totalProdutos+=subtotal;
    itens.push({codigo:String(p[0]),nome:String(p[1]||""),tamanho:String(item.tamanho||"").slice(0,30),quantidade:q,precoUnitario:preco,subtotal});
  });
  totalProdutos=Math.round(totalProdutos*100)/100;const taxa=Number(vendedor[6]||0),comissao=Math.round(totalProdutos*taxa)/100;
  const id=String(pedido.id||Utilities.getUuid()).replace(/[^A-Za-z0-9_-]/g,"").slice(0,60);if(!id)throw new Error("Identificador do pedido inválido.");
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const sh=abaPedidosVendedores_(),last=sh.getLastRow();
    if(last>=2&&sh.getRange(2,1,last-1,1).getValues().flat().some(v=>String(v)===id))return {registrado:true,duplicado:true,id};
    const recebimento=String(pedido.recebimento||"").slice(0,30);
    sh.appendRow([id,String(vendedor[0]),codigo,String(vendedor[1]),cliente,JSON.stringify(itens),totalProdutos,frete,
      Math.round((totalProdutos+frete)*100)/100,taxa,comissao,"aguardando",agora,agora,obs+(recebimento?" | Recebimento: "+recebimento:"")]);
    return {registrado:true,id,totalProdutos,comissao};
  }finally{lock.releaseLock();}
}
function listarPedidosVendedores_(){
  const sh=abaPedidosVendedores_(),last=sh.getLastRow();if(last<2)return [];
  return sh.getRange(2,1,last-1,COLUNAS_PEDIDOS_VENDEDORES.length).getValues().filter(l=>l[0]).map(l=>({
    id:String(l[0]),vendedorId:String(l[1]),codigoVendedor:String(l[2]),vendedor:String(l[3]),cliente:String(l[4]),itens:String(l[5]||"[]"),
    totalProdutos:Number(l[6]||0),frete:Number(l[7]||0),total:Number(l[8]||0),comissaoPercentual:Number(l[9]||0),comissao:Number(l[10]||0),
    status:normalizarStatusPedido_(l[11]),pagamentoId:String(l[15]||""),criadoEm:l[12] instanceof Date?l[12].toISOString():l[12],observacoes:String(l[14]||"")
  })).sort((a,b)=>new Date(b.criadoEm).getTime()-new Date(a.criadoEm).getTime());
}
function atualizarPedidoVendedor_(id,status){
  const recebido=String(status||"").toLowerCase(),st=({confirmada:"confirmando",concluida:"finalizado"})[recebido]||recebido;if(!id||!["aguardando","confirmando","finalizado","cancelada"].includes(st))throw new Error("Status do pedido inválido.");
  const sh=abaPedidosVendedores_(),last=sh.getLastRow();if(last<2)throw new Error("Pedido não encontrado.");
  const ids=sh.getRange(2,1,last-1,1).getValues().flat(),i=ids.findIndex(v=>String(v)===String(id));if(i<0)throw new Error("Pedido não encontrado.");
  sh.getRange(i+2,12).setValue(st);sh.getRange(i+2,14).setValue(new Date());return {id:String(id),status:st};
}

// ======================================================
// CONFIGURAÇÕES DE ENTREGA
// ======================================================

function obterConfiguracaoEntregas_() {
  const props = PropertiesService.getScriptProperties();
  return {
    valorKm: Number(props.getProperty("ENTREGA_VALOR_KM") || 2),
    enderecoOrigem: props.getProperty("ENTREGA_ENDERECO_ORIGEM") || "Avenida México, 87, Petrovale, Betim - MG, CEP 32668-052",
    enderecoRetirada: props.getProperty("ENTREGA_ENDERECO_RETIRADA") || "Avenida México, 87, Petrovale, Betim - MG, CEP 32668-052"
  };
}

function calcularEntrega_(destinoRecebido) {
  const config = obterConfiguracaoEntregas_();
  const origem = String(config.enderecoOrigem || "").trim();
  const destino = String(destinoRecebido || "").trim();

  if (!origem) throw new Error("O endereço de origem da entrega não está configurado.");
  if (destino.length < 12 || destino.length > 300) {
    throw new Error("Informe um endereço completo, incluindo rua, número, bairro, cidade e CEP.");
  }

  const valorKm = Number(config.valorKm);
  if (!Number.isFinite(valorKm) || valorKm < 0 || valorKm > 1000) {
    throw new Error("O valor por quilômetro configurado é inválido. Confira o painel administrativo.");
  }

  // Guarda a distância da rota por até 6 horas. O preço é recalculado sempre
  // com a tarifa atual, portanto mudar o valor por km no painel tem efeito imediato.
  const chaveTexto = origem.toLocaleLowerCase("pt-BR") + "|" + destino.toLocaleLowerCase("pt-BR");
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    chaveTexto,
    Utilities.Charset.UTF_8
  ).map(b => ("0" + ((b + 256) % 256).toString(16)).slice(-2)).join("");
  const cache = CacheService.getScriptCache();
  const chaveCache = "rota_entrega_" + hash;
  let rota = null;
  const salva = cache.get(chaveCache);
  if (salva) {
    try { rota = JSON.parse(salva); } catch (_) { rota = null; }
  }

  if (!rota || !Number.isFinite(Number(rota.distanciaMetros))) {
    let direcoes;
    try {
      direcoes = Maps.newDirectionFinder()
        .setOrigin(origem)
        .setDestination(destino)
        .setMode(Maps.DirectionFinder.Mode.DRIVING)
        .setLanguage("pt-BR")
        .setRegion("br")
        .getDirections();
    } catch (err) {
      throw new Error("Não foi possível consultar a rota agora. Confira o endereço e tente novamente.");
    }

    const rotas = direcoes && Array.isArray(direcoes.routes) ? direcoes.routes : [];
    const pernas = rotas.length && Array.isArray(rotas[0].legs) ? rotas[0].legs : [];
    if (!pernas.length) {
      throw new Error("Não encontramos uma rota de carro para esse endereço. Confira o CEP e o número.");
    }

    const distanciaMetros = pernas.reduce((soma, perna) => soma + Number(perna.distance && perna.distance.value || 0), 0);
    const duracaoSegundos = pernas.reduce((soma, perna) => soma + Number(perna.duration && perna.duration.value || 0), 0);
    if (!Number.isFinite(distanciaMetros) || distanciaMetros <= 0 || distanciaMetros > 1000000) {
      throw new Error("A distância retornada para essa entrega é inválida ou está fora da área atendida.");
    }
    rota = { distanciaMetros: distanciaMetros, duracaoSegundos: duracaoSegundos };
    cache.put(chaveCache, JSON.stringify(rota), 21600);
  }

  const distanciaKm = Number(rota.distanciaMetros) / 1000;
  // Cobrança proporcional: não arredonda a distância para quilômetros inteiros.
  const frete = Math.round(distanciaKm * valorKm * 100) / 100;

  return {
    origem: origem,
    destino: destino,
    distanciaKm: distanciaKm,
    valorKm: valorKm,
    frete: frete,
    duracaoMinutos: Math.ceil(Number(rota.duracaoSegundos || 0) / 60),
    metodo: "rota_de_carro",
    calculo: "proporcional"
  };
}

function salvarConfiguracaoEntregas_(config) {
  const valorKm = Number(config.valorKm);
  const enderecoOrigem = String(config.enderecoOrigem || "").trim();
  const enderecoRetirada = String(config.enderecoRetirada || "").trim();
  if (!Number.isFinite(valorKm) || valorKm < 0 || valorKm > 1000) {
    throw new Error("Informe um valor por quilômetro entre R$ 0 e R$ 1.000.");
  }
  if (!enderecoOrigem || enderecoOrigem.length > 250) {
    throw new Error("Informe um endereço de origem válido.");
  }
  if (!enderecoRetirada || enderecoRetirada.length > 250) {
    throw new Error("Informe um endereço de retirada válido.");
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    ENTREGA_VALOR_KM: String(valorKm),
    ENTREGA_ENDERECO_ORIGEM: enderecoOrigem,
    ENTREGA_ENDERECO_RETIRADA: enderecoRetirada
  }, false);
  return obterConfiguracaoEntregas_();
}

function midia_(p) {
  if (!p.codigo || !p.nomeArquivo || !p.base64) {
    throw new Error("Dados do arquivo incompletos.");
  }

  const pastaId = PropertiesService.getScriptProperties()
    .getProperty(CONFIG.PASTA);

  if (!pastaId) throw new Error("Pasta de arquivos não configurada.");

  const isVideo = p.tipoMidia === "video";

  const tipos = isVideo
    ? ["video/mp4", "video/webm"]
    : ["image/jpeg", "image/png", "image/webp", "image/gif"];

  if (!tipos.includes(p.tipo)) {
    throw new Error("Formato de arquivo não permitido.");
  }

  const bytes = Utilities.base64Decode(p.base64);
  const limite = isVideo ? 8 * 1024 * 1024 : 5 * 1024 * 1024;

  if (bytes.length > limite) {
    throw new Error("O arquivo excede o limite permitido.");
  }

  const blob = Utilities.newBlob(bytes, p.tipo, p.nomeArquivo);
  const arquivo = DriveApp.getFolderById(pastaId).createFile(blob);

  arquivo.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return {
    codigo: p.codigo,
    arquivoId: arquivo.getId(),
    url: isVideo
      ? "https://drive.google.com/uc?export=download&id=" +
        encodeURIComponent(arquivo.getId())
      : "https://drive.google.com/thumbnail?id=" +
        encodeURIComponent(arquivo.getId()) + "&sz=w1200"
  };
}

// ======================================================
// CONFIGURAÇÃO DO VÍDEO DO BANNER PRINCIPAL
// ======================================================

function obterBanner_() {
  const url = PropertiesService.getScriptProperties()
    .getProperty("BANNER_VIDEO_URL") || "";

  return {
    video_url: url
  };
}

function salvarBanner_(url) {
  const valor = String(url || "").trim();

  if (valor && !/^https:\/\/drive\.google\.com\/uc\?export=download&id=[A-Za-z0-9_-]+$/.test(valor)) {
    throw new Error("Endereço do vídeo do banner inválido.");
  }

  PropertiesService.getScriptProperties()
    .setProperty("BANNER_VIDEO_URL", valor);

  return {
    salvo: true,
    video_url: valor
  };
}