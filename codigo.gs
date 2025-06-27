// ------------------------------
// Configurações gerais
// ------------------------------
const ALLOWED_SENDERS = [
  'email1@email.com.br',
  'email2@email.com.br',
];
const NFE_NAMESPACE = XmlService.getNamespace('http://www.portalfiscal.inf.br/nfe');
const SISATEC_NAMESPACE = XmlService.getNamespace('http://schemas.datacontract.org/2004/07/WS.Entity');


// ------------------------------
// 1️⃣ Função principal: busca novos e-mails com XML e atualiza
// ------------------------------
function parseNewEmailsAndPopulateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Automático do e-mail");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Aba 'Automático do e-mail' não encontrada!");
    return;
  }

  const query = `is:unread has:attachment filename:xml (${ALLOWED_SENDERS.map(e => 'from:' + e).join(' OR ')})`;
  const threads = GmailApp.search(query);
  const newRows = [];

  threads.forEach(thread =>
    thread.getMessages().forEach(message => {
      const sender = message.getFrom().match(/<(.+?)>/)?.[1] || message.getFrom();
      if (ALLOWED_SENDERS.includes(sender)) {
        // extrai e prepara linhas
        processMessage(message, NFE_NAMESPACE, newRows);
        message.markRead();
      }
    })
  );

  if (newRows.length) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }

  // depois de importar e-mails, importa Sisatec e atualiza notas
  buscarAbastecimentosSisatecXML();
  atualizarNotas();
  enviarConfirmacoesDeNota();
}


// ------------------------------
// 2️⃣ Processa cada mensagem XML e acumula em rows[]
// ------------------------------
function processMessage(message, ns, rows) {
  message.getAttachments().forEach(att => {
    if (!att.getName().endsWith('.xml')) return;

    try {
      const xmlDoc = XmlService.parse(att.getDataAsString());
      const root = xmlDoc.getRootElement();
      // suporta <NFe> direto ou aninhado em <nfeProc>
      const nfeElem = root.getChild('NFe', ns)
        || root.getChild('nfeProc')?.getChild('NFe', ns);
      if (!nfeElem) throw new Error('Elemento NFe não encontrado');

      const infNFe = nfeElem.getChild('infNFe', ns);
      const ide = infNFe.getChild('ide', ns);
      const dhEmi = ide.getChild('dhEmi', ns).getText();
      const det = infNFe.getChild('det', ns);
      const prod = det.getChild('prod', ns);
      const pag = infNFe.getChild('pag', ns)
        .getChild('detPag', ns)
        .getChild('card', ns);
      const infAdic = infNFe.getChild('infAdic', ns);
      // Captura código de autorização do cartão
      let cAut = '';
      try { cAut = pag.getChild('cAut', ns).getText(); } catch { cAut = 'not found'; }

      // extrai dados adicionais (abnum, placa, odômetro)
      const { abnum, placa, odometro } = extractAdditionalInfo(infAdic, ns);

      // monta linha: [ABNUM, Data(Date), Série, Nota, Emitente, Qtde, Valor]
      const row = [
        cAut,
        new Date(dhEmi),
        infNFe.getChild('ide', ns).getChild('serie', ns).getText(),
        infNFe.getChild('ide', ns).getChild('nNF', ns).getText(),
        infNFe.getChild('emit', ns).getChild('xNome', ns).getText(),
        parseFloat(prod.getChild('qCom', ns).getText()),
        parseFloat(
          infNFe
            .getChild('pag', ns)
            .getChild('detPag', ns)
            .getChild('vPag', ns)
            .getText()
        )
      ];

      rows.push(row);
    } catch (e) {
      Logger.log(`Erro no XML (${att.getName()}): ${e}`);
    }
  });
}

// extrai abnum, placa, odômetro de infAdic -- não usada
function extractAdditionalInfo(infAdic, ns) {
  const info = { abnum: '', placa: '', odometro: '' };
  if (!infAdic) return info;

  infAdic.getChildren('obsCont', ns).forEach(obs => {
    const campo = obs.getAttribute('xCampo')?.getValue();
    const texto = obs.getChild('xTexto', ns)?.getText();
    if (campo === 'abnum') info.abnum = texto;
    if (campo === 'placa') info.placa = texto;
    if (campo === 'odometro') info.odometro = texto;
  });
  return info;
}


// ------------------------------
// 3️⃣ Importa Sisatec em batch
// ------------------------------
function buscarAbastecimentosSisatecXML() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("TABELA ABASTECIMENTO");
  if (!aba) {
    SpreadsheetApp.getUi().alert("Aba 'TABELA ABASTECIMENTO' não encontrada!");
    return;
  }

  // configura datas
  const hoje = new Date();
  const fim = Utilities.formatDate(hoje, 'GMT-3', 'MM-dd-yyyy');
  const ini = (() => {
    const d = new Date(hoje);
    d.setDate(d.getDate() - 4);
    return Utilities.formatDate(d, 'GMT-3', 'MM-dd-yyyy');
  })();

  const url = `https://ws.sisatec.com.br/api/abastecimento/dataeunidade?` +
    `codigo=9438&key=YOURKEYHERE` +
    `&dataInicio=${ini}&dataFim=${fim}` +
    `&codigoUnidade=126&status=1&type=xml`;

  let resp;
  try {
    resp = UrlFetchApp.fetch(url);
  } catch (e) {
    Logger.log(`Erro na requisição Sisatec: ${e}`);
    return;
  }

  const xml = XmlService.parse(resp.getContentText());
  const root = xml.getRootElement();
  const itens = root
    .getChild('abastecimentos', root.getNamespace())
    .getChildren('Abastecimento', SISATEC_NAMESPACE);
  if (!itens || !itens.length) {
    Logger.log("Nenhum abastecimento Sisatec retornado.");
    return;
  }

  // monta Set de códigos existentes (numérico)
  const existentes = new Set(
    aba
      .getRange(8, 1, aba.getLastRow() - 7, 1)
      .getValues()
      .flat()
      .filter(c => c !== '' && c != null)
      .map(Number)
  );

  const novas = [];
  itens.forEach(item => {
    const codVenda = Number(item.getChildText('codAbastecimento', SISATEC_NAMESPACE));
    if (existentes.has(codVenda)) return;

    // converte data para Date
    const dataObj = converterDataSisatec(item.getChildText('data', SISATEC_NAMESPACE));

    // prepara linha com 40 colunas
    const linha = [
      codVenda,
      "", "",                // B, C
      dataObj,               // D
      item.getChildText('hora', SISATEC_NAMESPACE),
      item.getChildText('prefixo', SISATEC_NAMESPACE),
      item.getChildText('placa', SISATEC_NAMESPACE),
      "PMES",
      item.getChildText('centroDeCustoVeiculo', SISATEC_NAMESPACE),
      item.getChildText('combustivel', SISATEC_NAMESPACE),
      item.getChildText('condutor', SISATEC_NAMESPACE),
      item.getChildText('matricula_condutor', SISATEC_NAMESPACE),
      item.getChildText('registroCondutor', SISATEC_NAMESPACE),
      item.getChildText('posto', SISATEC_NAMESPACE),
      item.getChildText('estado', SISATEC_NAMESPACE),
      item.getChildText('endereco', SISATEC_NAMESPACE),
      item.getChildText('Cidade', SISATEC_NAMESPACE),
      item.getChildText('TipoFrota', SISATEC_NAMESPACE) || "",
      item.getChildText('marca', SISATEC_NAMESPACE),
      item.getChildText('modelo', SISATEC_NAMESPACE),
      "",                     // U
      item.getChildText('ano_veiculo', SISATEC_NAMESPACE),
      "",                     // W
      item.getChildText('CNPJ', SISATEC_NAMESPACE),
      "",                     // Y
      item.getChildText('kmAtual', SISATEC_NAMESPACE),
      parseFloat(item.getChildText('quantidadeLitros', SISATEC_NAMESPACE) || 0),
      parseFloat(item.getChildText('valorLitro', SISATEC_NAMESPACE) || 0),
      "",                     // AD?
      // col 30 = valor bruto formatado
      "R$" + parseFloat(item.getChildText('valor', SISATEC_NAMESPACE).replace(',', '.'))
        .toFixed(2).replace('.', ','),
      0,                      // col 31
      item.getChildText('KmHoraPorLitro', SISATEC_NAMESPACE),
      item.getChildText('KmHoraRodado', SISATEC_NAMESPACE),
      "",                     // 34
      "POS/TEF",              // 35
      "Cartão",               // 36
      "Não",                  // 37
      "PMES",                 // 38
      "PMES - 10ª CIA INDEPEND", // 39
      ""                      // 40 = controle e-mail
    ];
    novas.push(linha);
  });

  if (novas.length) {
    aba
      .getRange(aba.getLastRow() + 1, 1, novas.length, novas[0].length)
      .setValues(novas);
  }
}

// converte "DD/MM/YYYY" em Date (meio-dia)
function converterDataSisatec(txt) {
  const p = txt.split('/');
  return new Date(+p[2], +p[1] - 1, +p[0], 12, 0, 0);
}


// ------------------------------
// 4️⃣ Atualiza notas na TABELA ABASTECIMENTO
// ------------------------------
function atualizarNotas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAbas = ss.getSheetByName("TABELA ABASTECIMENTO");
  const abaEmail = ss.getSheetByName("Automático do e-mail");
  const abaCadNota = ss.getSheetByName("CAD_NOTA");
  if (!abaAbas || !abaEmail || !abaCadNota) {
    SpreadsheetApp.getUi().alert("Uma das abas não foi encontrada!");
    return;
  }

  // constantes de índice (coluna → array index = col - 1)
  const C_COD = 1, C_NOTA = 2, C_SERIE = 3, C_DATA = 4;
  const C_VAL_BRUTO = 30;

  const linhaIni = 8;
  const ultLinha = abaAbas.getLastRow();
  const dadosAbast = abaAbas
    .getRange(linhaIni, 1, ultLinha - linhaIni + 1, C_VAL_BRUTO)
    .getValues();
  const dadosEmail = abaEmail
    .getRange(2, 1, abaEmail.getLastRow() - 1, 7)
    .getValues();
  const dadosCadNota = abaCadNota
    .getRange(2, 1, abaCadNota.getLastRow() - 1, 3)
    .getValues();

  // cria mapas de busca
  const mapaEmail = {};
  dadosEmail.forEach(([cod, data, serie, nota, , , valor]) => {
    if (cod) mapaEmail[cod.toString().trim()] = { nota, serie, data, valor };
  });
  const mapaCadNota = {};
  dadosCadNota.forEach(([cod, nota, classe]) => {
    if (cod) mapaCadNota[cod.toString().trim()] = { nota, classe };
  });

  // percorre cada linha
  dadosAbast.forEach((linha, i) => {
    const lp = linhaIni + i;
    const cod = linha[C_COD - 1];
    if (!cod) return;
    const key = cod.toString().trim();

    // 1️⃣ CAD_NOTA
    if (mapaCadNota[key]) {
      const { nota, classe } = mapaCadNota[key];
      abaAbas.getRange(lp, C_NOTA).setValue(nota);
      abaAbas.getRange(lp, C_SERIE).setValue(classe);
      return;
    }
    // 2️⃣ Automático do e-mail
    if (mapaEmail[key]) {
      const { nota, serie } = mapaEmail[key];
      abaAbas.getRange(lp, C_NOTA).setValue(nota);
      abaAbas.getRange(lp, C_SERIE).setValue(serie);
      return;
    }
    // 3️⃣ Data + Valor
    const rawData = linha[C_DATA - 1];
    const rawValor = linha[C_VAL_BRUTO - 1];

    // 1) Converte a data da TABELA ABASTECIMENTO para Date (se for string)
    let dataAb;
    if (rawData instanceof Date) {
      dataAb = rawData;
    } else if (typeof rawData === 'string' && rawData.includes('/')) {
      const [dia, mes, ano] = rawData.split('/');
      dataAb = new Date(+ano, +mes - 1, +dia);
    }

    // 2) Converte o valor bruto (string “R$233,71” ou número)
    let valorAb = NaN;
    if (typeof rawValor === 'number') {
      valorAb = rawValor;  // já é número
    } else if (typeof rawValor === 'string') {
      // remove tudo que não seja dígito, vírgula ou ponto
      const s = rawValor
        .trim()
        .replace(/[^0-9,.\-]+/g, '')
        // remove separador de milhar (pontos) e troca vírgula por ponto
        .replace(/\./g, '')
        .replace(',', '.');
      valorAb = parseFloat(s);
    }

    if (dataAb instanceof Date && !isNaN(valorAb)) {
      const dataLimpa = dataSemHora(dataAb);

      // percorre cada e-mail procurando match de data+valor
      for (const [, dataEmailRaw, serie, nota, , , valorEmailRaw] of dadosEmail) {
        // converte data do e-mail
        let dataEmail;
        if (dataEmailRaw instanceof Date) {
          dataEmail = dataEmailRaw;
        } else if (typeof dataEmailRaw === 'string' && dataEmailRaw.includes('/')) {
          const [d, m, a] = dataEmailRaw.split('/');
          dataEmail = new Date(+a, +m - 1, +d);
        } else {
          continue;
        }
        const dataEmailLimpa = dataSemHora(dataEmail);

        // converte valor do e-mail (mesma lógica)
        let valorEmail = NaN;
        if (typeof valorEmailRaw === 'number') {
          valorEmail = valorEmailRaw;
        } else if (typeof valorEmailRaw === 'string') {
          const t = valorEmailRaw
            .trim()
            .replace(/[^0-9,.\-]+/g, '')
            .replace(/\./g, '')
            .replace(',', '.');
          valorEmail = parseFloat(t);
        }

        // compara dia e valor (tolerância 1 centavo)
        if (
          dataEmailLimpa.getTime() === dataLimpa.getTime() &&
          Math.abs(valorEmail - valorAb) < 0.01
        ) {
          // preenche nota + série e destaca
          abaAbas
            .getRange(lp, C_NOTA, 1, 2)
            .setValues([[nota, serie]])
            .setBackground('#cfe2f3');
          break;
        } else {
          abaAbas.getRange(lp, C_NOTA).setValue("Não encontrado");
        }
      }
    }
  });
}

// remove hora de um Date
function dataSemHora(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


// ------------------------------
// 5️⃣ Envia confirmações de nota por e-mail
// ------------------------------
function enviarConfirmacoesDeNota() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAbas = ss.getSheetByName("TABELA ABASTECIMENTO");
  const abaMil = SpreadsheetApp
    .openById("XXXXXXXXXX")
    .getSheetByName("CAD_MILITARES");
  if (!abaAbas || !abaMil) {
    SpreadsheetApp.getUi().alert("Abas necessárias não encontradas.");
    return;
  }

  // índices de coluna
  const COL_ABNUM = 1;
  const COL_NOTA = 2;
  const COL_SERIE = 3;
  const COL_DATA = 4;
  const COL_POSTO = 14;
  const COL_VIATURA = 6;
  const COL_QUANTIDADE = 27;
  const COL_VAL_BRUTO = 30;
  const COL_FUNCIONAL = 13;
  const COL_CONTROLE = 40;

  const linhaIni = 8;
  const dados = abaAbas.getRange(linhaIni, 1, abaAbas.getLastRow() - linhaIni + 1, COL_CONTROLE)
    .getValues();
  const militares = abaMil.getRange(2, 1, abaMil.getLastRow() - 1, 9).getValues();

  // mapa de e-mails
  const mapaEmails = {};
  militares.forEach(([funcional, , , nome, , , , , email]) => {
    if (funcional && email) {
      mapaEmails[funcional.toString().trim()] = { nome, email };
    }
  });

  dados.forEach((linha, i) => {
    const lp = linhaIni + i;
    const ctrl = linha[COL_CONTROLE - 1];
    if (ctrl && ctrl.toString().toLowerCase().includes("e-mail enviado")) return;

    const nota = linha[COL_NOTA - 1];
    const serie = linha[COL_SERIE - 1];


    const funcional = linha[COL_FUNCIONAL - 1];
    if (!funcional) return;

    const mil = mapaEmails[funcional.toString().trim()];
    if (!mil) {
      abaAbas.getRange(lp, COL_CONTROLE).setValue("e-mail não enviado");
      return;
    }

    if (!nota || nota.toString().trim().toLowerCase() === "não encontrado") {
      if ((ctrl || "").toString().toLowerCase() != "e-mail com formulário enviado") {
        enviarEmailFormularioParaLinha(mil, linha, abaAbas, lp);
      }
      return;
    };


    // formata data
    const rawDate = linha[COL_DATA - 1];
    let dt = rawDate instanceof Date
      ? rawDate
      : (() => {
        const p = rawDate.toString().split('/');
        return new Date(+p[2], +p[1] - 1, +p[0]);
      })();
    const dataFmt = Utilities.formatDate(dt, "GMT-3", "dd/MM/yyyy");

    const corpo = `
    Senhor(a) ${mil.nome},

    Recebemos sua nota de abastecimento:

    VIATURA: ${linha[COL_VIATURA - 1]}
    ABNUM: ${linha[COL_ABNUM - 1]}
    Nota Fiscal: ${nota}
    Série: ${serie}
    Data: ${dataFmt}
    Valor Bruto: ${linha[COL_VAL_BRUTO - 1]}
    Quantidade: ${linha[COL_QUANTIDADE - 1]}
    Posto: ${linha[COL_POSTO - 1]}

    Confira os dados com o original e descarte a nota caso tudo esteja nos conformes.
    
    Respeitosamente,

    3º SGT FILIPE DIAS PEREIRA CAMUZZI
    SSLT - 10ª Cia Ind
    `;

    try {
      GmailApp.sendEmail(mil.email, "Confirmação de Recebimento de Nota de Abastecimento", corpo);
      abaAbas.getRange(lp, COL_CONTROLE).setValue("e-mail enviado");
    } catch (e) {
      Logger.log(`Erro ao enviar para ${mil.email}: ${e}`);
      abaAbas.getRange(lp, COL_CONTROLE).setValue("e-mail não enviado");
    }
  });
}


// ------------------------------
// 6️⃣ Menu customizado (remove chamada automática de atualizarNotas)
// ------------------------------
function onOpen() {
  onOpenMenu();
}

function onOpenMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Atualizar Dados')
    .addItem('Executar atualização de notas', 'atualizarNotas')
    .addItem('Enviar confirmações de notas', 'enviarConfirmacoesDeNota')
    .addToUi();
}


/**
 * Envia um e-mail com o link do formulário para uma única linha de abastecimento.
 * @param {object} militar O objeto com nome e e-mail do militar.
 * @param {Array} linha A linha de dados da planilha de abastecimento.
 * @param {Sheet} aba A aba da planilha para atualizar o status.
 * @param {number} linhaPlanilha O número da linha na planilha para a atualização.
 */
function enviarEmailFormularioParaLinha(militar, linha, aba, linhaPlanilha) {
  // Índices das colunas necessárias para o e-mail
  const COL_VIATURA = 6, COL_ABNUM = 1, COL_DATA = 4, COL_VAL_BRUTO = 30, COL_QUANTIDADE = 27, COL_POSTO = 14, COL_CONTROLE = 40;

  // Extrai os dados da linha para montar o corpo do e-mail
  const viatura = linha[COL_VIATURA - 1];
  const abnum = linha[COL_ABNUM - 1];
  const data = linha[COL_DATA - 1];
  const valorBruto = linha[COL_VAL_BRUTO - 1];
  const quantidade = linha[COL_QUANTIDADE - 1];
  const posto = linha[COL_POSTO - 1];

  const assunto = "Abastecimento fora dos postos parceiros - 10ª Cia Ind";
  const corpo = `
  Olá, ${militar.nome}

  O senhor(a) está recebendo este e-mail porque fez o seguinte abastecimento fora dos postos parceiros da 10ª Cia Ind.:

  VIATURA: ${viatura}
  ABNUM: ${abnum}
  Nota Fiscal: NÃO INFORMADO
  Série: NÃO INFORMADO
  Data: ${data instanceof Date ? Utilities.formatDate(data, "GMT-3", "dd/MM/yyyy") : data}
  Valor Bruto: ${valorBruto}
  Quantidade: ${quantidade}
  Posto: ${posto}

  Portanto será necessário preencher o formulário abaixo:

  https://forms.gle/formadress

  A nota original também deverá ser protocolada no livro da sua subunidade.

  Respeitosamente,

  3º Sgt Filipe
  SSLT 10ª Cia Ind PMES
  `;

  try {
    GmailApp.sendEmail(militar.email, assunto, corpo);
    aba.getRange(linhaPlanilha, COL_CONTROLE).setValue("e-mail com formulário enviado");
  } catch (e) {
    aba.getRange(linhaPlanilha, COL_CONTROLE).setValue("falha ao enviar formulário");
    Logger.log(`Erro ao enviar e-mail de formulário para ${militar.email}: ${e}`);
  }
}
