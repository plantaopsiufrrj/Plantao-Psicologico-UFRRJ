const PLANILHA_ID = '1WV6xUp5yE_blAHwZiwbIyfq_J_4h5huJLb3sJfdx3iQ';
const ABA_CONTAS = 'Conta dos Plantonistas';

const PLANILHA_DISTRIBUICAO_ID = '16fEGfkzAXScoQu3WiQU-FFCibewyk493pvfXXNX-M4Q';


// ==========================================
// SERVIR O SITE
// ==========================================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Plantão Psicológico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ==========================================
// LOGIN
// ==========================================

function verificarLogin(email, senha) {

  // Validação básica dos campos recebidos
  if (!email || !senha) {
    return {
      sucesso: false,
      mensagem: 'Preencha e-mail e senha.'
    };
  }

  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = planilha.getSheetByName(ABA_CONTAS);

  if (!aba) {
    return {
      sucesso: false,
      mensagem: 'A aba "Conta dos Plantonistas" não foi encontrada.'
    };
  }

  const dados = aba.getDataRange().getValues();
  const emailDigitado = String(email).trim().toLowerCase();
  const senhaDigitada = String(senha).trim();

  let emailEncontrado = false;

  for (let i = 1; i < dados.length; i++) {

    const emailPlanilha = String(dados[i][0]).trim().toLowerCase();

    if (emailPlanilha === '') continue; // ignora linhas vazias

    if (emailPlanilha === emailDigitado) {

      emailEncontrado = true;

      const nome = String(dados[i][1]).trim();
      const senhaPlanilha = String(dados[i][2]).trim();
      const ativo = String(dados[i][3]).trim().toLowerCase();
      const perfil = String(dados[i][4]).trim();
      const cabecalhos = dados[0].map(v => String(v).trim().toLowerCase());
      const idxHabilitado = cabecalhos.indexOf('habilitado');
      const habilitadoValor = idxHabilitado >= 0 ? String(dados[i][idxHabilitado]).trim().toLowerCase() : 'sim';
      const habilitado = !['não','nao','inabilitado','false','0'].includes(habilitadoValor);

      // Senha errada
      if (senhaPlanilha !== senhaDigitada) {
        return {
          sucesso: false,
          mensagem: 'E-mail ou senha incorretos.'
        };
      }

      // Sistema em manutenção: só o bolsista consegue entrar
      const manutencaoLigada = PropertiesService.getScriptProperties().getProperty('MANUTENCAO') === 'ON';
      if (manutencaoLigada && perfil.toLowerCase() !== 'bolsista') {
        return {
          sucesso: false,
          mensagem: 'O sistema está em manutenção no momento. Por favor, aguarde.'
        };
      }

      // Conta inativa: login permitido, mas com acesso restrito
      // (o front-end usa esse "ativo" para mostrar só o botão Meus Plantões)
      const contaAtiva = (
        ativo === 'sim' ||
        ativo === 'ativo' ||
        ativo === 'true'
      );

      // Login válido
      return {
        sucesso: true,
        nome: nome,
        email: emailPlanilha,
        perfil: perfil,
        habilitado: habilitado,
        ativo: contaAtiva
      };
    }
  }

  // E-mail não encontrado na planilha
  if (!emailEncontrado) {
    return {
      sucesso: false,
      mensagem: 'E-mail ou senha incorretos.'
    };
  }
}


// ==========================================
// SOLICITAR CLIENTE
// ==========================================
// Chamada pelo botão "Solicitar Cliente" no painel.
// Usa os dados do plantonista já logado (nome, e-mail, perfil)
// em vez de ler uma resposta de formulário.

function verificarHabilitacaoPlantonista_(email) {
  const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_CONTAS);
  if (!sh) throw new Error('Aba de contas não encontrada.');
  const dados = sh.getDataRange().getValues();
  if (!dados.length) throw new Error('Conta não encontrada.');
  const cab = dados[0].map(v => String(v).trim().toLowerCase());
  const idxEmail = cab.indexOf('e-mail');
  const idxHabilitado = cab.indexOf('habilitado');
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][idxEmail >= 0 ? idxEmail : 0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      if (idxHabilitado < 0) return true;
      const v = String(dados[i][idxHabilitado]).trim().toLowerCase();
      return !['não','nao','inabilitado','false','0'].includes(v);
    }
  }
  throw new Error('Conta não encontrada.');
}

// Checa a coluna "Ativo" da planilha de contas. Contas inativas podem
// logar (ver "Meus Plantões"), mas não podem usar as demais funções.
function verificarContaAtiva_(email) {
  const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_CONTAS);
  if (!sh) throw new Error('Aba de contas não encontrada.');
  const dados = sh.getDataRange().getValues();
  const cab = dados[0].map(v => String(v).trim().toLowerCase());
  const idxEmail = cab.indexOf('e-mail');
  const idxAtivo = cab.indexOf('ativo');
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][idxEmail >= 0 ? idxEmail : 0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      if (idxAtivo < 0) return true;
      const v = String(dados[i][idxAtivo]).trim().toLowerCase();
      return v === 'sim' || v === 'ativo' || v === 'true';
    }
  }
  return false;
}

function solicitarCliente(email, nome, perfil) {

  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    if (!verificarContaAtiva_(email)) throw new Error('Sua conta está inativa. Disponível apenas: Meus Plantões.');

    if (!plantonistaCumpriuFrequenciaMinima_(nome)) {
      return {
        sucesso: false,
        mensagem: 'Você não atingiu o mínimo de 3 presenças em supervisão/treinamento nos últimos 30 dias. A distribuição de clientes está bloqueada até regularizar sua frequência.'
      };
    }

    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const abaClientes = planilha.getSheetByName('Lista de clientes');
    const abaTreinamento = planilha.getSheetByName('Lista de clientes - treinamento');

    const tipoPlantonista = perfil ? perfil.toString().toLowerCase().trim() : '';
    const sheet = (tipoPlantonista === 'treinando') ? abaTreinamento : abaClientes;

    if (!sheet) {
      registrarSolicitacao(planilha, nome, email, tipoPlantonista, 'Aba de clientes não encontrada');
      return {
        sucesso: false,
        mensagem: 'Não encontrei a aba de clientes correspondente na planilha de distribuição.'
      };
    }

    const dados = sheet.getDataRange().getValues();
    const header = dados[0];
    const idxDist = header.indexOf('Distribuído?');
    const idxPlant = header.indexOf('Plantonista');
    const idxSit = header.indexOf('Situação');

    let cliente, linha;

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][idxDist] !== '✅') {
        cliente = dados[i];
        linha = i + 1;
        break;
      }
    }

    // Nenhum cliente disponível
    if (!cliente) {
      MailApp.sendEmail({
        to: email,
        subject: 'Nenhum cliente disponível no momento',
        body: `Olá ${nome},\n\nTodos os clientes ${
          tipoPlantonista === 'treinando' ? 'de treinamento ' : ''
        }já foram distribuídos. Aguarde antes de solicitar novamente.`
      });

      registrarSolicitacao(planilha, nome, email, tipoPlantonista, 'Nenhum disponível');

      return {
        sucesso: false,
        mensagem: 'Nenhum cliente disponível no momento. Você receberá um aviso por e-mail.'
      };
    }

    // Marca o cliente como distribuído
    sheet.getRange(linha, idxDist + 1).setValue('✅');
    sheet.getRange(linha, idxPlant + 1).setValue(nome);
    sheet.getRange(linha, idxSit + 1).setValue('Pendente');
    const idxSolicitacao = garantirColuna(sheet, 'Data da solicitação');
    sheet.getRange(linha, idxSolicitacao + 1).setValue(new Date());

    // Colunas que podem ir no corpo do e-mail
    const permitidas = [
      'Código',
      'Nome Completo',
      'Número de celular',
      'Idade',
      'Gênero',
      'Cidade',
      'Estado',
      'País',
      'Identificação Étnico-Racial',
      'Orientação afetiva/sexual',
      'Relação com a UFRRJ?',
      'Como soube...',
      'Motivo da procura...',
      'Curso/Função',
      'Melhor turno para atendimento:'
    ];

    let msg = `Olá ${nome},\n\nVocê recebeu os dados de um novo cliente:\n\n`;

    header.forEach((h, i) => {
      if (permitidas.includes(h)) msg += `${h}: ${cliente[i]}\n`;
    });

    msg += `\n⚠️ Lembre-se que o cliente possivelmente está em sofrimento psíquico e aguardando por este atendimento.\n` +
           `❗ Se não for possível rapidamente entrar em contato;\n` +
           `❗ Caso já conheça esse cliente;\n` +
           `❗ Se sentir desconforto em atendê-lo.\n\n` +
           `Não esqueça de narrar em supervisão e voltar o relatório via formulário: https://forms.gle/6sLueQMWuTTpqxjs6 😉\n\n` +
           `Abraços,\nEquipe do Plantão Psicológico da UFRRJ 🌻`;

    MailApp.sendEmail({
      to: email,
      subject: 'Novo cliente designado para você',
      body: msg
    });

    const idxCodigo = header.indexOf('Código');
    const idxNomeCliente = header.indexOf('Nome Completo');
    const codigoCliente = (idxCodigo !== -1) ? cliente[idxCodigo] : cliente[0];
    const nomeCliente = (idxNomeCliente !== -1) ? cliente[idxNomeCliente] : cliente[1];
    const resumoCliente = `(${codigoCliente}) ${nomeCliente}`;

    registrarSolicitacao(planilha, nome, email, tipoPlantonista, resumoCliente);

    return {
      sucesso: true,
      mensagem: 'Cliente designado! Confira seu e-mail para ver os dados.'
    };

  } catch (err) {
    Logger.log('Erro em solicitarCliente: ' + err);
    return {
      sucesso: false,
      mensagem: 'Ocorreu um erro ao solicitar o cliente. Tente novamente.'
    };
  }
}


// ==========================================
// REGISTRO DE SOLICITAÇÕES
// ==========================================
// Grava uma linha na aba "Registro de Solicitações" (criando a aba
// com cabeçalho se ainda não existir) toda vez que solicitarCliente roda.

function registrarSolicitacao(planilha, nome, email, tipoPlantonista, clienteInfo) {

  const NOME_ABA_REGISTRO = 'Registro de Solicitações';
  let aba = planilha.getSheetByName(NOME_ABA_REGISTRO);

  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA_REGISTRO);
    aba.appendRow(['Data/Hora', 'Plantonista', 'E-mail', 'Tipo', 'Cliente']);
  }

  const tipoTexto = (tipoPlantonista === 'treinando') ? 'Treinando' : 'Plantonista';

  aba.appendRow([new Date(), nome, email, tipoTexto, clienteInfo || '']);
}


// ==========================================
// MEUS PLANTÕES
// ==========================================
// Usa as colunas "Plantonista" e "Situação" da planilha de distribuição
// (Q e R na planilha, localizadas dinamicamente pelo nome do cabeçalho)
// para contar quantos clientes do plantonista logado estão em cada situação.

function obterResumoPlantao(email, nome, perfil) {

  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const abaClientes = planilha.getSheetByName('Lista de clientes');
    const abaTreinamento = planilha.getSheetByName('Lista de clientes - treinamento');

    const tipoPlantonista = perfil ? perfil.toString().toLowerCase().trim() : '';
    const sheet = (tipoPlantonista === 'treinando') ? abaTreinamento : abaClientes;

    if (!sheet) {
      return {
        sucesso: false,
        mensagem: 'Não encontrei a aba de clientes correspondente na planilha de distribuição.'
      };
    }

    const contagem = contarClientesPorSituacao(sheet, nome);

    return {
      sucesso: true,
      pendentes: contagem.pendentes,
      atendidos: contagem.atendidos,
      desistentes: contagem.desistentes,
      transferidos: contagem.transferidos
    };

  } catch (err) {
    Logger.log('Erro em obterResumoPlantao: ' + err);
    return {
      sucesso: false,
      mensagem: 'Ocorreu um erro ao buscar o resumo do plantão.'
    };
  }
}


// Função auxiliar: percorre a aba e conta, por situação, os clientes
// cuja coluna "Plantonista" bate com o nome informado.

function contarClientesPorSituacao(sheet, nomePlantonista) {

  const dados = sheet.getDataRange().getValues();
  const header = dados[0];
  const idxPlant = header.indexOf('Plantonista');
  const idxSit = header.indexOf('Situação');

  const nomeComparar = String(nomePlantonista).trim().toLowerCase();

  let pendentes = 0, atendidos = 0, desistentes = 0, transferidos = 0;

  for (let i = 1; i < dados.length; i++) {

    const plantonistaLinha = String(dados[i][idxPlant]).trim().toLowerCase();
    if (plantonistaLinha !== nomeComparar) continue;

    const situacao = String(dados[i][idxSit]).trim().toLowerCase();

    if (situacao === 'pendente') pendentes++;
    else if (situacao === 'atendido') atendidos++;
    else if (situacao === 'desistente') desistentes++;
    else if (situacao === 'transferido') transferidos++;
  }

  return { pendentes, atendidos, desistentes, transferidos };
}


// ==========================================
// SOLICITAR DADOS DOS MEUS CLIENTES
// ==========================================
// Reúne todas as linhas em que o plantonista logado aparece na coluna
// "Plantonista" e manda os dados desses clientes por e-mail, junto
// com o resumo do plantão.

function solicitarDadosClientes(email, nome, perfil) {

  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const abaClientes = planilha.getSheetByName('Lista de clientes');
    const abaTreinamento = planilha.getSheetByName('Lista de clientes - treinamento');

    const tipoPlantonista = perfil ? perfil.toString().toLowerCase().trim() : '';
    const sheet = (tipoPlantonista === 'treinando') ? abaTreinamento : abaClientes;

    if (!sheet) {
      return {
        sucesso: false,
        mensagem: 'Não encontrei a aba de clientes correspondente na planilha de distribuição.'
      };
    }

    const dados = sheet.getDataRange().getValues();
    const header = dados[0];
    const idxPlant = header.indexOf('Plantonista');

    const nomeComparar = String(nome).trim().toLowerCase();
    const linhasCliente = [];

    for (let i = 1; i < dados.length; i++) {
      const plantonistaLinha = String(dados[i][idxPlant]).trim().toLowerCase();
      if (plantonistaLinha === nomeComparar) {
        linhasCliente.push(dados[i]);
      }
    }

    if (linhasCliente.length === 0) {
      return {
        sucesso: false,
        mensagem: 'Você ainda não tem clientes registrados.'
      };
    }

    // Mesmas colunas permitidas do e-mail de "Solicitar Cliente", com a Situação a mais
    const permitidas = [
      'Código',
      'Nome Completo',
      'Número de celular',
      'Idade',
      'Gênero',
      'Cidade',
      'Estado',
      'País',
      'Identificação Étnico-Racial',
      'Orientação afetiva/sexual',
      'Relação com a UFRRJ?',
      'Como soube...',
      'Motivo da procura...',
      'Curso/Função',
      'Melhor turno para atendimento:',
      'Situação'
    ];

    let msg = `Olá ${nome},\n\nSegue os dados dos seus clientes:\n\n`;

    linhasCliente.forEach((cliente, indice) => {
      msg += `--- Cliente ${indice + 1} ---\n`;
      header.forEach((h, i) => {
        if (permitidas.includes(h)) msg += `${h}: ${cliente[i]}\n`;
      });
      msg += `\n`;
    });

    const contagem = contarClientesPorSituacao(sheet, nome);

    msg += `Meu resumo do plantão:\n` +
           `${contagem.pendentes} pendente(s)\n` +
           `${contagem.atendidos} atendido(s)\n` +
           `${contagem.desistentes} desistente(s)\n` +
           `${contagem.transferidos} transferido(s)\n\n` +
           `Abraços,\nEquipe do Plantão Psicológico da UFRRJ 🌻`;

    MailApp.sendEmail({
      to: email,
      subject: 'Dados dos seus clientes - Plantão Psicológico',
      body: msg
    });

    return {
      sucesso: true,
      mensagem: 'Dados enviados! Confira seu e-mail.'
    };

  } catch (err) {
    Logger.log('Erro em solicitarDadosClientes: ' + err);
    return {
      sucesso: false,
      mensagem: 'Ocorreu um erro ao solicitar os dados. Tente novamente.'
    };
  }
}


// ==========================================
// JARDIM DO ADMINISTRADOR — JANELAS DE FREQUÊNCIA
// ==========================================
// Área restrita ao bolsista (perfil === 'bolsista' na planilha de contas).
// Cada janela representa um período em que o registro de frequência
// fica liberado para os plantonistas.

const NOME_ABA_JANELAS = 'Janelas de Frequência';

// Confere no backend se quem está chamando é realmente o bolsista.
// Importante: o botão escondido no front-end não é suficiente sozinho,
// porque alguém poderia chamar a função do Apps Script diretamente.
function ehBolsista(perfil) {
  return perfil && perfil.toString().trim().toLowerCase() === 'bolsista';
}

function abrirJanelaFrequencia(perfil, data, horaInicio, horaFim, tipo) {

  if (!ehBolsista(perfil)) {
    return { sucesso: false, mensagem: 'Acesso restrito ao bolsista.' };
  }

  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    let aba = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!aba) {
      aba = planilha.insertSheet(NOME_ABA_JANELAS);
      aba.appendRow(['ID', 'Data', 'Hora Início', 'Hora Fim', 'Tipo', 'Status']);
    }

    const tiposValidos = ['supervisao_plantonistas', 'supervisao_treinandos', 'treinamento_teorico'];
    const tipoFinal = tiposValidos.includes(tipo) ? tipo : 'supervisao_plantonistas';

    const id = Utilities.getUuid();

    aba.appendRow([id, data, horaInicio, horaFim, tipoFinal, 'Aberta']);

    return { sucesso: true, mensagem: 'Janela de frequência aberta.' };

  } catch (err) {
    Logger.log('Erro em abrirJanelaFrequencia: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao abrir a janela.' };
  }
}

function fecharJanelaFrequencia(perfil, id) {

  if (!ehBolsista(perfil)) {
    return { sucesso: false, mensagem: 'Acesso restrito ao bolsista.' };
  }

  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const aba = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!aba) {
      return { sucesso: false, mensagem: 'Nenhuma janela cadastrada ainda.' };
    }

    const dados = aba.getDataRange().getValues();
    const header = dados[0];
    const idxId = header.indexOf('ID');
    const idxStatus = header.indexOf('Status');

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][idxId] === id) {
        aba.getRange(i + 1, idxStatus + 1).setValue('Fechada');
        return { sucesso: true, mensagem: 'Janela fechada.' };
      }
    }

    return { sucesso: false, mensagem: 'Janela não encontrada.' };

  } catch (err) {
    Logger.log('Erro em fecharJanelaFrequencia: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao fechar a janela.' };
  }
}

function listarJanelasFrequencia(perfil) {

  if (!ehBolsista(perfil)) {
    return { sucesso: false, mensagem: 'Acesso restrito ao bolsista.' };
  }

  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const aba = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!aba) {
      return { sucesso: true, janelas: [] };
    }

    const dados = aba.getDataRange().getValues();
    const header = dados[0];
    const idxId = header.indexOf('ID');
    const idxData = header.indexOf('Data');
    const idxInicio = header.indexOf('Hora Início');
    const idxFim = header.indexOf('Hora Fim');
    const idxTipo = header.indexOf('Tipo');
    const idxStatus = header.indexOf('Status');

    const janelas = [];

    for (let i = 1; i < dados.length; i++) {
      janelas.push({
        id: dados[i][idxId],
        data: formatarDataExibicao(dados[i][idxData]),
        horaInicio: formatarHoraExibicao(dados[i][idxInicio]),
        horaFim: formatarHoraExibicao(dados[i][idxFim]),
        tipo: dados[i][idxTipo],
        status: dados[i][idxStatus]
      });
    }

    // Mais recentes primeiro
    janelas.reverse();

    return { sucesso: true, janelas: janelas };

  } catch (err) {
    Logger.log('Erro em listarJanelasFrequencia: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao buscar as janelas.' };
  }
}

// O Apps Script guarda datas/horas digitadas em <input type="date"/"time">
// como objetos Date. Essas funções só convertem de volta para texto legível.
function formatarDataExibicao(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return valor;
}

function formatarHoraExibicao(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'HH:mm');
  }
  return valor;
}

// Junta a data (da coluna "Data") com a hora (de "Hora Início" ou "Hora Fim")
// num único objeto Date real, pra dar pra comparar com o momento atual.
function construirDataHora(dataValor, horaValor) {

  let data;
  if (Object.prototype.toString.call(dataValor) === '[object Date]') {
    data = new Date(dataValor.getFullYear(), dataValor.getMonth(), dataValor.getDate());
  } else {
    const partesData = String(dataValor).split('-'); // formato vindo do <input type="date">: AAAA-MM-DD
    data = new Date(Number(partesData[0]), Number(partesData[1]) - 1, Number(partesData[2]));
  }

  let horas = 0, minutos = 0;
  if (Object.prototype.toString.call(horaValor) === '[object Date]') {
    horas = horaValor.getHours();
    minutos = horaValor.getMinutes();
  } else {
    const partesHora = String(horaValor).split(':'); // formato vindo do <input type="time">: HH:MM
    horas = Number(partesHora[0]);
    minutos = Number(partesHora[1]);
  }

  data.setHours(horas, minutos, 0, 0);
  return data;
}

// Verifica se o momento atual está dentro do intervalo [data+horaInicio, data+horaFim].
function janelaDentroDoHorario(dataValor, horaInicioValor, horaFimValor) {
  const agora = new Date();
  const inicio = construirDataHora(dataValor, horaInicioValor);
  const fim = construirDataHora(dataValor, horaFimValor);
  return agora >= inicio && agora <= fim;
}


// ==========================================
// BLOQUEIO DE SOLICITAÇÃO POR FREQUÊNCIA MÍNIMA
// ==========================================
// Regra: em 30 dias, o plantonista precisa ter comparecido a pelo menos
// 3 janelas de frequência já encerradas. Registros de frequência
// realizados pelo plantonista contam como comparecimento.
// Se ainda não houve pelo menos 3 janelas encerradas no período, ninguém
// é bloqueado (não é possível exigir 3 presenças se só houve 1 ou 2 chances).

// Monta um mapa { janelaId: {data, fimCompleto} } com todas as janelas cadastradas.
function obterMapaJanelas_() {
  const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
  const aba = planilha.getSheetByName(NOME_ABA_JANELAS);
  const mapa = {};

  if (!aba) return mapa;

  const dados = aba.getDataRange().getValues();
  const header = dados[0];
  const idxId = header.indexOf('ID');
  const idxData = header.indexOf('Data');
  const idxFim = header.indexOf('Hora Fim');

  for (let i = 1; i < dados.length; i++) {
    const id = dados[i][idxId];
    mapa[id] = {
      data: dados[i][idxData],
      fimCompleto: construirDataHora(dados[i][idxData], dados[i][idxFim])
    };
  }

  return mapa;
}

function plantonistaCumpriuFrequenciaMinima_(nome) {

  const MINIMO_PRESENCAS = 3;
  const DIAS_JANELA = 30;

  const agora = new Date();
  const limite = new Date(agora.getTime() - DIAS_JANELA * 24 * 60 * 60 * 1000);

  const mapaJanelas = obterMapaJanelas_();

  // Quantas janelas já encerradas existiram nos últimos 30 dias (oportunidades reais)
  let oportunidades = 0;
  Object.keys(mapaJanelas).forEach(id => {
    const fim = mapaJanelas[id].fimCompleto;
    if (fim <= agora && fim >= limite) oportunidades++;
  });

  if (oportunidades < MINIMO_PRESENCAS) {
    return true; // ainda não houve chances suficientes para exigir o mínimo
  }

  const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
  const abaRegistros = planilha.getSheetByName(NOME_ABA_REGISTROS_FREQ);

  if (!abaRegistros) return false; // havia oportunidades, mas nunca houve nenhum registro

  const dados = abaRegistros.getDataRange().getValues();
  const header = dados[0];
  const idxJanelaId = header.indexOf('Janela ID');
  const idxNome = header.indexOf('Plantonista');
  const nomeComparar = String(nome).trim().toLowerCase();

  const janelasContadas = {};
  let presencas = 0;

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][idxNome]).trim().toLowerCase() !== nomeComparar) continue;

    const janelaId = dados[i][idxJanelaId];
    const janela = mapaJanelas[janelaId];
    if (!janela) continue;

    const dentroDoPeriodo = janela.fimCompleto <= agora && janela.fimCompleto >= limite;

    if (dentroDoPeriodo && !janelasContadas[janelaId]) {
      presencas++;
      janelasContadas[janelaId] = true; // evita contar a mesma janela duas vezes
    }
  }

  return presencas >= MINIMO_PRESENCAS;
}


// ==========================================
// REGISTRAR FREQUÊNCIA — LADO DO PLANTONISTA
// ==========================================

const NOME_ABA_REGISTROS_FREQ = 'Registros de Frequência';

// Lista as janelas com status "Aberta" para qualquer plantonista logado ver.
function listarJanelasAbertas() {

  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const aba = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!aba) {
      return { sucesso: true, janelas: [] };
    }

    const dados = aba.getDataRange().getValues();
    const header = dados[0];
    const idxId = header.indexOf('ID');
    const idxData = header.indexOf('Data');
    const idxInicio = header.indexOf('Hora Início');
    const idxFim = header.indexOf('Hora Fim');
    const idxTipo = header.indexOf('Tipo');
    const idxStatus = header.indexOf('Status');

    const janelas = [];

    for (let i = 1; i < dados.length; i++) {
      const statusAberta = String(dados[i][idxStatus]).trim() === 'Aberta';
      const dentroDoHorario = janelaDentroDoHorario(dados[i][idxData], dados[i][idxInicio], dados[i][idxFim]);

      if (statusAberta && dentroDoHorario) {
        janelas.push({
          id: dados[i][idxId],
          data: formatarDataExibicao(dados[i][idxData]),
          horaInicio: formatarHoraExibicao(dados[i][idxInicio]),
          horaFim: formatarHoraExibicao(dados[i][idxFim]),
          tipo: dados[i][idxTipo]
        });
      }
    }

    return { sucesso: true, janelas: janelas };

  } catch (err) {
    Logger.log('Erro em listarJanelasAbertas: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao buscar as janelas abertas.' };
  }
}

// Registra (ou sobrescreve) a presença do plantonista numa janela.
function registrarFrequencia(email, nome, perfil, idJanela, narrou, palavraDia) {

  try {
    if (!verificarContaAtiva_(email)) {
      return { sucesso: false, mensagem: 'Sua conta está inativa. Disponível apenas: Meus Plantões.' };
    }

    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const abaJanelas = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!abaJanelas) {
      return { sucesso: false, mensagem: 'Nenhuma janela de frequência cadastrada.' };
    }

    // Confirma que a janela existe e ainda está aberta
    // (checagem de segurança: o front-end já filtra isso, mas o estado
    // pode ter mudado entre o plantonista abrir a tela e confirmar o envio)
    const dadosJanelas = abaJanelas.getDataRange().getValues();
    const headerJ = dadosJanelas[0];
    const idxId = headerJ.indexOf('ID');
    const idxData = headerJ.indexOf('Data');
    const idxStatus = headerJ.indexOf('Status');

    let janela = null;

    for (let i = 1; i < dadosJanelas.length; i++) {
      if (dadosJanelas[i][idxId] === idJanela) {
        janela = dadosJanelas[i];
        break;
      }
    }

    if (!janela) {
      return { sucesso: false, mensagem: 'Essa janela não foi encontrada.' };
    }

    if (String(janela[idxStatus]).trim() !== 'Aberta') {
      return { sucesso: false, mensagem: 'Essa janela não está mais aberta para registro.' };
    }

    const idxInicio = headerJ.indexOf('Hora Início');
    const idxFim = headerJ.indexOf('Hora Fim');

    if (!janelaDentroDoHorario(janela[idxData], janela[idxInicio], janela[idxFim])) {
      return {
        sucesso: false,
        mensagem: 'Essa janela não está mais dentro do horário permitido para registro (' +
          formatarHoraExibicao(janela[idxInicio]) + '–' + formatarHoraExibicao(janela[idxFim]) + ').'
      };
    }

    // Supervisão Plantonistas e Supervisão Treinandos exigem narrou + palavra
    // do dia. Treinamento teórico não exige nada além da confirmação.
    const idxTipo = headerJ.indexOf('Tipo');
    const tipoJanela = janela[idxTipo];
    const exigeCampos = (tipoJanela === 'supervisao_plantonistas' || tipoJanela === 'supervisao_treinandos');

    if (exigeCampos) {
      if (narrou !== true && narrou !== false) {
        return { sucesso: false, mensagem: 'Informe se narrou um caso ou não.' };
      }
      if (!palavraDia || !palavraDia.toString().trim()) {
        return { sucesso: false, mensagem: 'A palavra do dia é obrigatória para esse tipo de supervisão.' };
      }
    }

    // Prepara (ou cria) a aba de registros
    let abaRegistros = planilha.getSheetByName(NOME_ABA_REGISTROS_FREQ);

    if (!abaRegistros) {
      abaRegistros = planilha.insertSheet(NOME_ABA_REGISTROS_FREQ);
      abaRegistros.appendRow([
        'Janela ID', 'Data da Janela', 'Plantonista', 'E-mail',
        'Narrou', 'Palavra do Dia', 'Registrado em'
      ]);
    }

    const dadosRegistros = abaRegistros.getDataRange().getValues();
    const headerR = dadosRegistros[0];
    const idxJanelaIdR = headerR.indexOf('Janela ID');
    const idxEmailR = headerR.indexOf('E-mail');

    const emailComparar = String(email).trim().toLowerCase();
    let linhaExistente = -1;

    for (let i = 1; i < dadosRegistros.length; i++) {
      const mesmaJanela = dadosRegistros[i][idxJanelaIdR] === idJanela;
      const mesmoEmail = String(dadosRegistros[i][idxEmailR]).trim().toLowerCase() === emailComparar;
      if (mesmaJanela && mesmoEmail) {
        linhaExistente = i + 1; // +1 porque getRange é 1-indexado
        break;
      }
    }

    const narrouTexto = exigeCampos ? (narrou ? 'Sim' : 'Não') : 'Presença confirmada';
    const linha = [
      idJanela,
      formatarDataExibicao(janela[idxData]),
      nome,
      email,
      narrouTexto,
      exigeCampos ? (palavraDia || '') : '',
      new Date()
    ];

    if (linhaExistente !== -1) {
      // Sobrescreve o registro já existente (permitido pela regra do projeto)
      abaRegistros.getRange(linhaExistente, 1, 1, linha.length).setValues([linha]);
    } else {
      abaRegistros.appendRow(linha);
    }

    return { sucesso: true, mensagem: 'Presença registrada com sucesso!' };

  } catch (err) {
    Logger.log('Erro em registrarFrequencia: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao registrar a frequência. Tente novamente.' };
  }
}


// ==========================================
// REGISTRAR MOTIVO DE FALTA
// ==========================================

const NOME_ABA_JUSTIFICATIVAS = 'Justificativas de Falta';

// Tolera a aba ter sido criada como "Justificativa de Falta" (singular)
// em vez do nome padrão, pra nunca ficar sem achar os dados por causa
// de uma diferença de plural.
function obterAbaJustificativas_(planilha) {
  return planilha.getSheetByName('Justificativas de Falta') ||
         planilha.getSheetByName('Justificativa de Falta');
}

// Lista TODAS as janelas (abertas e fechadas), para o plantonista
// poder justificar a ausência em qualquer dia já cadastrado.
function listarTodasJanelas() {

  try {
    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const aba = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!aba) {
      return { sucesso: true, janelas: [] };
    }

    const dados = aba.getDataRange().getValues();
    const header = dados[0];
    const idxId = header.indexOf('ID');
    const idxData = header.indexOf('Data');
    const idxInicio = header.indexOf('Hora Início');
    const idxFim = header.indexOf('Hora Fim');
    const idxTipo = header.indexOf('Tipo');
    const idxStatus = header.indexOf('Status');

    const janelas = [];

    for (let i = 1; i < dados.length; i++) {
      janelas.push({
        id: dados[i][idxId],
        data: formatarDataExibicao(dados[i][idxData]),
        horaInicio: formatarHoraExibicao(dados[i][idxInicio]),
        horaFim: formatarHoraExibicao(dados[i][idxFim]),
        tipo: dados[i][idxTipo],
        status: dados[i][idxStatus]
      });
    }

    janelas.reverse(); // mais recentes primeiro

    return { sucesso: true, janelas: janelas };

  } catch (err) {
    Logger.log('Erro em listarTodasJanelas: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao buscar os dias cadastrados.' };
  }
}

// O plantonista envia o motivo da ausência para um dia específico.
// O motivo é registrado para consulta, sem fluxo de aprovação pelo site.
function enviarJustificativaFalta(email, nome, perfil, idJanela, motivo) {

  try {
    if (!verificarContaAtiva_(email)) {
      return { sucesso: false, mensagem: 'Sua conta está inativa. Disponível apenas: Meus Plantões.' };
    }

    if (!motivo || !motivo.toString().trim()) {
      return { sucesso: false, mensagem: 'Escreva o motivo da ausência.' };
    }

    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const abaJanelas = planilha.getSheetByName(NOME_ABA_JANELAS);

    if (!abaJanelas) {
      return { sucesso: false, mensagem: 'Nenhum dia cadastrado ainda.' };
    }

    const dadosJanelas = abaJanelas.getDataRange().getValues();
    const headerJ = dadosJanelas[0];
    const idxId = headerJ.indexOf('ID');
    const idxData = headerJ.indexOf('Data');

    let janela = null;
    for (let i = 1; i < dadosJanelas.length; i++) {
      if (dadosJanelas[i][idxId] === idJanela) {
        janela = dadosJanelas[i];
        break;
      }
    }

    if (!janela) {
      return { sucesso: false, mensagem: 'Esse dia não foi encontrado.' };
    }

    let abaJust = obterAbaJustificativas_(planilha);
    if (!abaJust) {
      abaJust = planilha.insertSheet(NOME_ABA_JUSTIFICATIVAS);
      abaJust.appendRow([
        'ID', 'Janela ID', 'Data da Janela', 'Plantonista', 'E-mail',
        'Motivo', 'Status', 'Enviado em', 'Decidido em'
      ]);
    }

    const id = Utilities.getUuid();

    abaJust.appendRow([
      id,
      idJanela,
      formatarDataExibicao(janela[idxData]),
      nome,
      email,
      motivo.toString().trim(),
      'Pendente',
      new Date(),
      ''
    ]);

    return { sucesso: true, mensagem: 'Motivo da ausência enviado com sucesso.' };

  } catch (err) {
    Logger.log('Erro em enviarJustificativaFalta: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao enviar a justificativa. Tente novamente.' };
  }
}


// ==========================================
// ANÁLISE DE JUSTIFICATIVAS (BOLSISTA/ADMIN)
// ==========================================

function verificarBolsista_(email) {
  const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_CONTAS);
  if (!sh) throw new Error('Aba de contas não encontrada.');
  const dados = sh.getDataRange().getValues();
  if (!dados.length) throw new Error('Aba de contas vazia.');
  const cab = dados[0].map(v => String(v).trim().toLowerCase());
  const idxEmail = cab.indexOf('e-mail');
  const idxPerfil = cab.indexOf('perfil');
  if (idxEmail < 0 || idxPerfil < 0) throw new Error('Colunas E-mail e Perfil não encontradas.');
  const emailBusca = String(email || '').trim().toLowerCase();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][idxEmail]).trim().toLowerCase() === emailBusca) {
      return String(dados[i][idxPerfil]).trim().toLowerCase() === 'bolsista';
    }
  }
  return false;
}

function listarJustificativasPendentes(email) {
  try {
    if (!verificarBolsista_(email)) {
      return { sucesso: false, mensagem: 'Acesso restrito à equipe administrativa.' };
    }

    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const aba = obterAbaJustificativas_(planilha);
    if (!aba || aba.getLastRow() < 2) return { sucesso: true, justificativas: [] };

    const dados = aba.getDataRange().getValues();
    const cab = dados[0].map(v => String(v).trim());
    const idx = {
      id: cab.indexOf('ID'),
      janela: cab.indexOf('Janela ID'),
      data: cab.indexOf('Data da Janela'),
      nome: cab.indexOf('Plantonista'),
      email: cab.indexOf('E-mail'),
      motivo: cab.indexOf('Motivo'),
      status: cab.indexOf('Status'),
      enviado: cab.indexOf('Enviado em')
    };
    if (Object.values(idx).some(i => i < 0)) throw new Error('Cabeçalhos da aba de justificativas estão incompletos.');

    const justificativas = [];
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][idx.status]).trim().toLowerCase() === 'pendente') {
        justificativas.push({
          id: String(dados[i][idx.id]),
          janelaId: String(dados[i][idx.janela]),
          data: dados[i][idx.data] instanceof Date
            ? formatarDataExibicao(dados[i][idx.data])
            : String(dados[i][idx.data] || ''),
          nome: String(dados[i][idx.nome] || ''),
          email: String(dados[i][idx.email] || ''),
          motivo: String(dados[i][idx.motivo] || ''),
          enviadoEm: dados[i][idx.enviado] instanceof Date
            ? Utilities.formatDate(dados[i][idx.enviado], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
            : String(dados[i][idx.enviado] || '')
        });
      }
    }
    justificativas.reverse();
    return { sucesso: true, justificativas: justificativas };
  } catch (err) {
    Logger.log('Erro em listarJustificativasPendentes: ' + err);
    return { sucesso: false, mensagem: 'Não foi possível carregar as justificativas.' };
  }
}

function decidirJustificativaFalta(email, id, decisao) {
  try {
    if (!verificarBolsista_(email)) {
      return { sucesso: false, mensagem: 'Acesso restrito à equipe administrativa.' };
    }
    const escolha = String(decisao || '').trim().toLowerCase();
    if (!['aprovada', 'recusada'].includes(escolha)) {
      return { sucesso: false, mensagem: 'Decisão inválida.' };
    }

    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const aba = obterAbaJustificativas_(planilha);
    if (!aba || aba.getLastRow() < 2) {
      return { sucesso: false, mensagem: 'Justificativa não encontrada.' };
    }
    const dados = aba.getDataRange().getValues();
    const cab = dados[0].map(v => String(v).trim());
    const idxId = cab.indexOf('ID');
    const idxStatus = cab.indexOf('Status');
    const idxDecidido = cab.indexOf('Decidido em');
    if (idxId < 0 || idxStatus < 0 || idxDecidido < 0) {
      return { sucesso: false, mensagem: 'Cabeçalhos da aba de justificativas estão incompletos.' };
    }

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][idxId]) === String(id)) {
        if (String(dados[i][idxStatus]).trim().toLowerCase() !== 'pendente') {
          return { sucesso: false, mensagem: 'Esta justificativa já foi analisada.' };
        }
        aba.getRange(i + 1, idxStatus + 1).setValue(escolha === 'aprovada' ? 'Aprovada' : 'Recusada');
        aba.getRange(i + 1, idxDecidido + 1).setValue(new Date());
        return { sucesso: true, mensagem: escolha === 'aprovada' ? 'Justificativa aprovada.' : 'Justificativa recusada.' };
      }
    }
    return { sucesso: false, mensagem: 'Justificativa não encontrada.' };
  } catch (err) {
    Logger.log('Erro em decidirJustificativaFalta: ' + err);
    return { sucesso: false, mensagem: 'Não foi possível registrar a decisão.' };
  }
}

const PASTA_RELATORIOS_NOME = 'Relatórios de Atendimento Recebidos';

function obterAbaDoPerfil(perfil) {
  const ss = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
  const treinamento = String(perfil || '').trim().toLowerCase() === 'treinando';
  const nomeAba = treinamento ? 'Lista de clientes - treinamento' : 'Lista de clientes';
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) throw new Error('A aba ' + nomeAba + ' não foi encontrada.');
  return { ss: ss, aba: aba };
}

function indicesCabecalho(aba) {
  const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(String);
  const idx = {};
  cab.forEach((v, i) => idx[v.trim()] = i);
  return { cabecalho: cab, indices: idx };
}

function garantirColuna(aba, nome) {
  let info = indicesCabecalho(aba);
  if (info.indices[nome] === undefined) {
    const col = aba.getLastColumn() + 1;
    aba.getRange(1, col).setValue(nome);
    info = indicesCabecalho(aba);
  }
  return info.indices[nome];
}

function listarClientesPendentes(email, nome, perfil) {
  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    if (!verificarContaAtiva_(email)) throw new Error('Sua conta está inativa. Disponível apenas: Meus Plantões.');
    const {aba} = obterAbaDoPerfil(perfil);
    const dados = aba.getDataRange().getValues();
    if (!dados.length) return {sucesso:true, clientes:[]};
    const idx = {};
    dados[0].forEach((h,i)=>idx[String(h).trim()]=i);
    if (idx['Plantonista'] === undefined || idx['Situação'] === undefined || idx['Código'] === undefined || idx['Nome Completo'] === undefined)
      throw new Error('Confira se a planilha possui os cabeçalhos Código, Nome Completo, Plantonista e Situação.');
    const clientes=[];
    for(let i=1;i<dados.length;i++) {
      if(String(dados[i][idx['Plantonista']]).trim().toLowerCase()===String(nome).trim().toLowerCase() && String(dados[i][idx['Situação']]).trim().toLowerCase()==='pendente') {
        clientes.push({codigo:String(dados[i][idx['Código']]), nome:String(dados[i][idx['Nome Completo']])});
      }
    }
    return {sucesso:true,clientes:clientes};
  } catch(e) { Logger.log(e); return {sucesso:false,mensagem:e.message}; }
}

function atualizarSituacaoCliente(email,nome,perfil,codigo,situacao,detalhe,arquivo) {
  const lock=LockService.getScriptLock();
  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    if (!verificarContaAtiva_(email)) throw new Error('Sua conta está inativa. Disponível apenas: Meus Plantões.');
    if (!['Atendido','Desistente','Transferido'].includes(situacao)) throw new Error('Situação inválida.');
    if (situacao==='Desistente' && !String(detalhe||'').trim()) throw new Error('Informe o motivo da desistência.');
    if (situacao==='Transferido' && !String(detalhe||'').trim()) throw new Error('Informe o nome do novo plantonista.');
    if (situacao==='Atendido' && !arquivo) throw new Error('Anexe o relatório de atendimento.');
    lock.waitLock(30000);
    const {ss,aba}=obterAbaDoPerfil(perfil); const dados=aba.getDataRange().getValues();
    const idx={}; dados[0].forEach((h,i)=>idx[String(h).trim()]=i);
    ['Código','Plantonista','Situação'].forEach(h=>{if(idx[h]===undefined)throw new Error('Cabeçalho ausente: '+h)});
    let row=-1;
    for(let i=1;i<dados.length;i++) if(String(dados[i][idx['Código']]).trim()===String(codigo).trim()){row=i+1;break;}
    if(row<0) throw new Error('Cliente não encontrado.');
    const atual=aba.getRange(row,idx['Plantonista']+1).getValue();
    const status=String(aba.getRange(row,idx['Situação']+1).getValue()).trim();
    if(String(atual).trim().toLowerCase()!==String(nome).trim().toLowerCase() || status.toLowerCase()!=='pendente') throw new Error('Este cliente não está mais pendente sob sua responsabilidade. Atualize a lista.');
    if(situacao==='Atendido') {
      const mime=String(arquivo.mimeType||'').toLowerCase(); const fname=String(arquivo.nome||'');
      const ext=(fname.match(/\.([^.]+)$/)||[])[1];
      if(!['pdf','doc','docx'].includes(String(ext||'').toLowerCase()) || !['application/pdf','application/msword','application/vnd.openxmlformats-officedocument.wordprocessingml.document'].includes(mime)) throw new Error('Envie um arquivo PDF, DOC ou DOCX.');
      const pasta=obterOuCriarPastaRelatorios();
      const nomeArquivo=String(nome)+' - '+String(codigo)+' - Relatório Plantão Psicológico da UFRRJ.'+String(ext).toLowerCase();
      const bytes=Utilities.base64Decode(String(arquivo.base64).split(',').pop());
      const blob=Utilities.newBlob(bytes,mime,nomeArquivo); const file=pasta.createFile(blob);
      const idxRel=garantirColuna(aba,'Relatório de atendimento'); aba.getRange(row,idxRel+1).setValue(file.getUrl());
    } else if(situacao==='Desistente') {
      const c=garantirColuna(aba,'Motivo da desistência'); aba.getRange(row,c+1).setValue(String(detalhe).trim());
    } else {
      const c=garantirColuna(aba,'Plantonista anterior'); aba.getRange(row,c+1).setValue(nome);
      aba.getRange(row,idx['Plantonista']+1).setValue(String(detalhe).trim());
      const c2=garantirColuna(aba,'Data da transferência'); aba.getRange(row,c2+1).setValue(new Date());
    }
    aba.getRange(row,idx['Situação']+1).setValue(situacao);
    const idxDevolutiva = garantirColuna(aba, 'Data e hora da devolutiva');
    aba.getRange(row, idxDevolutiva + 1).setValue(new Date());
    const cEmail=garantirColuna(aba,'E-mail do plantonista responsável'); aba.getRange(row,cEmail+1).setValue(email||'');
    return {sucesso:true,mensagem:'Cliente atualizado para '+situacao+'.'};
  } catch(e) { Logger.log(e); return {sucesso:false,mensagem:e.message||'Erro ao atualizar cliente.'}; }
  finally { try{lock.releaseLock();}catch(ignore){} }
}

function obterOuCriarPastaRelatorios() {
  const pastas=DriveApp.getFoldersByName(PASTA_RELATORIOS_NOME);
  return pastas.hasNext()?pastas.next():DriveApp.createFolder(PASTA_RELATORIOS_NOME);
}

function cadastrarPlantaoPresencial(email,nome,perfil,informacoes,local,arquivo) {
  const lock = LockService.getScriptLock();
  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    if (!verificarContaAtiva_(email)) throw new Error('Sua conta está inativa. Disponível apenas: Meus Plantões.');
    if(!String(local||'').trim()) return {sucesso:false,mensagem:'Informe o local de atendimento.'};
    if(!arquivo) return {sucesso:false,mensagem:'Anexe o relatório do plantão presencial.'};
    lock.waitLock(30000);
    const {aba}=obterAbaDoPerfil(perfil);
    const info=indicesCabecalho(aba), dados=aba.getDataRange().getValues();
    const idxCodigo=info.indices['Código'];
    if(idxCodigo===undefined) throw new Error('Cabeçalho Código não encontrado.');
    // A sequência é global: considera os códigos PRES-N de todas as abas
    // da planilha de distribuição, incluindo clientes e treinamento.
    let maiorNumero=0;
    const ss=SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    ss.getSheets().forEach(abaPlanilha=>{
      const ultimaLinha=abaPlanilha.getLastRow();
      const ultimaColuna=abaPlanilha.getLastColumn();
      if(ultimaLinha<2 || ultimaColuna<1) return;
      const valores=abaPlanilha.getRange(1,1,ultimaLinha,ultimaColuna).getValues();
      const cabecalhos=valores[0].map(v=>String(v).trim());
      const colunaCodigo=cabecalhos.indexOf('Código');
      if(colunaCodigo<0) return;
      for(let i=1;i<valores.length;i++) {
        const existente=String(valores[i][colunaCodigo]||'').trim().match(/^PRES-(\d+)$/i);
        if(existente) maiorNumero=Math.max(maiorNumero,Number(existente[1]));
      }
    });
    const codigo='PRES-'+(maiorNumero+1);
    const fname=String(arquivo.nome||''), ext=(fname.match(/\.([^.]+)$/)||[])[1];
    const extensao=String(ext||'').toLowerCase(), mimeInformado=String(arquivo.mimeType||'').toLowerCase();
    const tipos={pdf:'application/pdf',doc:'application/msword',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'};
    if(!tipos[extensao]) throw new Error('Envie um arquivo PDF, DOC ou DOCX.');
    const mime=tipos[extensao];
    if(mimeInformado && mimeInformado!==mime) throw new Error('O tipo do arquivo não corresponde à extensão.');
    const dadosCliente = (informacoes && typeof informacoes === 'object') ? informacoes : {outras:String(informacoes||'').trim()};
    const campos = [
      ['Nome Completo',dadosCliente.nome], ['Curso/Função',dadosCliente.curso],
      ['Idade',dadosCliente.idade], ['Cidade',dadosCliente.cidade], ['Estado',dadosCliente.estado], ['País',dadosCliente.pais], ['Relação com a UFRRJ?',dadosCliente.relacaoUfrrj], ['Gênero',dadosCliente.genero],
      ['Identificação Étnico-Racial',dadosCliente.etnico], ['Orientação afetiva/sexual',dadosCliente.orientacao],
      ['Como soube...',dadosCliente.comoSoube], ['Informações básicas',dadosCliente.outras]
    ];
    campos.forEach(([cabecalho])=>garantirColuna(aba,cabecalho));
    const infoAtualizada=indicesCabecalho(aba);
    const row=Array(infoAtualizada.cabecalho.length).fill('');
    const set=(h,v)=>{if(infoAtualizada.indices[h]!==undefined)row[infoAtualizada.indices[h]]=v==null?'':String(v).trim();};
    set('Código',codigo); set('Nome Completo',dadosCliente.nome||'Plantão presencial');
    campos.slice(1).forEach(([cabecalho,valor])=>set(cabecalho,valor));
    set('Local de atendimento',String(local).trim());
    set('Distribuído?','✅'); set('Plantonista',nome); set('Situação','Atendido'); set('Atendido por',nome);
    set('Data do plantão presencial',new Date()); set('Data e hora da devolutiva',new Date()); set('Data da solicitação',new Date()); set('E-mail do plantonista responsável',email||'');
    const idxRel=garantirColuna(aba,'Relatório de atendimento');
    while(row.length<aba.getLastColumn()) row.push('');
    const nomeArquivo=String(nome)+' - '+codigo+' - Relatório Plantão Psicológico da UFRRJ.'+extensao;
    const bytes=Utilities.base64Decode(String(arquivo.base64).split(',').pop());
    const file=obterOuCriarPastaRelatorios().createFile(Utilities.newBlob(bytes,mime,nomeArquivo));
    row[idxRel]=file.getUrl(); aba.appendRow(row);
    return {sucesso:true,mensagem:'Plantão presencial cadastrado com relatório. Código: '+codigo};
  } catch(e) { Logger.log(e); return {sucesso:false,mensagem:e.message||'Erro ao cadastrar plantão presencial.'}; }
  finally { try{lock.releaseLock();}catch(ignore){} }
}


// ==========================================
// DADOS DO PLANTÃO E PARTICIPANTES (BOLSISTA)
// ==========================================
function validarBolsista_(email) {
  const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_CONTAS);
  if (!sh) throw new Error('Aba de contas não encontrada.');
  const v = sh.getDataRange().getValues();
  for (let i=1;i<v.length;i++) if (String(v[i][0]).trim().toLowerCase()===String(email).trim().toLowerCase() && String(v[i][4]).trim().toLowerCase()==='bolsista' && ['sim','ativo','true'].includes(String(v[i][3]).trim().toLowerCase())) return sh;
  throw new Error('Acesso restrito ao bolsista ativo.');
}
function dadosAbasClientes_(){
 const ss=SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
 return ['Lista de clientes','Lista de clientes - treinamento'].map(n=>ss.getSheetByName(n)).filter(Boolean);
}
function statusZerado_(){return {atendidos:0,desistentes:0,transferidos:0,pendentes:0};}
function obterDadosPlantao(email){
 try {
  validarBolsista_(email);
  const agora=new Date(), inicio7=new Date(agora.getTime()-7*24*60*60*1000);
  const mes=agora.getMonth(), ano=agora.getFullYear();
  const inicioPeriodo=mes>=1&&mes<=6?new Date(ano,1,1):new Date(mes===0?ano-1:ano,7,1);
  const fimPeriodo=mes>=1&&mes<=6?new Date(ano,7,1):(mes===0?new Date(ano,1,1):new Date(ano+1,1,1));
  const semanal=statusZerado_(), periodo=statusZerado_(), total=statusZerado_();
  const statusMap={atendido:'atendidos',desistente:'desistentes',transferido:'transferidos',pendente:'pendentes'};
  dadosAbasClientes_().forEach(sh=>{
   const d=sh.getDataRange().getValues(); if(d.length<2)return;
   const h=d[0].map(x=>String(x).trim()), ix={};h.forEach((x,i)=>ix[x]=i);
   const is=ix['Situação'], id=ix['Data e hora da devolutiva'], ip=ix['Data da solicitação'], it=ix['Data do plantão presencial'];
   if(is===undefined)return;
   for(let r=1;r<d.length;r++){
    const st=String(d[r][is]).trim().toLowerCase(), key=statusMap[st];if(!key)continue;
    total[key]++;
    const dt=st==='pendente'?(ip===undefined?null:d[r][ip]):(id===undefined?null:d[r][id]);
    const data=dt instanceof Date?dt:(dt?new Date(dt):null);
    if(data&&!isNaN(data.getTime())){if(data>=inicio7&&data<=agora)semanal[key]++;if(data>=inicioPeriodo&&data<fimPeriodo)periodo[key]++;}
   }
  });
  return {sucesso:true,semanal,periodo,total,periodoInicio:inicioPeriodo.toLocaleDateString('pt-BR'),periodoFim:new Date(fimPeriodo.getTime()-86400000).toLocaleDateString('pt-BR')};
 }catch(e){return {sucesso:false,mensagem:e.message};}
}
function obterParticipantesAtivos(email){
 try{
  const sh=validarBolsista_(email), d=sh.getDataRange().getValues(), participantes=[];
  const freq=SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
  const abas=dadosAbasClientes_();
  for(let i=1;i<d.length;i++){
   const row=d[i], mail=String(row[0]||'').trim(), nome=String(row[1]||'').trim();if(!mail&&!nome)continue;
   const cabContas=d[0].map(v=>String(v).trim().toLowerCase());const idxTel=Math.max(cabContas.indexOf('telefone'),cabContas.indexOf('número de telefone'));const idxHab=cabContas.indexOf('habilitado');
   const perfil=String(row[4]||'').trim()||'Plantonista', ativo=['sim','ativo','true'].includes(String(row[3]).trim().toLowerCase());const habilitado=idxHab<0?true:!['não','nao','inabilitado','false','0'].includes(String(row[idxHab]).trim().toLowerCase());
   const item={linha:i+1,email:mail,nome,telefone:idxTel>=0?(row[idxTel]||''):'',ativo,habilitado,perfil,presencas:0,narracoes:0,atendidos:0,desistentes:0,transferidos:0,pendentes:0};
   abas.forEach(a=>{const x=a.getDataRange().getValues();if(x.length<2)return;const h=x[0].map(v=>String(v).trim()), ie=h.indexOf('E-mail do plantonista responsável'), idxNome=h.indexOf('Plantonista'), ist=h.indexOf('Situação');for(let j=1;j<x.length;j++){const match=(ie>=0&&mail&&String(x[j][ie]).trim().toLowerCase()===mail.toLowerCase())||(idxNome>=0&&nome&&String(x[j][idxNome]).trim().toLowerCase()===nome.toLowerCase());if(match&&ist>=0){const st=String(x[j][ist]).trim().toLowerCase();if(st==='atendido')item.atendidos++;else if(st==='desistente')item.desistentes++;else if(st==='transferido')item.transferidos++;else if(st==='pendente')item.pendentes++;}}});
   // Busca registros de frequência em abas com cabeçalho reconhecível.
   freq.getSheets().forEach(a=>{const x=a.getDataRange().getValues();if(!x.length)return;for(let rr=0;rr<Math.min(3,x.length);rr++){const h=x[rr].map(v=>String(v).trim().toLowerCase()), ie=h.findIndex(v=>v==='e-mail'||v==='email');if(ie<0)continue;const im=h.findIndex(v=>v==='nomes'||v==='nome'||v==='plantonista');if((mail&&String(x.slice(rr+1).find(z=>String(z[ie]).trim().toLowerCase()===mail.toLowerCase())?.[ie]||'').trim())||(im>=0&&nome)){for(let j=rr+1;j<x.length;j++){if((mail&&String(x[j][ie]).trim().toLowerCase()===mail.toLowerCase())||(im>=0&&String(x[j][im]).trim().toLowerCase()===nome.toLowerCase())){const vals=x[j];item.presencas++;if(vals.some(v=>String(v).trim().toLowerCase()==='sim'))item.narracoes++;}}break;}}});
   participantes.push(item);
  }
  return {sucesso:true,participantes};
 }catch(e){return {sucesso:false,mensagem:e.message};}
}
function atualizarParticipante(emailBolsista, linha, dados){
 try{
  const sh=validarBolsista_(emailBolsista), r=Number(linha);if(!Number.isInteger(r)||r<2||r>sh.getLastRow())throw new Error('Linha inválida.');
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v).trim()), ix={};h.forEach((v,i)=>ix[v]=i+1);
  ['E-mail','Nome','Senha','Ativo','Perfil'].forEach(k=>{if(ix[k]===undefined)throw new Error('Cabeçalho ausente: '+k)});
  sh.getRange(r,ix['E-mail']).setValue(String(dados.email||'').trim());sh.getRange(r,ix['Nome']).setValue(String(dados.nome||'').trim());
  sh.getRange(r,ix['Ativo']).setValue(dados.ativo?'Sim':'Não');sh.getRange(r,ix['Perfil']).setValue(String(dados.perfil||'Plantonista').trim());
  if(ix['Telefone']!==undefined)sh.getRange(r,ix['Telefone']).setValue(String(dados.telefone||'').trim());else if(ix['Número de telefone']!==undefined)sh.getRange(r,ix['Número de telefone']).setValue(String(dados.telefone||'').trim());
  if(ix['Habilitado']===undefined){const nova=sh.getLastColumn()+1;sh.getRange(1,nova).setValue('Habilitado');ix['Habilitado']=nova;}
  sh.getRange(r,ix['Habilitado']).setValue(dados.habilitado?'Sim':'Não');
  if(dados.senha)sh.getRange(r,ix['Senha']).setValue(String(dados.senha));
  return {sucesso:true,mensagem:'Participante atualizado.'};
 }catch(e){return {sucesso:false,mensagem:e.message};}
}


// ==========================================
// ENVIAR FOTOS
// ==========================================
// Recebe fotos em JPG/PNG (em base64, vindas do front-end) e salva na
// pasta "Fotos dos Plantonistas" no Drive. Não disponível para contas
// inativas (restritas a "Meus Plantões").

function obterOuCriarPastaFotos_() {
  const pastas = DriveApp.getFoldersByName('Fotos dos Plantonistas');
  return pastas.hasNext() ? pastas.next() : DriveApp.createFolder('Fotos dos Plantonistas');
}

function enviarFotosPlantonista(email, nome, perfil, arquivos) {

  try {
    if (!verificarHabilitacaoPlantonista_(email)) throw new Error('Acesso limitado: plantonista inabilitado.');
    if (!verificarContaAtiva_(email)) throw new Error('Sua conta está inativa. Disponível apenas: Meus Plantões.');

    if (!arquivos || !arquivos.length) {
      return { sucesso: false, mensagem: 'Selecione ao menos uma foto.' };
    }

    const permitidos = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png' };
    const pasta = obterOuCriarPastaFotos_();
    let enviados = 0;

    arquivos.forEach(arq => {
      const fname = String(arq.nome || '');
      const ext = (fname.match(/\.([^.]+)$/) || [])[1];
      const extensao = String(ext || '').toLowerCase();

      if (!permitidos[extensao]) return; // ignora arquivos que não são jpg/png

      const mime = permitidos[extensao];
      const bytes = Utilities.base64Decode(String(arq.base64).split(',').pop());
      const carimbo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss');
      const nomeArquivo = nome + ' - ' + carimbo + ' - ' + fname;

      pasta.createFile(Utilities.newBlob(bytes, mime, nomeArquivo));
      enviados++;
    });

    if (enviados === 0) {
      return { sucesso: false, mensagem: 'Nenhuma foto válida foi enviada (envie apenas JPG ou PNG).' };
    }

    return {
      sucesso: true,
      mensagem: enviados + ' foto(s) enviada(s) com sucesso! Obrigado por compartilhar 🌳'
    };

  } catch (err) {
    Logger.log('Erro em enviarFotosPlantonista: ' + err);
    return { sucesso: false, mensagem: err.message || 'Ocorreu um erro ao enviar as fotos.' };
  }
}


// ==========================================
// NOTIFICAÇÃO SEMANAL DE CLIENTE PENDENTE
// ==========================================
// Verifica clientes com Situação = "Pendente" cuja "Data da solicitação"
// já passou de 14 dias, e manda um e-mail semanal pro plantonista
// responsável lembrando de dar retorno. Usa a coluna "Última Notificação
// de Pendência" (criada automaticamente) pra não notificar mais de uma
// vez por semana sobre o mesmo cliente.
//
// Esta função sozinha NÃO roda automaticamente — é preciso rodar
// instalarGatilhoNotificacoesSemanal() uma única vez (manualmente, pelo
// editor do Apps Script) pra criar o gatilho semanal que a chama.

function construirMapaEmailPorNome_() {
  const sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_CONTAS);
  const dados = sh.getDataRange().getValues();
  const mapa = {};
  for (let i = 1; i < dados.length; i++) {
    const email = String(dados[i][0]).trim();
    const nome = String(dados[i][1]).trim().toLowerCase();
    if (nome) mapa[nome] = email;
  }
  return mapa;
}

function verificarClientesPendentesENotificar() {

  const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
  const abas = [
    planilha.getSheetByName('Lista de clientes'),
    planilha.getSheetByName('Lista de clientes - treinamento')
  ].filter(Boolean);

  const mapaEmails = construirMapaEmailPorNome_();

  const agora = new Date();
  const LIMITE_INICIAL_MS = 14 * 24 * 60 * 60 * 1000; // 2 semanas
  const INTERVALO_REENVIO_MS = 7 * 24 * 60 * 60 * 1000; // 1 semana

  abas.forEach(aba => {
    const info = indicesCabecalho(aba);
    const idx = info.indices;

    if (idx['Situação'] === undefined || idx['Data da solicitação'] === undefined ||
        idx['Plantonista'] === undefined || idx['Código'] === undefined) {
      return; // aba sem a estrutura esperada, pula
    }

    const idxUltimaNotif = garantirColuna(aba, 'Última Notificação de Pendência');
    const dados = aba.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const situacao = String(dados[i][idx['Situação']]).trim().toLowerCase();
      if (situacao !== 'pendente') continue;

      const dataSolicitacao = dados[i][idx['Data da solicitação']];
      if (!(dataSolicitacao instanceof Date)) continue;
      if (agora - dataSolicitacao < LIMITE_INICIAL_MS) continue;

      const ultimaNotif = dados[i][idxUltimaNotif];
      if (ultimaNotif instanceof Date && (agora - ultimaNotif) < INTERVALO_REENVIO_MS) continue;

      const nomePlantonista = String(dados[i][idx['Plantonista']] || '').trim();
      const emailPlantonista = mapaEmails[nomePlantonista.toLowerCase()];
      if (!emailPlantonista) continue; // não achou o e-mail desse plantonista na planilha de contas

      const codigo = dados[i][idx['Código']];
      const nomeCliente = idx['Nome Completo'] !== undefined ? dados[i][idx['Nome Completo']] : '';

      MailApp.sendEmail({
        to: emailPlantonista,
        subject: 'Retorno pendente de cliente',
        body: `Olá ${nomePlantonista},\n\n` +
          `Identificamos que a ausência do retorno sobre a cliente ${codigo} - ${nomeCliente}. ` +
          `Por favor retorne através da plataforma a situação do atendimento e, em caso de atendimento realizado, o relatório do plantão.\n\n` +
          `Caso você ainda esteja tentando agendar um horário com a cliente, por favor ignore este e-mail por enquanto e retorne o atendimento assim que possível.\n\n` +
          `Abraços,\nEquipe do Plantão Psicológico da UFRRJ`
      });

      aba.getRange(i + 1, idxUltimaNotif + 1).setValue(agora);
    }
  });
}

// Rode esta função UMA VEZ, manualmente, pelo editor do Apps Script
// (selecione "instalarGatilhoNotificacoesSemanal" no menu de funções e
// clique em Executar). Ela cria um gatilho semanal, toda segunda-feira
// às 9h, que passa a chamar verificarClientesPendentesENotificar sozinho.
function instalarGatilhoNotificacoesSemanal() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'verificarClientesPendentesENotificar') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('verificarClientesPendentesENotificar')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}


// ==========================================
// VERIFICAR PRATICANTES — RELATÓRIOS POR PERÍODO
// ==========================================
// Gera uma planilha Google nova, já filtrada por situação e por período
// (com base em "Data da solicitação"), e devolve o link pra abrir.
// mesInicio/mesFim vêm como texto "AAAA-MM" (de um <input type="month">).

function gerarRelatorioPraticantes(perfil, filtroSituacao, mesInicio, mesFim) {

  if (!ehBolsista(perfil)) {
    return { sucesso: false, mensagem: 'Acesso restrito ao bolsista.' };
  }

  try {
    const [anoIni, mesIniNum] = mesInicio.split('-').map(Number);
    const [anoFim, mesFimNum] = mesFim.split('-').map(Number);
    const dataInicio = new Date(anoIni, mesIniNum - 1, 1, 0, 0, 0);
    const dataFim = new Date(anoFim, mesFimNum, 0, 23, 59, 59); // último dia do mês final

    const planilha = SpreadsheetApp.openById(PLANILHA_DISTRIBUICAO_ID);
    const abas = [
      planilha.getSheetByName('Lista de clientes'),
      planilha.getSheetByName('Lista de clientes - treinamento')
    ].filter(Boolean);

    const linhas = [['Tipo', 'Código', 'Nome Completo', 'Plantonista', 'Situação', 'Data da Solicitação']];

    abas.forEach(aba => {
      const info = indicesCabecalho(aba);
      const idx = info.indices;
      if (idx['Situação'] === undefined || idx['Data da solicitação'] === undefined) return;

      const dados = aba.getDataRange().getValues();
      const tipoAba = aba.getName().indexOf('treinamento') !== -1 ? 'Treinamento' : 'Normal';

      for (let i = 1; i < dados.length; i++) {
        const dataSolicitacao = dados[i][idx['Data da solicitação']];
        if (!(dataSolicitacao instanceof Date)) continue;
        if (dataSolicitacao < dataInicio || dataSolicitacao > dataFim) continue;

        const situacao = String(idx['Situação'] !== undefined ? dados[i][idx['Situação']] : '').trim();
        if (filtroSituacao !== 'geral' && situacao.toLowerCase() !== filtroSituacao) continue;

        linhas.push([
          tipoAba,
          idx['Código'] !== undefined ? dados[i][idx['Código']] : '',
          idx['Nome Completo'] !== undefined ? dados[i][idx['Nome Completo']] : '',
          idx['Plantonista'] !== undefined ? dados[i][idx['Plantonista']] : '',
          situacao,
          dataSolicitacao
        ]);
      }
    });

    if (linhas.length === 1) {
      return { sucesso: false, mensagem: 'Nenhum registro encontrado para esse período/situação.' };
    }

    const rotulos = { atendido: 'Atendidos', pendente: 'Pendentes', desistente: 'Desistentes', geral: 'Geral' };
    const nomeRelatorio = 'Relatório ' + (rotulos[filtroSituacao] || filtroSituacao) + ' - ' + mesInicio + ' a ' + mesFim;

    const novaPlanilha = SpreadsheetApp.create(nomeRelatorio);
    const abaNova = novaPlanilha.getSheets()[0];
    abaNova.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);
    abaNova.getRange(1, 1, 1, linhas[0].length).setFontWeight('bold');
    abaNova.autoResizeColumns(1, linhas[0].length);

    return { sucesso: true, mensagem: 'Relatório gerado!', url: novaPlanilha.getUrl() };

  } catch (err) {
    Logger.log('Erro em gerarRelatorioPraticantes: ' + err);
    return { sucesso: false, mensagem: 'Ocorreu um erro ao gerar o relatório.' };
  }
}


// ==========================================
// MANUTENÇÃO ON/OFF
// ==========================================
// Usa PropertiesService (não precisa de aba na planilha) pra guardar
// o status. Enquanto ligada, só o bolsista consegue logar — ver a
// checagem dentro de verificarLogin.

function obterStatusManutencao() {
  try {
    const valor = PropertiesService.getScriptProperties().getProperty('MANUTENCAO');
    return { sucesso: true, ligada: valor === 'ON' };
  } catch (err) {
    Logger.log('Erro em obterStatusManutencao: ' + err);
    return { sucesso: false, mensagem: 'Erro ao verificar o status de manutenção.' };
  }
}

function definirManutencao(perfil, ligar) {
  if (!ehBolsista(perfil)) {
    return { sucesso: false, mensagem: 'Acesso restrito ao bolsista.' };
  }
  try {
    PropertiesService.getScriptProperties().setProperty('MANUTENCAO', ligar ? 'ON' : 'OFF');
    return { sucesso: true, mensagem: 'Manutenção ' + (ligar ? 'ativada' : 'desativada') + '.' };
  } catch (err) {
    Logger.log('Erro em definirManutencao: ' + err);
    return { sucesso: false, mensagem: 'Erro ao alterar o status de manutenção.' };
  }
}


// ============================================================
// API PARA GITHUB PAGES
// ============================================================
// Mantém o Apps Script como backend. O index.html hospedado no GitHub
// chama esta API; quando o HTML roda dentro do Apps Script, google.script.run
// continua sendo usado diretamente.

const API_SESSAO_SEGUNDOS = 21600; // 6 horas

function respostaApi_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function criarSessaoApi_(usuario) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(
    'api_sessao_' + token,
    JSON.stringify({
      email: String(usuario.email || '').trim().toLowerCase(),
      nome: String(usuario.nome || ''),
      perfil: String(usuario.perfil || '').trim().toLowerCase(),
      ativo: usuario.ativo !== false,
      habilitado: usuario.habilitado !== false
    }),
    API_SESSAO_SEGUNDOS
  );
  return token;
}

function obterSessaoApi_(token) {
  if (!token) return null;
  const bruto = CacheService.getScriptCache().get('api_sessao_' + token);
  if (!bruto) return null;
  try { return JSON.parse(bruto); } catch (e) { return null; }
}

function validarChamadaApi_(funcao, args, sessao) {
  if (!sessao) throw new Error('Sessão expirada. Faça login novamente.');

  const somenteBolsista = [
    'obterDadosPlantao','obterParticipantesAtivos','atualizarParticipante',
    'gerarRelatorioPraticantes','listarJustificativasPendentes',
    'decidirJustificativaFalta','definirManutencao','abrirJanelaFrequencia',
    'listarJanelasFrequencia','fecharJanelaFrequencia'
  ];
  if (somenteBolsista.includes(funcao) && sessao.perfil !== 'bolsista') {
    throw new Error('Acesso restrito à equipe administrativa.');
  }

  const funcoesComEmailPrimeiro = [
    'solicitarCliente','enviarFotosPlantonista','listarClientesPendentes',
    'atualizarSituacaoCliente','cadastrarPlantaoPresencial',
    'listarJustificativasPendentes','decidirJustificativaFalta',
    'obterResumoPlantao','solicitarDadosClientes','obterDadosPlantao',
    'obterParticipantesAtivos','atualizarParticipante','registrarFrequencia',
    'enviarJustificativaFalta'
  ];

  if (funcoesComEmailPrimeiro.includes(funcao)) {
    const emailRecebido = String(args[0] || '').trim().toLowerCase();
    if (emailRecebido !== sessao.email) {
      throw new Error('A chamada não corresponde ao usuário autenticado.');
    }
  }

  // Funções administrativas antigas que recebem perfil como primeiro argumento.
  const funcoesComPerfilPrimeiro = [
    'gerarRelatorioPraticantes','definirManutencao','abrirJanelaFrequencia',
    'listarJanelasFrequencia','fecharJanelaFrequencia'
  ];
  if (funcoesComPerfilPrimeiro.includes(funcao)) {
    args[0] = sessao.perfil;
  }
}

function executarFuncaoApi_(funcao, args) {
  switch (funcao) {
    case 'solicitarCliente': return solicitarCliente.apply(null, args);
    case 'enviarFotosPlantonista': return enviarFotosPlantonista.apply(null, args);
    case 'gerarRelatorioPraticantes': return gerarRelatorioPraticantes.apply(null, args);
    case 'obterDadosPlantao': return obterDadosPlantao.apply(null, args);
    case 'obterParticipantesAtivos': return obterParticipantesAtivos.apply(null, args);
    case 'atualizarParticipante': return atualizarParticipante.apply(null, args);
    case 'listarClientesPendentes': return listarClientesPendentes.apply(null, args);
    case 'atualizarSituacaoCliente': return atualizarSituacaoCliente.apply(null, args);
    case 'cadastrarPlantaoPresencial': return cadastrarPlantaoPresencial.apply(null, args);
    case 'listarJustificativasPendentes': return listarJustificativasPendentes.apply(null, args);
    case 'decidirJustificativaFalta': return decidirJustificativaFalta.apply(null, args);
    case 'obterStatusManutencao': return obterStatusManutencao.apply(null, args);
    case 'definirManutencao': return definirManutencao.apply(null, args);
    case 'abrirJanelaFrequencia': return abrirJanelaFrequencia.apply(null, args);
    case 'listarJanelasFrequencia': return listarJanelasFrequencia.apply(null, args);
    case 'fecharJanelaFrequencia': return fecharJanelaFrequencia.apply(null, args);
    case 'listarJanelasAbertas': return listarJanelasAbertas.apply(null, args);
    case 'registrarFrequencia': return registrarFrequencia.apply(null, args);
    case 'listarTodasJanelas': return listarTodasJanelas.apply(null, args);
    case 'enviarJustificativaFalta': return enviarJustificativaFalta.apply(null, args);
    case 'obterResumoPlantao': return obterResumoPlantao.apply(null, args);
    case 'solicitarDadosClientes': return solicitarDadosClientes.apply(null, args);
    default: throw new Error('Função não autorizada pela API.');
  }
}

function doPost(e) {
  try {
    const corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const funcao = String(corpo.funcao || '');
    const args = Array.isArray(corpo.argumentos) ? corpo.argumentos : [];

    if (funcao === 'verificarLogin') {
      const resultado = verificarLogin.apply(null, args);
      if (resultado && resultado.sucesso) {
        const token = criarSessaoApi_(resultado);
        return respostaApi_({ ok: true, resultado: resultado, token: token });
      }
      return respostaApi_({ ok: true, resultado: resultado });
    }

    const sessao = obterSessaoApi_(corpo.token);
    validarChamadaApi_(funcao, args, sessao);
    const resultado = executarFuncaoApi_(funcao, args);

    // Renova a sessão a cada chamada válida.
    CacheService.getScriptCache().put(
      'api_sessao_' + corpo.token,
      JSON.stringify(sessao),
      API_SESSAO_SEGUNDOS
    );

    return respostaApi_({ ok: true, resultado: resultado });

  } catch (erro) {
    return respostaApi_({
      ok: false,
      erro: erro && erro.message ? erro.message : String(erro)
    });
  }
}
