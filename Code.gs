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

      // Senha errada
      if (senhaPlanilha !== senhaDigitada) {
        return {
          sucesso: false,
          mensagem: 'E-mail ou senha incorretos.'
        };
      }

      // Conta inativa
      const contaAtiva = (
        ativo === 'sim' ||
        ativo === 'ativo' ||
        ativo === 'true'
      );

      if (!contaAtiva) {
        return {
          sucesso: false,
          mensagem: 'Esta conta está desativada. Fale com a coordenação do plantão.'
        };
      }

      // Login válido
      return {
        sucesso: true,
        nome: nome,
        email: emailPlanilha,
        perfil: perfil
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

function solicitarCliente(email, nome, perfil) {

  try {
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

function obterResumoPlantao(nome, perfil) {

  try {
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
