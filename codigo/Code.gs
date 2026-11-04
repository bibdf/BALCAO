/**
 * =========================================================================
 *  BIBLIOTECA ESCOLAR — EB D. Fernando II · 2026/2027
 *  Code.gs (Google Apps Script) — painel de estatísticas, avisos aos DT
 *  e dados de teste.
 *  Conta: bibdf@essmaria.net · Fuso horário: Europe/Lisbon
 *  Powered by Claude IA & Ideias&Soltas — concepção de produtos multimédia
 * =========================================================================
 *
 *  INSTALAÇÃO
 *  1. Na folha de cálculo da biblioteca: Extensões → Apps Script.
 *  2. Apagar tudo em Code.gs e colar ESTE ficheiro.
 *  3. "+" → HTML → nome: Index (sem .html) → colar o Index.html.
 *  4. Definições do projeto (roda dentada) → Fuso horário: Europe/Lisbon.
 *  5. Guardar. Recarregar a folha: aparece o menu "Biblioteca".
 *
 *  DADOS DE TESTE
 *  - Pôr BLOQUEIO_SEGURANCA = false, guardar, menu Biblioteca →
 *    "Criar dados de teste". Depois voltar a pôr true e guardar.
 *
 *  PAINEL E BALCÃO (mesma aplicação web)
 *  - Implementar → Nova implementação → Aplicação Web:
 *      Executar como: Eu · Quem tem acesso: só a equipa da biblioteca.
 *  - Endereço sem nada no fim → painel de estatísticas.
 *  - Endereço com ?p=balcao no fim → app do balcão (requisições,
 *    atividade, reserva do espaço). Precisa do ficheiro HTML "Balcao".
 *
 *  CÓDIGO DE ACESSO
 *  - As duas páginas pedem um código. Está na folha "Config" (criada
 *    sozinha na primeira vez); para o mudar, escrever outro nessa célula.
 *
 *  CÓPIA DE SEGURANÇA
 *  - Correr uma vez a função ativarProtecoes: protege os títulos das
 *    folhas e faz uma cópia da folha de cálculo todas as sextas-feiras.
 *
 *  PASSAR A DADOS REAIS
 *  - Confirmar BLOQUEIO_SEGURANCA = true e mudar MODO_TESTE_EMAIL = false.
 * =========================================================================
 */

// ---------------------------------------------------------------------------
// Configuração
// ---------------------------------------------------------------------------
const BLOQUEIO_SEGURANCA = true;   // true = impede apagar a base de dados por engano
const MODO_TESTE_EMAIL = true;     // true = os avisos vão só para a biblioteca
const EMAIL_BIBLIOTECA = "bibdf@essmaria.net";
const TIMEZONE = "Europe/Lisbon";
const DIA_MS = 24 * 60 * 60 * 1000;
// Folha de cálculo "Biblioteca Escolar — Gestão 2026/27" (conta bibdf@essmaria.net)
const ID_FOLHA = "1RXJ7Hui0IV6ZzQWzU7QPdgXzyAUiqr6nsk1otqZ7sUo";

// Abre a folha da biblioteca (funciona no editor, no painel e no acionador semanal).
function abrirFolha_() {
  return SpreadsheetApp.openById(ID_FOLHA);
}

// ---------------------------------------------------------------------------
// Menu na folha de cálculo
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Biblioteca")
    .addItem("Criar dados de teste", "inicializarBaseDeDadosCompleta")
    .addSeparator()
    .addItem("Ativar envio semanal aos DT (segunda, 8h)", "configurarAcionadorSemanal")
    .addItem("Enviar avisos aos DT agora", "enviarAvisosSemanaisAtrasosDT")
    .addToUi();
}

// ---------------------------------------------------------------------------
// Aplicação web (painel)
// ---------------------------------------------------------------------------
// Sem parâmetros abre o painel de estatísticas; com ?p=balcao abre a app da equipa (Balcão).
function doGet(e) {
  const balcao = !!(e && e.parameter && e.parameter.p === "balcao");
  const t = HtmlService.createTemplateFromFile(balcao ? "Balcao" : "Index");
  t.urlBase = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle(balcao ? "Biblioteca — Balcão" : "Biblioteca Escolar — EB D. Fernando II")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ---------------------------------------------------------------------------
// Auxiliares de datas
// ---------------------------------------------------------------------------
// Diferença em dias de calendário, imune à mudança da hora.
function calcularDiferencaDias(dataMaior, dataMenor) {
  const u1 = Date.UTC(dataMaior.getFullYear(), dataMaior.getMonth(), dataMaior.getDate());
  const u2 = Date.UTC(dataMenor.getFullYear(), dataMenor.getMonth(), dataMenor.getDate());
  return Math.round((u1 - u2) / DIA_MS);
}

// Converte o valor de uma célula numa data (ou null se estiver vazia/inválida).
function paraData(v) {
  if (v === "" || v === null || v === undefined) return null;
  const d = (v instanceof Date) ? new Date(v.getTime()) : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// Segunda-feira da semana de uma data, às 00:00.
function segundaDaSemana(d) {
  const r = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const dia = r.getDay(); // 0 = domingo
  r.setDate(r.getDate() + (dia === 0 ? -6 : 1 - dia));
  return r;
}

function calcularIntervalo(periodo, dataInicioCustom, dataFimCustom) {
  const hoje = new Date();
  const h = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  let inicio, fim, rotulo;

  if (periodo === "este_mes") {
    inicio = new Date(h.getFullYear(), h.getMonth(), 1);
    fim = new Date(h.getFullYear(), h.getMonth() + 1, 0, 23, 59, 59, 999);
    rotulo = "Relatório mensal";
  } else if (periodo === "ano_letivo") {
    const ano = h.getMonth() < 8 ? h.getFullYear() - 1 : h.getFullYear(); // 8 = setembro
    inicio = new Date(ano, 8, 1);
    fim = new Date(ano + 1, 7, 31, 23, 59, 59, 999);
    rotulo = "Relatório anual";
  } else if (periodo === "personalizado" && dataInicioCustom && dataFimCustom) {
    const pi = String(dataInicioCustom).split("-").map(Number); // "aaaa-mm-dd"
    const pf = String(dataFimCustom).split("-").map(Number);
    inicio = new Date(pi[0], pi[1] - 1, pi[2]);
    fim = new Date(pf[0], pf[1] - 1, pf[2], 23, 59, 59, 999);
    if (fim < inicio) { const t = inicio; inicio = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate()); fim = new Date(t.getFullYear(), t.getMonth(), t.getDate(), 23, 59, 59, 999); }
    rotulo = "Relatório de período";
  } else {
    inicio = segundaDaSemana(h);
    fim = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate() + 6, 23, 59, 59, 999);
    rotulo = "Relatório semanal";
  }
  return { hoje: h, inicio: inicio, fim: fim, rotulo: rotulo };
}

function lerFolha(ss, nome) {
  const sh = ss.getSheetByName(nome);
  if (!sh) throw new Error("Falta a folha «" + nome + "».");
  return sh.getDataRange().getValues();
}

function eVerdadeiro(v) {
  return v === true || String(v).toUpperCase() === "TRUE" || String(v).toUpperCase() === "VERDADEIRO";
}
// N.º de prorrogações de uma requisição (registos antigos têm VERDADEIRO = 1; agora guarda-se 1, 2 ou 3).
function numProrrog_(v) {
  if (eVerdadeiro(v)) return 1;
  const n = parseInt(v, 10);
  return n > 0 ? n : 0;
}

// ---------------------------------------------------------------------------
// Estatísticas para o painel (o Index.html chama painel(), que confirma o código de acesso)
// ---------------------------------------------------------------------------
function getEstatisticasPainel_(periodo, dataInicioCustom, dataFimCustom) {
  const ss = abrirFolha_();
  const I = calcularIntervalo(periodo, dataInicioCustom, dataFimCustom);
  const hoje = I.hoje;

  function dentroPeriodo(v) {
    const d = paraData(v);
    return !!d && d >= I.inicio && d <= I.fim;
  }

  // Leitores e livros
  const dadosLei = lerFolha(ss, "Leitores");
  const leitoresMap = {};
  for (let i = 1; i < dadosLei.length; i++) {
    leitoresMap[dadosLei[i][0]] = { nome: dadosLei[i][1], turma: dadosLei[i][3], num: dadosLei[i][4], email: dadosLei[i][5] };
  }
  const dadosLiv = lerFolha(ss, "Livros");
  const livrosMap = {};
  for (let i = 1; i < dadosLiv.length; i++) livrosMap[dadosLiv[i][0]] = dadosLiv[i][1];

  // Presenças
  const dadosPres = lerFolha(ss, "Presencas");
  let totalPresencas = 0;
  const presPorAtividade = {}, presPorTurma = {}, usoComputadores = {};
  for (let i = 1; i < dadosPres.length; i++) {
    if (!dentroPeriodo(dadosPres[i][1])) continue;
    totalPresencas++;
    const t = dadosPres[i][3] || "Sem turma";
    const a = dadosPres[i][5] || "Outro";
    presPorTurma[t] = (presPorTurma[t] || 0) + 1;
    presPorAtividade[a] = (presPorAtividade[a] || 0) + 1;
    const pc = dadosPres[i][6];
    if (pc !== "" && pc !== undefined && pc !== null) usoComputadores["PC " + pc] = (usoComputadores["PC " + pc] || 0) + 1;
  }

  // Requisições
  const dadosReq = lerFolha(ss, "Requisicoes");
  let reqsFeitasPeriodo = 0, reqsDevolvidasPeriodo = 0, prorrogacoesPeriodo = 0, totalPorEntregar = 0;
  const listaAtrasos = [];
  const reqPorTurma = {}, reqPorLivro = {}, reqPorLeitor = {};
  for (let i = 1; i < dadosReq.length; i++) {
    const r = dadosReq[i];
    if (!r[0]) continue;
    const nProrrog = numProrrog_(r[6]), prorrogado = nProrrog > 0;
    const dataLimite = paraData(r[7]);
    const dataDev = paraData(r[8]);

    if (dentroPeriodo(r[4])) {
      reqsFeitasPeriodo++;
      const lt = leitoresMap[r[2]];
      const turmaReq = (lt && lt.turma) || r[1] || "Sem turma";
      const tituloReq = String(livrosMap[r[3]] || r[3]);
      const quemReq = (lt ? lt.nome : String(r[2])) + " (" + turmaReq + ")";
      reqPorTurma[turmaReq] = (reqPorTurma[turmaReq] || 0) + 1;
      reqPorLivro[tituloReq] = (reqPorLivro[tituloReq] || 0) + 1;
      reqPorLeitor[quemReq] = (reqPorLeitor[quemReq] || 0) + 1;
    }
    // A app não guarda a data da prorrogação: conta-se pela data de entrega original.
    if (prorrogado && dentroPeriodo(r[5])) prorrogacoesPeriodo += nProrrog;
    if (dataDev && dentroPeriodo(dataDev)) reqsDevolvidasPeriodo++;

    if (!dataDev && dataLimite) {
      const dias = calcularDiferencaDias(hoje, dataLimite);
      if (dias > 0) {
        const l = leitoresMap[r[2]] || { nome: "Desconhecido", turma: r[1], num: "" };
        listaAtrasos.push({
          idReq: r[0],
          turma: l.turma || r[1],
          aluno: l.nome,
          num: l.num,
          livro: livrosMap[r[3]] || r[3],
          limite: Utilities.formatDate(dataLimite, TIMEZONE, "dd/MM/yyyy"),
          dias: dias,
          prorrogado: prorrogado
        });
      } else {
        totalPorEntregar++;
      }
    }
  }
  listaAtrasos.sort(function (a, b) { return b.dias - a.dias; });
  // Só os 10 primeiros de um dicionário {nome: quantidade}.
  function primeiros(dic, n) {
    const out = {};
    Object.keys(dic).sort(function (a, b) { return (dic[b] - dic[a]) || a.localeCompare(b, "pt"); }).slice(0, n).forEach(function (k) { out[k] = dic[k]; });
    return out;
  }

  // Reservas de espaço
  const dadosRes = lerFolha(ss, "Reservas");
  let totalReservas = 0;
  const reservasPorProfessor = {};
  for (let i = 1; i < dadosRes.length; i++) {
    if (!dentroPeriodo(dadosRes[i][1])) continue;
    totalReservas++;
    const profId = dadosRes[i][4];
    const profNome = leitoresMap[profId] ? leitoresMap[profId].nome : profId;
    const chave = profNome + (dadosRes[i][5] ? " (" + dadosRes[i][5] + ")" : "");
    reservasPorProfessor[chave] = (reservasPorProfessor[chave] || 0) + 1;
  }

  // Material
  const dadosMat = lerFolha(ss, "Material");
  const matTipoMap = {};
  for (let i = 1; i < dadosMat.length; i++) matTipoMap[dadosMat[i][0]] = dadosMat[i][2];

  const dadosEmp = lerFolha(ss, "Emprestimos_Material");
  let materialFora = 0, totalEmprestimosPeriodo = 0, materialDanificadoPeriodo = 0;
  const materialPorTipo = {};
  for (let i = 1; i < dadosEmp.length; i++) {
    const e = dadosEmp[i];
    if (!e[0]) continue;
    const tipo = matTipoMap[e[3]] || "Outro";
    const dataDev = paraData(e[5]);
    if (dentroPeriodo(e[4])) {
      totalEmprestimosPeriodo++;
      materialPorTipo[tipo] = (materialPorTipo[tipo] || 0) + 1;
    }
    if (!dataDev) materialFora++;
    if (e[6] === "Danificado" && dataDev && dentroPeriodo(dataDev)) materialDanificadoPeriodo++;
  }

  return {
    intervalo: {
      inicio: Utilities.formatDate(I.inicio, TIMEZONE, "dd/MM/yyyy"),
      fim: Utilities.formatDate(I.fim, TIMEZONE, "dd/MM/yyyy"),
      tipoRotulo: I.rotulo
    },
    resumo: {
      totalPresencas: totalPresencas,
      totalPorEntregar: totalPorEntregar,
      totalAtrasados: listaAtrasos.length,
      materialFora: materialFora,
      totalReservas: totalReservas,
      reqsFeitasPeriodo: reqsFeitasPeriodo,
      reqsDevolvidasPeriodo: reqsDevolvidasPeriodo,
      prorrogacoesPeriodo: prorrogacoesPeriodo,
      totalEmprestimosPeriodo: totalEmprestimosPeriodo,
      materialDanificadoPeriodo: materialDanificadoPeriodo
    },
    presPorTurma: presPorTurma,
    presPorAtividade: presPorAtividade,
    reservasPorProfessor: reservasPorProfessor,
    materialPorTipo: materialPorTipo,
    usoComputadores: usoComputadores,
    reqPorTurma: reqPorTurma,
    livrosMaisRequisitados: primeiros(reqPorLivro, 10),
    leitoresMaisAtivos: primeiros(reqPorLeitor, 10),
    listaAtrasos: listaAtrasos
  };
}

// ---------------------------------------------------------------------------
// Avisos semanais aos Diretores de Turma
// ---------------------------------------------------------------------------
function configurarAcionadorSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "enviarAvisosSemanaisAtrasosDT") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("enviarAvisosSemanaisAtrasosDT")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(0)
    .inTimezone(TIMEZONE)
    .create();
  avisar_("Envio semanal ativado: todas as segundas-feiras por volta das 8h.");
}

function enviarAvisosSemanaisAtrasosDT() {
  const ss = abrirFolha_();
  const dadosReq = lerFolha(ss, "Requisicoes");
  const dadosLei = lerFolha(ss, "Leitores");
  const dadosTur = lerFolha(ss, "Turmas");
  const dadosLiv = lerFolha(ss, "Livros");

  const leitoresMap = {};
  for (let i = 1; i < dadosLei.length; i++) {
    leitoresMap[dadosLei[i][0]] = { nome: dadosLei[i][1], turma: dadosLei[i][3], num: dadosLei[i][4], email: String(dadosLei[i][5] || "").trim() };
  }
  const livrosMap = {};
  for (let i = 1; i < dadosLiv.length; i++) livrosMap[dadosLiv[i][0]] = dadosLiv[i][1];

  const dtEmailMap = {};
  for (let i = 1; i < dadosTur.length; i++) {
    const dt = leitoresMap[dadosTur[i][2]];
    dtEmailMap[dadosTur[i][0]] = dt ? dt.email : "";
  }

  const agora = new Date();
  const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
  const atrasosPorTurma = {};

  for (let i = 1; i < dadosReq.length; i++) {
    const r = dadosReq[i];
    if (!r[0] || paraData(r[8])) continue;
    const dataLimite = paraData(r[7]);
    if (!dataLimite) continue;
    const dias = calcularDiferencaDias(hoje, dataLimite);
    if (dias <= 0) continue;
    const leitor = leitoresMap[r[2]];
    if (!leitor || !leitor.turma) continue;
    (atrasosPorTurma[leitor.turma] = atrasosPorTurma[leitor.turma] || []).push({
      aluno: leitor.nome,
      num: leitor.num,
      livro: livrosMap[r[3]] || r[3],
      limite: Utilities.formatDate(dataLimite, TIMEZONE, "dd/MM/yyyy"),
      dias: dias
    });
  }

  let enviados = 0;
  const semDT = [];
  Object.keys(atrasosPorTurma).sort().forEach(function (turma) {
    const emailDT = dtEmailMap[turma];
    if (!emailDT || emailDT.indexOf("@") < 0) { semDT.push(turma); return; }
    const lista = atrasosPorTurma[turma].sort(function (a, b) { return b.dias - a.dias; });
    const linhas = lista.map(function (a) {
      return "• " + a.aluno + " (n.º " + a.num + "): «" + a.livro + "» — data limite " + a.limite + " (" + a.dias + (a.dias === 1 ? " dia" : " dias") + " de atraso)";
    }).join("\n");

    let assunto = "[Biblioteca Escolar] Alunos com livros em atraso — turma " + turma;
    if (MODO_TESTE_EMAIL) assunto = "[TESTE · seria para " + emailDT + "] " + assunto;

    const corpo =
      "Caro(a) Diretor(a) de Turma do " + turma + ",\n\n" +
      "Segue a lista dos alunos da sua turma com livros da Biblioteca Escolar em atraso, à data de " +
      Utilities.formatDate(hoje, TIMEZONE, "dd/MM/yyyy") + ":\n\n" +
      linhas + "\n\n" +
      "Agradecemos a sua colaboração para que os alunos devolvam os livros com a maior brevidade.\n\n" +
      "Com os melhores cumprimentos,\nEquipa da Biblioteca Escolar\n" + EMAIL_BIBLIOTECA;

    const opcoes = { name: "Biblioteca Escolar", replyTo: EMAIL_BIBLIOTECA };
    if (MODO_TESTE_EMAIL) {
      GmailApp.sendEmail(EMAIL_BIBLIOTECA, assunto, corpo, opcoes);
    } else {
      opcoes.cc = EMAIL_BIBLIOTECA;
      GmailApp.sendEmail(emailDT, assunto, corpo, opcoes);
    }
    enviados++;
  });

  avisar_(enviados + (enviados === 1 ? " aviso enviado." : " avisos enviados.") +
    (MODO_TESTE_EMAIL ? " (Modo de teste: tudo para " + EMAIL_BIBLIOTECA + ".)" : "") +
    (semDT.length ? " Turmas com atrasos mas sem email de DT: " + semDT.join(", ") + "." : ""));
}

// Mensagem na folha quando corrido pelo menu; em registo quando corrido pelo acionador.
function avisar_(msg) {
  try { abrirFolha_().toast(msg, "Biblioteca", 8); } catch (e) {}
  console.log(msg);
}

// ---------------------------------------------------------------------------
// Dados de teste (APAGA as folhas — só com BLOQUEIO_SEGURANCA = false)
// ---------------------------------------------------------------------------
function inicializarBaseDeDadosCompleta() {
  if (BLOQUEIO_SEGURANCA) {
    throw new Error("Ação bloqueada: BLOQUEIO_SEGURANCA está a true. Para recriar os dados de teste, mude para false, guarde, corra de novo e volte a pôr true.");
  }
  const ss = abrirFolha_();
  const agora = new Date();
  const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());

  // Data "n dias atrás" (negativo = no futuro), por calendário: imune à mudança da hora.
  function diaRel(diasAtras) {
    return new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - diasAtras, 12); // meio-dia: a diferença de fuso entre o script e a folha não muda o dia
  }
  function diaHora(diasAtras, hh, mm) {
    return new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - diasAtras, hh, mm);
  }
  function preencher(nome, cabecalho, linhas) {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome); else sh.clear();
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
      .setFontWeight("bold").setBackground("#e6eef8");
    sh.setFrozenRows(1);
    if (linhas.length) sh.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
    return sh;
  }

  // Turmas
  const turmas = [["5A", 5, "PROF-01"], ["6B", 6, "PROF-02"], ["7A", 7, "PROF-03"],
                  ["7B", 7, "PROF-04"], ["8A", 8, "PROF-05"], ["9C", 9, "PROF-06"]];
  preencher("Turmas", ["Turma", "Ano", "DT"], turmas);

  // Leitores: 12 professores (6 DT com email fictício) + 120 alunos
  const profs = ["Ana Maria Santos", "Carlos Manuel Pereira", "Isabel Sofia Ramos", "João Pedro Ferreira",
                 "Maria Helena Carvalho", "Rui Jorge Oliveira", "Teresa Cristina Lucas", "António José Silva",
                 "Paula Cristina Pinto", "Fernando Jorge Rocha", "Marta Susana Rodrigues", "Luís Miguel Gonçalves"];
  const leitores = profs.map(function (n, i) {
    const id = "PROF-" + ("0" + (i + 1)).slice(-2);
    return [id, n, "Professor", "", "", i < 6 ? "dt" + (i + 1) + "@exemplo.pt" : ""];
  });
  const nomes = ["Afonso", "Beatriz", "Bernardo", "Carolina", "Diogo", "Duarte", "Francisca", "Gabriel", "Gonçalo", "Inês",
                 "Leonor", "Lourenço", "Manuel", "Margarida", "Mariana", "Martim", "Matilde", "Rodrigo", "Santiago", "Tomás"];
  const apelidos = ["Silva", "Santos", "Ferreira", "Pereira", "Oliveira", "Costa", "Rodrigues", "Martins", "Jesus", "Sousa",
                    "Fernandes", "Gonçalves", "Gomes", "Lopes", "Marques", "Alves", "Ribeiro", "Pinto", "Carvalho", "Teixeira"];
  let k = 1;
  turmas.forEach(function (t) {
    for (let num = 1; num <= 20; num++, k++) {
      leitores.push(["ALU-" + ("0000" + k).slice(-4), nomes[(k - 1) % 20] + " " + apelidos[(k * 3) % 20], "Aluno", t[0], num, ""]);
    }
  });
  preencher("Leitores", ["ID_Leitor", "Nome", "Tipo", "Turma", "Numero", "Email"], leitores);

  // Livros: 001–012 requisitados, restantes disponíveis
  const titulos = [
    ["O Cavaleiro da Dinamarca", "Sophia de Mello Breyner Andresen", "82-3 BRE"],
    ["A Fada Oriana", "Sophia de Mello Breyner Andresen", "82-3 BRE"],
    ["O Principezinho", "Antoine de Saint-Exupéry", "82-3 SAI"],
    ["História de uma Gaivota e do Gato que a Ensinou a Voar", "Luis Sepúlveda", "82-3 SEP"],
    ["A Menina do Mar", "Sophia de Mello Breyner Andresen", "82-3 BRE"],
    ["Ulisses", "Maria Alberta Menéres", "82-3 MEN"],
    ["A Viúva e o Papagaio", "Virginia Woolf", "82-3 WOO"],
    ["Sexta-Feira ou a Vida Selvagem", "Michel Tournier", "82-3 TOU"],
    ["O Velho e o Mar", "Ernest Hemingway", "82-3 HEM"],
    ["Diário de Anne Frank", "Anne Frank", "82-94 FRA"],
    ["Os Maias (adaptação juvenil)", "Eça de Queirós", "82-3 QUE"],
    ["Auto da Barca do Inferno", "Gil Vicente", "82-2 VIC"],
    ["O Rapaz do Pijama às Riscas", "John Boyne", "82-3 BOY"],
    ["Percy Jackson e o Ladrão de Raios", "Rick Riordan", "82-3 RIO"],
    ["Harry Potter e a Pedra Filosofal", "J.K. Rowling", "82-3 ROW"],
    ["Harry Potter e a Câmara dos Segredos", "J.K. Rowling", "82-3 ROW"],
    ["O Hobbit", "J.R.R. Tolkien", "82-3 TOL"],
    ["As Crónicas de Nárnia", "C.S. Lewis", "82-3 LEW"],
    ["Uma Aventura na Cidade", "Ana Maria Magalhães e Isabel Alçada", "82-3 MAG"],
    ["Uma Aventura na Escola", "Ana Maria Magalhães e Isabel Alçada", "82-3 MAG"]
  ];
  const livros = [];
  for (let i = 1; i <= 40; i++) {
    const b = titulos[(i - 1) % 20];
    livros.push(["LIV-" + ("000" + i).slice(-3), b[0] + (i > 20 ? " (exemplar 2)" : ""), b[1], b[2], i <= 12 ? "Requisitado" : "Disponível"]);
  }
  preencher("Livros", ["ID_Livro", "Titulo", "Autor", "Cota", "Estado"], livros);

  // Requisições: [id, turma, aluno, livro, dias desde a requisição, prorrogado, dias desde a devolução ou null]
  const reqDef = [
    ["REQ-001", "5A", "ALU-0001", "LIV-001", 60, false, null], // > 30 dias de atraso
    ["REQ-002", "7A", "ALU-0041", "LIV-002", 55, true,  null],
    ["REQ-003", "8A", "ALU-0081", "LIV-003", 52, false, null],
    ["REQ-004", "5A", "ALU-0005", "LIV-004", 25, false, null], // atraso recente
    ["REQ-005", "6B", "ALU-0022", "LIV-005", 22, true,  null],
    ["REQ-006", "7B", "ALU-0065", "LIV-006", 18, false, null],
    ["REQ-007", "9C", "ALU-0102", "LIV-007", 28, true,  null],
    ["REQ-008", "5A", "ALU-0010", "LIV-008", 5,  false, null], // dentro do prazo
    ["REQ-009", "6B", "ALU-0025", "LIV-009", 10, false, null],
    ["REQ-010", "7A", "ALU-0045", "LIV-010", 14, true,  null],
    ["REQ-011", "8A", "ALU-0085", "LIV-011", 2,  false, null],
    ["REQ-012", "9C", "ALU-0110", "LIV-012", 14, false, null], // termina amanhã
    ["REQ-013", "5A", "ALU-0003", "LIV-013", 30, false, 18],   // devolvidos
    ["REQ-014", "6B", "ALU-0028", "LIV-014", 25, true,  7],
    ["REQ-015", "7A", "ALU-0050", "LIV-015", 20, false, 6],
    ["REQ-016", "8A", "ALU-0088", "LIV-016", 15, false, 1]
  ];
  const reqs = reqDef.map(function (r) {
    const entrega = r[4] - 15;                 // +15 dias corridos
    const limite = r[5] ? entrega - 5 : entrega; // +5 se prorrogado
    return [r[0], r[1], r[2], r[3], diaRel(r[4]), diaRel(entrega), r[5], diaRel(limite), r[6] === null ? "" : diaRel(r[6])];
  });
  preencher("Requisicoes", ["ID_Requisicao", "Turma", "Leitor", "Livro", "Data_Requisicao", "Data_Entrega", "Prorrogado", "Data_Limite", "Data_Devolucao"], reqs)
    .getRange("E:I").setNumberFormat("dd/mm/yyyy");
  ss.getSheetByName("Requisicoes").getRange("G:G").setNumberFormat("General");

  // Tempos (hora de início e de fim separadas, para o calendário da AppSheet)
  preencher("Tempos", ["ID_Tempo", "Ordem", "Designacao", "Hora_Inicio", "Hora_Fim"], [
    ["T1", 1, "1.º tempo", "08:15", "09:05"], ["T2", 2, "2.º tempo", "09:05", "09:55"],
    ["T3", 3, "3.º tempo", "10:25", "11:15"], ["T4", 4, "4.º tempo", "11:15", "12:05"],
    ["T5", 5, "5.º tempo", "12:15", "13:05"], ["T6", 6, "6.º tempo", "13:20", "14:10"],
    ["T7", 7, "7.º tempo", "14:10", "15:00"], ["T8", 8, "8.º tempo", "15:10", "16:00"],
    ["T9", 9, "9.º tempo", "16:20", "17:10"], ["T10", 10, "10.º tempo", "17:10", "18:00"]
  ]).getRange("D:E").setNumberFormat("HH:mm");

  // Espaços
  preencher("Espacos", ["ID_Espaco", "Nome_Espaco", "Capacidade", "Descricao"], [
    ["ESP-01", "Zona de leitura e pesquisa", 30, "Mesas para leitura e estudo"],
    ["ESP-02", "Sala multimédia / audiovisuais", 20, "Computadores e projetor"]
  ]);

  // Reservas (inclui uma no futuro dentro da semana)
  preencher("Reservas", ["ID_Reserva", "Data", "Tempo", "Espaco", "Professor", "Turma", "Finalidade"], [
    ["RES-01", diaRel(1),  "T2", "ESP-01", "PROF-01", "5A", "Pesquisa de História"],
    ["RES-02", diaRel(2),  "T3", "ESP-02", "PROF-03", "7A", "Apresentação oral"],
    ["RES-03", diaRel(-2), "T4", "ESP-01", "PROF-05", "8A", "Leitura orientada"],
    ["RES-04", diaRel(10), "T1", "ESP-02", "PROF-02", "6B", "Trabalho de Ciências"],
    ["RES-05", diaRel(15), "T5", "ESP-01", "PROF-04", "7B", "Projeto de Cidadania"]
  ]).getRange("B:B").setNumberFormat("dd/mm/yyyy");

  // Material
  preencher("Material", ["ID_Material", "Num_Inventario", "Tipo", "Estado"], [
    ["MAT-01", "INV-F01", "Fones", "Disponível"], ["MAT-02", "INV-F02", "Fones", "Emprestado"],
    ["MAT-03", "INV-T01", "Tablet", "Disponível"], ["MAT-04", "INV-T02", "Tablet", "Emprestado"],
    ["MAT-05", "INV-C01", "Cabo", "Avariado"],     ["MAT-06", "INV-O01", "Outro", "Disponível"]
  ]);

  // Empréstimos de material
  preencher("Emprestimos_Material", ["ID_Emprestimo", "Turma", "Leitor", "Material", "DataHora_Saida", "DataHora_Devolucao", "Estado_Devolucao", "Observacoes"], [
    ["EMP-01", "5A", "ALU-0004", "MAT-02", diaHora(0, 9, 15),  "", "", "Uso na sala de estudo"],
    ["EMP-02", "7B", "ALU-0062", "MAT-04", diaHora(1, 10, 30), "", "", "Trabalho de TIC"],
    ["EMP-03", "6B", "ALU-0024", "MAT-01", diaHora(4, 11, 20), diaHora(4, 12, 0), "OK", "Devolvido sem problemas"],
    ["EMP-04", "8A", "ALU-0083", "MAT-05", diaHora(8, 14, 15), diaHora(8, 15, 0), "Danificado", "Ficha do cabo partida"]
  ]).getRange("E:F").setNumberFormat("dd/mm/yyyy HH:mm");

  // Presenças: só dias úteis das últimas 4 semanas, alunos certos de cada turma
  const atividades = ["Estudar", "Trabalho de grupo", "Desenhar", "Utilizar computador", "Ler"];
  const pres = [];
  let pid = 1;
  for (let d = 0; d < 28; d++) {
    const dia = diaRel(d);
    const ds = dia.getDay();
    if (ds === 0 || ds === 6) continue;
    turmas.forEach(function (t, tIdx) {
      const base = tIdx * 20 + 1; // 5A: 1–20, 6B: 21–40, 7A: 41–60, 7B: 61–80, 8A: 81–100, 9C: 101–120
      const n = (d % 2 === 0) ? 2 : 1;
      for (let reg = 0; reg < n; reg++) {
        const aluno = "ALU-" + ("0000" + (base + (reg * 5 + d) % 20)).slice(-4);
        pres.push(["PRE-" + ("0000" + pid++).slice(-4), dia, "10:30", t[0], aluno, atividades[(d + reg + tIdx) % 5]]);
      }
    });
  }
  const shP = preencher("Presencas", ["ID_Presenca", "Data", "Hora", "Turma", "Leitor", "Atividade"], pres);
  shP.getRange("B:B").setNumberFormat("dd/mm/yyyy");
  shP.getRange("C:C").setNumberFormat("HH:mm");

  SpreadsheetApp.flush();
  avisar_("Dados de teste criados: 6 turmas, 120 alunos, 12 professores, 16 requisições, " + pres.length + " presenças. Volte a pôr BLOQUEIO_SEGURANCA = true.");
}

// ===========================================================================
// BALCÃO — app da equipa da biblioteca (?p=balcao)
// Requisições domiciliárias, atividade na biblioteca e reserva do espaço.
// Escreve nas mesmas folhas que o painel lê, pela mesma ordem de colunas.
// ===========================================================================
const ATIVIDADES = ["Usar computadores", "Estudar", "Jogar", "Fazer desenhos", "Fazer trabalhos de grupo", "Conviver na biblioteca"];
const DIAS_EMPRESTIMO = 15;
const DIAS_PRORROGACAO = 5;
const MAX_PRORROGACOES = 3;    // até três pedidos de prorrogação, de 5 dias cada
const DIAS_HISTORICO = 30;     // o balcão carrega as devoluções dos últimos 30 dias (o resto vê-se em "histórico completo")
const NUM_COMPUTADORES = 15;   // computadores da biblioteca (numerados de 1 a 15)

function hojeMeioDia_() {
  const a = new Date();
  return new Date(a.getFullYear(), a.getMonth(), a.getDate(), 12);
}
// "aaaa-mm-dd" → data ao meio-dia (a diferença de fuso entre script e folha não muda o dia)
function isoParaData_(s) {
  const p = String(s || "").split("-").map(Number);
  if (p.length !== 3 || !p[0] || !p[1] || !p[2]) throw new Error("Data inválida.");
  return new Date(p[0], p[1] - 1, p[2], 12);
}
function dataIso_(v) {
  const d = paraData(v);
  return d ? Utilities.formatDate(d, TIMEZONE, "yyyy-MM-dd") : "";
}
function somaDias_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n, 12);
}
// 1 de setembro do ano letivo em curso: limite para registos com data passada.
function inicioAnoLetivo_() {
  const h = hojeMeioDia_();
  return new Date(h.getMonth() >= 8 ? h.getFullYear() : h.getFullYear() - 1, 8, 1, 12);
}
// Data de um registo: vazio = hoje; não pode ser depois de hoje nem antes do início do ano letivo.
function dataRegisto_(iso) {
  if (!iso) return hojeMeioDia_();
  const d = isoParaData_(iso), di = dataIso_(d);
  if (di > dataIso_(hojeMeioDia_())) throw new Error("A data não pode ser depois de hoje.");
  if (di < dataIso_(inicioAnoLetivo_())) throw new Error("A data não pode ser anterior ao início do ano letivo (" + Utilities.formatDate(inicioAnoLetivo_(), TIMEZONE, "dd/MM/yyyy") + ").");
  return d;
}
function novoId_(prefixo) {
  return prefixo + "-" + Utilities.formatDate(new Date(), TIMEZONE, "yyMMddHHmmss") + "-" + (100 + Math.floor(Math.random() * 900));
}
function texto_(v, max) {
  return String(v === null || v === undefined ? "" : v).replace(/\s+/g, " ").trim().slice(0, max || 300);
}
function comBloqueio_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error("A folha está ocupada. Tente de novo dentro de alguns segundos.");
  try { return fn(); } finally { lock.releaseLock(); }
}
function linhaDoId_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function leitores_(ss) {
  const v = lerFolha(ss, "Leitores"), m = {};
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    m[String(v[i][0])] = { id: String(v[i][0]), nome: texto_(v[i][1]), tipo: texto_(v[i][2]), turma: texto_(v[i][3]), num: v[i][4] === "" ? "" : String(v[i][4]) };
  }
  return m;
}
function requisicaoCompacta_(r) {
  // [id, turma, leitor, livro, dataReq, dataEntrega, prorrogado, dataLimite, dataDevolucao] (datas "aaaa-mm-dd")
  return [String(r[0]), texto_(r[1]), String(r[2]), texto_(r[3]), dataIso_(r[4]), dataIso_(r[5]), numProrrog_(r[6]), dataIso_(r[7]), dataIso_(r[8])];
}

// Tudo o que o balcão precisa para arrancar.
function getBalcaoDados_() {
  const ss = abrirFolha_();
  const leit = leitores_(ss);
  const alunos = [], professores = [];
  Object.keys(leit).forEach(function (k) {
    const l = leit[k];
    if (l.tipo === "Aluno") alunos.push([l.id, l.nome, l.turma, l.num]);
    else if (l.tipo === "Professor") professores.push([l.id, l.nome]);
  });
  professores.sort(function (a, b) { return a[1].localeCompare(b[1], "pt"); });

  const tv = lerFolha(ss, "Turmas"), turmas = [];
  for (let i = 1; i < tv.length; i++) if (tv[i][0] !== "") turmas.push(texto_(tv[i][0]));

  // Catálogo por título: [título, autor, n.º de exemplares, registos]. Vários registos do mesmo título somam-se.
  const lv = lerFolha(ss, "Livros"), cat = {}, livrosId = {};
  for (let i = 1; i < lv.length; i++) {
    const t = texto_(lv[i][1]);
    if (!t) continue;
    if (lv[i][0]) livrosId[String(lv[i][0])] = t;
    let n = Number(lv[i][5]);
    if (!(n >= 1)) n = 1;
    const k = t.toLowerCase();
    if (!cat[k]) cat[k] = [t, texto_(lv[i][2], 80), 0, []];
    cat[k][2] += n;
    if (lv[i][0] && cat[k][3].length < 6) cat[k][3].push(String(lv[i][0]));
  }
  const catalogo = Object.keys(cat).map(function (k) { return [cat[k][0], cat[k][1], cat[k][2], cat[k][3].join(", ")]; });
  catalogo.sort(function (a, b) { return a[0].localeCompare(b[0], "pt"); });

  const rv = lerFolha(ss, "Requisicoes"), requisicoes = [];
  // Para a app não ficar lenta: vêm as requisições em aberto e as devolvidas nos últimos DIAS_HISTORICO dias.
  const corte = dataIso_(somaDias_(hojeMeioDia_(), -DIAS_HISTORICO));
  for (let i = 1; i < rv.length; i++) {
    if (rv[i][0] === "") continue;
    const c = requisicaoCompacta_(rv[i]);
    if (livrosId[c[3]]) c[3] = livrosId[c[3]];   // requisições antigas guardadas pelo código do livro
    if (!c[8] || c[8] >= corte) requisicoes.push(c);
  }

  const tempos = [];
  const tpv = ss.getSheetByName("Tempos").getDataRange().getDisplayValues();
  for (let i = 1; i < tpv.length; i++) if (tpv[i][0]) tempos.push([tpv[i][0], tpv[i][2], tpv[i][3], tpv[i][4]]);

  const ev = lerFolha(ss, "Espacos"), espacos = [];
  for (let i = 1; i < ev.length; i++) if (ev[i][0]) espacos.push([String(ev[i][0]), texto_(ev[i][1])]);

  const hoje = hojeMeioDia_(), hojeIso = dataIso_(hoje);
  const pv = ss.getSheetByName("Presencas").getDataRange().getDisplayValues(), pvv = lerFolha(ss, "Presencas");
  const presencasHoje = [];
  for (let i = 1; i < pvv.length; i++) if (dataIso_(pvv[i][1]) === hojeIso) presencasHoje.push([String(pv[i][2]).slice(0, 5), texto_(pvv[i][3]), String(pvv[i][4]), texto_(pvv[i][5]), pv[i][6] || "", String(pvv[i][0]), String(pv[i][7] || "").slice(0, 5)]);

  const mat = materialDados_(ss, leit);
  return { hoje: hojeIso, turmas: turmas, alunos: alunos, professores: professores, catalogo: catalogo,
           material: mat.material, materialFora: mat.fora, requisicoes: requisicoes, tempos: tempos, espacos: espacos, atividades: ATIVIDADES,
           presencasHoje: presencasHoje, diasEmprestimo: DIAS_EMPRESTIMO, diasProrrogacao: DIAS_PRORROGACAO,
           maxProrrogacoes: MAX_PRORROGACOES, inicioAno: dataIso_(inicioAnoLetivo_()),
           numComputadores: NUM_COMPUTADORES, diasHistorico: DIAS_HISTORICO };
}

// --- Requisições -----------------------------------------------------------
function criarRequisicao_(o) {
  o = o || {};
  const ss = abrirFolha_();
  const leit = leitores_(ss)[String(o.leitor)];
  if (!leit || leit.tipo !== "Aluno") throw new Error("Aluno não encontrado.");
  const livro = texto_(o.livro, 200);
  if (!livro) throw new Error("Escreva ou escolha o título do livro.");
  const dReq = dataRegisto_(o.dataReq);
  const dEnt = somaDias_(dReq, DIAS_EMPRESTIMO);
  // Requisição antiga que já foi devolvida: a data de devolução vem junto.
  let dDev = "";
  if (o.dataDev) {
    dDev = dataRegisto_(o.dataDev);
    if (dataIso_(dDev) < dataIso_(dReq)) throw new Error("A devolução não pode ser antes da requisição.");
  }
  const linha = [novoId_("REQ"), leit.turma, leit.id, livro, dReq, dEnt, false, dEnt, dDev];
  return comBloqueio_(function () {
    ss.getSheetByName("Requisicoes").appendRow(linha);
    return requisicaoCompacta_(linha);
  });
}

function prorrogarRequisicao_(id) {
  return comBloqueio_(function () {
    const sh = abrirFolha_().getSheetByName("Requisicoes");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Requisição não encontrada. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 9).getValues()[0];
    if (paraData(r[8])) throw new Error("Este livro já foi devolvido.");
    const feitas = numProrrog_(r[6]);
    if (feitas >= MAX_PRORROGACOES) throw new Error("Esta requisição já teve " + MAX_PRORROGACOES + " prorrogações (é o máximo).");
    // Cada prorrogação soma 5 dias ao prazo que estiver em vigor.
    const base = (feitas > 0 && paraData(r[7])) ? paraData(r[7]) : paraData(r[5]);
    const lim = somaDias_(base, DIAS_PRORROGACAO);
    sh.getRange(n, 7, 1, 2).setValues([[feitas + 1, lim]]);
    r[6] = feitas + 1; r[7] = lim;
    return requisicaoCompacta_(r);
  });
}

function registarDevolucao_(id, dataIso) {
  return comBloqueio_(function () {
    const sh = abrirFolha_().getSheetByName("Requisicoes");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Requisição não encontrada. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 9).getValues()[0];
    if (paraData(r[8])) throw new Error("Este livro já estava devolvido.");
    const dDev = dataRegisto_(dataIso);
    if (dataIso_(dDev) < dataIso_(r[4])) throw new Error("A devolução não pode ser antes da requisição (" + Utilities.formatDate(paraData(r[4]), TIMEZONE, "dd/MM/yyyy") + ").");
    sh.getRange(n, 9).setValue(dDev);
    r[8] = dDev;
    return requisicaoCompacta_(r);
  });
}

// --- Atividade --------------------------------------------------------------
function registarPresenca_(o) {
  o = o || {};
  const ss = abrirFolha_();
  const leit = leitores_(ss)[String(o.leitor)];
  if (!leit || leit.tipo !== "Aluno") throw new Error("Aluno não encontrado.");
  if (ATIVIDADES.indexOf(o.atividade) < 0) throw new Error("Atividade inválida.");
  const dia = dataRegisto_(o.data), passado = dataIso_(dia) !== dataIso_(hojeMeioDia_());
  let pc = "";
  if (o.atividade === "Usar computadores") {
    pc = parseInt(o.computador, 10);
    if (!(pc >= 1 && pc <= NUM_COMPUTADORES)) {
      if (!passado || o.computador) throw new Error("Escolha o número do computador (1 a " + NUM_COMPUTADORES + ").");
      pc = "";   // entrada de um dia anterior: o número do computador é opcional
    }
  }
  // Entrada de um dia anterior: fica a hora que se escrever (ou em branco).
  const hora = passado ? (/^\d{1,2}:\d{2}$/.test(String(o.hora || "")) ? String(o.hora) : "") : Utilities.formatDate(new Date(), TIMEZONE, "HH:mm");
  return comBloqueio_(function () {
    const sh = ss.getSheetByName("Presencas");
    if (sh.getRange(1, 7).getValue() === "") sh.getRange(1, 7).setValue("Computador");
    if (sh.getRange(1, 8).getValue() === "") sh.getRange(1, 8).setValue("Hora_Saida");
    // Se o computador ainda estava atribuído a outro aluno hoje, fica libertado a esta hora.
    if (pc !== "" && !passado) {
      const v = sh.getDataRange().getValues(), hojeIso = dataIso_(hojeMeioDia_());
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][6]) === String(pc) && (v[i][7] === "" || v[i][7] === undefined) && dataIso_(v[i][1]) === hojeIso) sh.getRange(i + 1, 8).setValue(hora);
      }
    }
    const id = novoId_("PRE");
    sh.appendRow([id, dia, hora, leit.turma, leit.id, o.atividade, pc, ""]);
    return [hora, leit.turma, leit.id, o.atividade, pc === "" ? "" : String(pc), id, "", dataIso_(dia)];
  });
}

// --- Reserva do espaço ------------------------------------------------------
// Reservas de uma semana (segundaIso = segunda-feira, "aaaa-mm-dd").
function getReservasSemana_(segundaIso) {
  const ini = isoParaData_(segundaIso), fim = somaDias_(ini, 6);
  const iniIso = dataIso_(ini), fimIso = dataIso_(fim);
  const v = lerFolha(abrirFolha_(), "Reservas"), out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    const d = dataIso_(v[i][1]);
    if (d && d >= iniIso && d <= fimIso) out.push([String(v[i][0]), d, String(v[i][2]), String(v[i][3]), String(v[i][4]), texto_(v[i][5]), texto_(v[i][6], 500), String(v[i][7] === undefined ? "" : v[i][7])]);
  }
  return out;
}

function criarReserva_(o) {
  o = o || {};
  const ss = abrirFolha_();
  const prof = leitores_(ss)[String(o.professor)];
  if (!prof || prof.tipo !== "Professor") throw new Error("Escolha o professor.");
  const finalidade = texto_(o.finalidade, 500);
  if (!finalidade) throw new Error("Escreva o que vão fazer (tipo de atividade).");
  const data = isoParaData_(o.data);
  const dataIso = dataIso_(data);
  if (dataIso < dataIso_(inicioAnoLetivo_())) throw new Error("Não é possível registar antes do início do ano letivo.");
  const tempo = String(o.tempo || ""), espaco = String(o.espaco || "");
  const tpv = ss.getSheetByName("Tempos").getDataRange().getValues();
  if (!tpv.some(function (r) { return String(r[0]) === tempo; })) throw new Error("Tempo inválido.");
  const turma = texto_(o.turma, 20);
  let semanas = parseInt(o.semanas, 10);
  if (!(semanas >= 1)) semanas = 1;
  if (semanas > 40) semanas = 40;
  return comBloqueio_(function () {
    const sh = ss.getSheetByName("Reservas"), v = sh.getDataRange().getValues();
    if (sh.getRange(1, 8).getValue() === "") sh.getRange(1, 8).setValue("Serie");
    const ocupado = {};
    for (let i = 1; i < v.length; i++) {
      if (v[i][0] && String(v[i][2]) === tempo && String(v[i][3]) === espaco) ocupado[dataIso_(v[i][1])] = true;
    }
    if (ocupado[dataIso]) throw new Error("Esse tempo já foi reservado entretanto. Carregue em Atualizar.");
    const serie = semanas > 1 ? novoId_("SER") : "";
    const criadas = [], ocupadas = [];
    for (let k = 0; k < semanas; k++) {
      const d = somaDias_(data, 7 * k), iso = dataIso_(d);
      if (ocupado[iso]) { ocupadas.push(iso); continue; }
      const id = novoId_("RES") + (k ? "-" + k : "");
      sh.appendRow([id, d, tempo, espaco, prof.id, turma, finalidade, serie]);
      criadas.push([id, iso, tempo, espaco, prof.id, turma, finalidade, serie]);
    }
    return { criadas: criadas, ocupadas: ocupadas };
  });
}

function apagarReserva_(id, quem, seguintes) {
  quem = texto_(quem, 80);
  if (!quem) throw new Error("Escreva quem está a cancelar a reserva.");
  return comBloqueio_(function () {
    const ss = abrirFolha_(), sh = ss.getSheetByName("Reservas");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Reserva não encontrada. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 8).getValues()[0];
    const prof = leitores_(ss)[String(r[4])];
    const serie = String(r[7] || ""), desde = dataIso_(r[1]), apagadas = [];
    if (seguintes && serie) {
      const v = sh.getDataRange().getValues();
      for (let i = v.length - 1; i >= 1; i--) {
        if (String(v[i][7] || "") === serie && dataIso_(v[i][1]) >= desde) { apagadas.push(String(v[i][0])); sh.deleteRow(i + 1); }
      }
    } else {
      apagadas.push(String(r[0])); sh.deleteRow(n);
    }
    registo_(ss, apagadas.length > 1 ? "Reservas canceladas (" + apagadas.length + " semanas)" : "Reserva cancelada",
      desde + " · " + r[2] + " · " + (prof ? prof.nome : r[4]) + (r[5] ? " · " + r[5] : "") + " · cancelada por: " + quem);
    return apagadas;
  });
}

// --- Anular enganos e libertar computadores --------------------------------
// Fica tudo anotado na folha "Registo" (data e hora, ação, detalhe).
function registo_(ss, acao, detalhe) {
  let sh = ss.getSheetByName("Registo");
  if (!sh) { sh = ss.insertSheet("Registo"); sh.getRange(1, 1, 1, 3).setValues([["DataHora", "Acao", "Detalhe"]]); }
  sh.appendRow([Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm"), acao, detalhe]);
}

function anularPresenca_(id) {
  return comBloqueio_(function () {
    const ss = abrirFolha_(), sh = ss.getSheetByName("Presencas");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Entrada não encontrada. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 7).getValues()[0];
    if (dataIso_(r[1]) !== dataIso_(hojeMeioDia_())) throw new Error("Só é possível anular entradas de hoje.");
    const leit = leitores_(ss)[String(r[4])];
    sh.deleteRow(n);
    registo_(ss, "Entrada anulada", (leit ? leit.nome : r[4]) + " (" + r[3] + ") · " + r[5] + (r[6] !== "" ? " · PC " + r[6] : ""));
    return true;
  });
}

function libertarComputador_(id) {
  const hora = Utilities.formatDate(new Date(), TIMEZONE, "HH:mm");
  return comBloqueio_(function () {
    const sh = abrirFolha_().getSheetByName("Presencas");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Entrada não encontrada. Carregue em Atualizar.");
    if (sh.getRange(1, 8).getValue() === "") sh.getRange(1, 8).setValue("Hora_Saida");
    sh.getRange(n, 8).setValue(hora);
    return hora;
  });
}

function anularDevolucao_(id) {
  return comBloqueio_(function () {
    const ss = abrirFolha_(), sh = ss.getSheetByName("Requisicoes");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Requisição não encontrada. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 9).getValues()[0];
    if (!paraData(r[8])) throw new Error("Este livro não estava marcado como devolvido.");
    sh.getRange(n, 9).setValue("");
    const leit = leitores_(ss)[String(r[2])];
    registo_(ss, "Devolução anulada", (leit ? leit.nome : r[2]) + " (" + r[1] + ") · " + r[3]);
    r[8] = "";
    return requisicaoCompacta_(r);
  });
}

function anularRequisicao_(id) {
  return comBloqueio_(function () {
    const ss = abrirFolha_(), sh = ss.getSheetByName("Requisicoes");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Requisição não encontrada. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 9).getValues()[0];
    const hoje = Utilities.formatDate(new Date(), TIMEZONE, "yyMMdd");
    if (String(r[0]).indexOf("REQ-" + hoje) !== 0) throw new Error("Só é possível anular requisições registadas hoje. Para as outras, registe a devolução.");
    const leit = leitores_(ss)[String(r[2])];
    sh.deleteRow(n);
    registo_(ss, "Requisição anulada", (leit ? leit.nome : r[2]) + " (" + r[1] + ") · " + r[3]);
    return true;
  });
}

// --- Material e jogos (fones, tablets, cabos, jogos…) -------------------------
// Folha "Material": ID_Material | Num_Inventario (nome ou número) | Tipo | Estado
// Folha "Emprestimos_Material": ID | Turma | Leitor | Material | Saída | Devolução | Estado_Devolucao | Observações
function materialDados_(ss) {
  const mv = lerFolha(ss, "Material"), material = [];
  for (let i = 1; i < mv.length; i++) if (mv[i][0] !== "") material.push([String(mv[i][0]), texto_(mv[i][1], 80), texto_(mv[i][2], 30) || "Outro", texto_(mv[i][3], 20)]);
  const ev = lerFolha(ss, "Emprestimos_Material"), fora = [];
  for (let i = 1; i < ev.length; i++) {
    if (ev[i][0] === "" || paraData(ev[i][5])) continue;
    const d = paraData(ev[i][4]);
    fora.push([String(ev[i][0]), texto_(ev[i][1]), String(ev[i][2]), String(ev[i][3]), d ? Utilities.formatDate(d, TIMEZONE, "dd/MM HH:mm") : ""]);
  }
  return { material: material, fora: fora };
}

function emprestarMaterial_(o) {
  o = o || {};
  const ss = abrirFolha_();
  const leit = leitores_(ss)[String(o.leitor)];
  if (!leit || leit.tipo !== "Aluno") throw new Error("Aluno não encontrado.");
  return comBloqueio_(function () {
    const shM = ss.getSheetByName("Material"), n = linhaDoId_(shM, o.material);
    if (n < 0) throw new Error("Esse material não existe na folha «Material». Carregue em Atualizar.");
    const m = shM.getRange(n, 1, 1, 4).getValues()[0];
    if (/avariad|danificad/i.test(String(m[3]))) throw new Error("Esse material está marcado como avariado.");
    const shE = ss.getSheetByName("Emprestimos_Material"), ev = shE.getDataRange().getValues();
    for (let i = 1; i < ev.length; i++) {
      if (ev[i][0] !== "" && String(ev[i][3]) === String(m[0]) && !paraData(ev[i][5])) throw new Error("Esse material já está emprestado. Carregue em Atualizar.");
    }
    const id = novoId_("EMP"), agora = new Date();
    shE.appendRow([id, leit.turma, leit.id, m[0], agora, "", "", ""]);
    shM.getRange(n, 4).setValue("Emprestado");
    return [id, leit.turma, leit.id, String(m[0]), Utilities.formatDate(agora, TIMEZONE, "dd/MM HH:mm")];
  });
}

function devolverMaterial_(id, o) {
  o = o || {};
  const problema = !!o.problema, obs = texto_(o.obs, 300);
  if (problema && !obs) throw new Error("Escreva qual é o problema.");
  return comBloqueio_(function () {
    const ss = abrirFolha_(), shE = ss.getSheetByName("Emprestimos_Material");
    const n = linhaDoId_(shE, id);
    if (n < 0) throw new Error("Empréstimo não encontrado. Carregue em Atualizar.");
    const e = shE.getRange(n, 1, 1, 8).getValues()[0];
    if (paraData(e[5])) throw new Error("Este material já estava devolvido.");
    shE.getRange(n, 6, 1, 3).setValues([[new Date(), problema ? "Danificado" : "OK", obs]]);
    const shM = ss.getSheetByName("Material"), nm = linhaDoId_(shM, e[3]);
    const estado = problema ? "Avariado" : "Disponível";
    if (nm > 0) shM.getRange(nm, 4).setValue(estado);
    return [String(e[3]), estado];
  });
}

// --- Vários alunos de uma vez ----------------------------------------------
function registarPresencas_(o) {
  o = o || {};
  if (ATIVIDADES.indexOf(o.atividade) < 0) throw new Error("Atividade inválida.");
  if (o.atividade === "Usar computadores") throw new Error("Os computadores registam-se um aluno de cada vez, para escolher o número.");
  const ids = Array.isArray(o.leitores) ? o.leitores : [];
  if (!ids.length) throw new Error("Escolha pelo menos um aluno.");
  if (ids.length > 40) throw new Error("No máximo 40 alunos de cada vez.");
  const ss = abrirFolha_(), leit = leitores_(ss), vistos = {}, lista = [];
  ids.forEach(function (id) {
    const l = leit[String(id)];
    if (!l || l.tipo !== "Aluno") throw new Error("Aluno não encontrado. Carregue em Atualizar.");
    if (!vistos[l.id]) { vistos[l.id] = 1; lista.push(l); }
  });
  const dia = dataRegisto_(o.data), passado = dataIso_(dia) !== dataIso_(hojeMeioDia_());
  const hora = passado ? (/^\d{1,2}:\d{2}$/.test(String(o.hora || "")) ? String(o.hora) : "") : Utilities.formatDate(new Date(), TIMEZONE, "HH:mm");
  return comBloqueio_(function () {
    const sh = ss.getSheetByName("Presencas"), hoje = dia;
    const linhas = lista.map(function (l, k) { return [novoId_("PRE") + "-" + k, hoje, hora, l.turma, l.id, o.atividade, "", ""]; });
    sh.getRange(sh.getLastRow() + 1, 1, linhas.length, 8).setValues(linhas);
    return linhas.map(function (r) { return [hora, r[3], String(r[4]), r[5], "", r[0], "", dataIso_(dia)]; });
  });
}

// --- Histórico completo de um aluno ------------------------------------------
function historicoAluno_(id) {
  const rv = lerFolha(abrirFolha_(), "Requisicoes"), out = [];
  for (let i = 1; i < rv.length; i++) if (rv[i][0] !== "" && String(rv[i][2]) === String(id)) out.push(requisicaoCompacta_(rv[i]));
  return out;
}

// --- Gerir alunos (acrescentar, corrigir, saiu da escola) --------------------
function guardarAluno_(o) {
  o = o || {};
  const nome = texto_(o.nome, 120), turma = texto_(o.turma, 20);
  if (nome.length < 3) throw new Error("Escreva o nome do aluno.");
  let num = texto_(o.num, 4);
  if (num !== "" && !/^\d{1,2}$/.test(num)) throw new Error("O número na turma tem de ser um número (ou ficar vazio).");
  num = num === "" ? "" : Number(num);
  return comBloqueio_(function () {
    const ss = abrirFolha_(), sh = ss.getSheetByName("Leitores");
    const tv = lerFolha(ss, "Turmas");
    if (!tv.some(function (r, i) { return i > 0 && texto_(r[0]) === turma; })) throw new Error("Essa turma não existe na folha «Turmas».");
    if (o.novo) {
      const id = texto_(o.id, 10);
      if (!/^\d{3,8}$/.test(id)) throw new Error("Escreva o número de processo do aluno (só algarismos).");
      if (linhaDoId_(sh, id) > 0) throw new Error("Já existe um leitor com o número de processo " + id + ".");
      sh.appendRow([Number(id), nome, "Aluno", turma, num, ""]);
      registo_(ss, "Aluno acrescentado", nome + " (" + turma + ") · proc. " + id);
      return [id, nome, turma, num === "" ? "" : String(num)];
    }
    const n = linhaDoId_(sh, o.id);
    if (n < 0) throw new Error("Aluno não encontrado. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 5).getValues()[0];
    if (texto_(r[2]) !== "Aluno") throw new Error("Esse leitor não é um aluno.");
    sh.getRange(n, 2).setValue(nome);
    sh.getRange(n, 4, 1, 2).setValues([[turma, num]]);
    if (texto_(r[1]) !== nome || texto_(r[3]) !== turma) registo_(ss, "Aluno alterado", texto_(r[1]) + " (" + texto_(r[3]) + ") → " + nome + " (" + turma + ") · proc. " + r[0]);
    return [String(r[0]), nome, turma, num === "" ? "" : String(num)];
  });
}

// O aluno deixa de aparecer no balcão, mas não se apaga nada (fica com o tipo "Saiu").
function retirarAluno_(id) {
  return comBloqueio_(function () {
    const ss = abrirFolha_(), sh = ss.getSheetByName("Leitores");
    const n = linhaDoId_(sh, id);
    if (n < 0) throw new Error("Aluno não encontrado. Carregue em Atualizar.");
    const r = sh.getRange(n, 1, 1, 4).getValues()[0];
    if (texto_(r[2]) !== "Aluno") throw new Error("Esse leitor não é um aluno.");
    const rv = lerFolha(ss, "Requisicoes");
    for (let i = 1; i < rv.length; i++) {
      if (rv[i][0] !== "" && String(rv[i][2]) === String(id) && !paraData(rv[i][8])) throw new Error("Este aluno ainda tem livros por devolver. Registe primeiro as devoluções.");
    }
    sh.getRange(n, 3).setValue("Saiu");
    registo_(ss, "Aluno saiu da escola", texto_(r[1]) + " (" + texto_(r[3]) + ") · proc. " + r[0]);
    return true;
  });
}

// ===========================================================================
// CÓPIA DE SEGURANÇA E PROTEÇÃO DOS CABEÇALHOS
// Correr UMA VEZ no editor a função ativarProtecoes (botão Executar).
// ===========================================================================
// Faz uma cópia da folha de cálculo inteira no Drive (no máximo uma por dia).
function copiaSeguranca() {
  const p = PropertiesService.getScriptProperties();
  const hoje = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
  if (p.getProperty("COPIA_DIA") === hoje) return "Já foi feita uma cópia hoje.";
  abrirFolha_().copy("Biblioteca — cópia de segurança " + hoje);
  p.setProperty("COPIA_DIA", hoje);
  return "Cópia feita: Biblioteca — cópia de segurança " + hoje;
}

// Protege a linha dos títulos de cada folha (aviso antes de alterar) e liga a cópia semanal (sexta, 19h).
function ativarProtecoes() {
  const ss = abrirFolha_();
  ss.getSheets().forEach(function (sh) {
    const ja = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).some(function (x) { return x.getDescription() === "Cabeçalhos (não alterar)"; });
    if (!ja && sh.getLastRow() > 0) sh.getRange("1:1").protect().setDescription("Cabeçalhos (não alterar)").setWarningOnly(true);
  });
  const tem = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "copiaSeguranca"; });
  if (!tem) ScriptApp.newTrigger("copiaSeguranca").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(19).create();
  Logger.log("Cabeçalhos protegidos. Cópia semanal ativa (sexta-feira, 19h). " + copiaSeguranca());
}

// ===========================================================================
// CÓDIGO DE ACESSO
// O código está na folha "Config" (célula ao lado de "PIN"); para o mudar,
// basta escrever outro nessa célula. As páginas só recebem dados depois de
// o código ser confirmado aqui no servidor.
// ===========================================================================
function pinAtual_() {
  const ss = abrirFolha_();
  let sh = ss.getSheetByName("Config");
  if (!sh) {
    sh = ss.insertSheet("Config");
    sh.getRange(1, 1, 2, 2).setValues([["Chave", "Valor"], ["PIN", String(100000 + Math.floor(Math.random() * 900000))]]);
  }
  const v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim().toUpperCase() === "PIN") return String(v[i][1]).trim();
  throw new Error("Falta a linha PIN na folha «Config».");
}
function segredo_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty("SEGREDO");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty("SEGREDO", s); }
  return s;
}
function tokenDe_(pin) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, segredo_() + "|" + pin, Utilities.Charset.UTF_8));
}
function tokenValido_() {
  const cache = CacheService.getScriptCache();
  let t = cache.get("TOKEN");
  if (!t) { t = tokenDe_(pinAtual_()); cache.put("TOKEN", t, 120); }
  return t;
}
function exigirSessao_(token) {
  if (!token || String(token) !== tokenValido_()) throw new Error("SESSAO: é preciso introduzir o código de acesso.");
}

// Cria a folha "Config" com um código ao acaso (correr uma vez no editor, se ainda não existir).
function criarCodigoDeAcesso() {
  Logger.log("Código de acesso atual: " + pinAtual_() + " (folha Config)");
}

// Chamada pelas páginas quando alguém escreve o código. Devolve a chave da sessão.
function entrar(pin) {
  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get("FALHAS") || 0);
  if (falhas >= 10) throw new Error("Demasiadas tentativas erradas. Aguarde 10 minutos e tente de novo.");
  const certo = pinAtual_();
  if (certo.length < 4) throw new Error("O código na folha «Config» tem de ter pelo menos 4 caracteres.");
  if (String(pin === null || pin === undefined ? "" : pin).trim() !== certo) {
    cache.put("FALHAS", String(falhas + 1), 600);
    Utilities.sleep(800);
    throw new Error("Código errado.");
  }
  cache.remove("TOKEN");
  return tokenDe_(certo);
}

// Porta única do Balcão: confirma a sessão e só depois executa a ação pedida.
function balcao(token, acao, args) {
  exigirSessao_(token);
  const ACOES = {
    getBalcaoDados: getBalcaoDados_, criarRequisicao: criarRequisicao_, prorrogarRequisicao: prorrogarRequisicao_,
    registarDevolucao: registarDevolucao_, anularDevolucao: anularDevolucao_, anularRequisicao: anularRequisicao_,
    registarPresenca: registarPresenca_, anularPresenca: anularPresenca_, libertarComputador: libertarComputador_,
    getReservasSemana: getReservasSemana_, criarReserva: criarReserva_, apagarReserva: apagarReserva_,
    registarPresencas: registarPresencas_, historicoAluno: historicoAluno_, guardarAluno: guardarAluno_, retirarAluno: retirarAluno_,
    emprestarMaterial: emprestarMaterial_, devolverMaterial: devolverMaterial_
  };
  if (!Object.prototype.hasOwnProperty.call(ACOES, acao)) throw new Error("Ação desconhecida.");
  return ACOES[acao].apply(null, args || []);
}

// Porta única do painel de estatísticas.
function painel(token, periodo, dataInicioCustom, dataFimCustom) {
  exigirSessao_(token);
  return getEstatisticasPainel_(periodo, dataInicioCustom, dataFimCustom);
}
