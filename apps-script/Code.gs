// Script de agendamento da Yurt (Google Apps Script)
// Recebe pedidos do site, checa conflito na agenda, cria o evento PENDENTE,
// avisa o Roberto e envia a confirmação ao cliente quando o título virar [CONFIRMADO].
// Teste: o Roberto é o João (joao.valerio.j2013@gmail.com).

const EMAIL_ROBERTO = 'joao.valerio.j2013@gmail.com';
const TZ = 'America/Sao_Paulo';
const ENDERECO = 'R. Luiz Gama, 933, Loja 2 — Maringá/PR';
const DIAS_ANTECEDENCIA = 60;
const DIAS_RETENCAO = 30; // agendamentos com mais de 30 dias são apagados da agenda

// Disponibilidade: o site pede ?data=AAAA-MM-DD e recebe os horários ocupados
function doGet(e) {
  try {
    return disponibilidade(e);
  } catch (err) {
    avisarFalha('consulta de horários', err);
    return json({ erro: 'falha interna' });
  }
}

function disponibilidade(e) {
  const dia = (e && e.parameter && e.parameter.data) || '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dia)) return json({ erro: 'data inválida' });

  const ini = new Date(dia + 'T00:00:00-03:00');
  const fim = new Date(ini.getTime() + 24 * 60 * 60 * 1000);
  const busy = CalendarApp.getDefaultCalendar()
    .getEvents(ini, fim)
    .filter(ev => !ev.isAllDayEvent() && !ev.getTitle().startsWith('[CANCELADO]'))
    .map(ev => ({
      ini: Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm'),
      fim: Utilities.formatDate(ev.getEndTime(), TZ, 'HH:mm')
    }));
  return json({ busy: busy });
}

// Avisa o Roberto quando algo dá errado (no máximo 1 aviso por tipo a cada 6 horas)
function avisarFalha(onde, err, ignorarLimite) {
  try {
    const cache = CacheService.getScriptCache();
    if (!ignorarLimite && cache.get('falha_' + onde)) return;
    cache.put('falha_' + onde, '1', 6 * 60 * 60);
    const msg = String(err && err.message ? err.message : err);
    const quando = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    const html =
      '<div style="font-family:Arial,sans-serif;max-width:560px">' +
      '<div style="background:#C62828;color:#fff;padding:20px;font-size:26px;font-weight:bold;text-align:center">' +
      '🚨 ALERTA: O AGENDAMENTO DO SITE FALHOU 🚨</div>' +
      '<div style="border:4px solid #C62828;padding:20px;font-size:16px;line-height:1.5">' +
      '<p style="font-size:18px"><b>Pedidos novos podem NÃO estar chegando.</b> Avise quem cuida do site agora.</p>' +
      '<p><b>Onde:</b> ' + onde + '<br><b>Erro:</b> ' + msg + '<br><b>Quando:</b> ' + quando + '</p></div></div>';
    MailApp.sendEmail({
      to: EMAIL_ROBERTO,
      subject: '🚨🚨 URGENTE - ALERTA YURT: agendamento do site com ERRO (' + onde + ')',
      body: 'ALERTA URGENTE: o agendamento do site falhou. Onde: ' + onde + '. Erro: ' + msg + '. Quando: ' + quando + '. Pedidos novos podem não estar chegando.',
      htmlBody: html,
      name: '🚨 ALERTA YURT'
    });
  } catch (e2) { /* sem como avisar */ }
}

// Teste manual: rode esta função no editor para ver como o alerta chega
function testarAlerta() {
  avisarFalha('TESTE (pode ignorar)', new Error('Este é só um teste do alerta.'), true);
}

// Novo pedido vindo do site
function doPost(e) {
  try {
    return processarPedido(e);
  } catch (err) {
    avisarFalha('novo pedido', err);
    return json({ ok: false, erro: 'falha interna' });
  }
}

function processarPedido(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, erro: 'dados inválidos' });
  }

  const nome = String(d.nome || '').trim();
  const fone = String(d.whatsapp || '').replace(/\D/g, '');
  const email = String(d.email || '').trim();
  const servicos = String(d.servicos || '').trim();
  const inicio = new Date(d.inicio_iso);
  const fim = new Date(d.fim_iso);
  const agora = new Date();
  const limite = new Date(agora.getTime() + DIAS_ANTECEDENCIA * 24 * 60 * 60 * 1000);

  if (nome.length < 2 || nome.length > 80) return json({ ok: false, erro: 'nome' });
  if (fone.length < 10 || fone.length > 11) return json({ ok: false, erro: 'whatsapp' });
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ ok: false, erro: 'email' });
  if (!servicos || servicos.length > 300) return json({ ok: false, erro: 'serviços' });
  if (isNaN(inicio) || isNaN(fim) || fim <= inicio) return json({ ok: false, erro: 'horário' });
  if (inicio < agora || inicio > limite) return json({ ok: false, erro: 'data' });

  // Limite: um pedido por WhatsApp a cada 60 segundos
  const cache = CacheService.getScriptCache();
  if (cache.get('rl_' + fone)) return json({ ok: false, erro: 'limite' });
  cache.put('rl_' + fone, '1', 60);

  // Conflito: horário já ocupado na agenda
  const agenda = CalendarApp.getDefaultCalendar();
  const conflitos = agenda.getEvents(inicio, fim).filter(ev =>
    !ev.isAllDayEvent() &&
    !ev.getTitle().startsWith('[CANCELADO]') &&
    ev.getStartTime() < fim && ev.getEndTime() > inicio
  );
  if (conflitos.length) {
    MailApp.sendEmail(
      EMAIL_ROBERTO,
      'Pedido em horário ocupado: ' + nome,
      'Um cliente tentou agendar ' + d.data + ' às ' + d.horario + ' (' + servicos + '), mas o horário já está ocupado.\n' +
      'Nome: ' + nome + '\nWhatsApp: ' + fone + '\n\nEntre em contato para remarcar.'
    );
    return json({ ok: false, erro: 'ocupado' });
  }

  const ev = agenda.createEvent(
    '[PENDENTE] ' + servicos + ' - ' + nome,
    inicio, fim,
    { description: 'WhatsApp: ' + fone + '\nValor: ' + d.valor_total + '\nObservação: ' + (d.observacao || '—') }
  );
  ev.setTag('cliente_nome', nome);
  ev.setTag('cliente_email', email);
  ev.setTag('servicos', servicos);
  ev.setTag('data', d.data || '');
  ev.setTag('horario', d.horario || '');
  ev.setTag('notificado', '0');

  MailApp.sendEmail(
    EMAIL_ROBERTO,
    'Agendamento para o Roberto confirmar: ' + nome,
    'Cliente: ' + nome + '\nWhatsApp: ' + fone + '\nServiços: ' + servicos +
    '\nData: ' + (d.data || '') + ' às ' + (d.horario || '') + '\nValor: ' + d.valor_total +
    '\n\nPara confirmar, abra o evento na agenda do Roberto e troque [PENDENTE] por [CONFIRMADO] no título.'
  );
  return json({ ok: true });
}

// Verificador (acionador de tempo, a cada minuto): envia a confirmação ao cliente
function checarConfirmacoes() {
  try {
    verificarConfirmacoes();
  } catch (err) {
    avisarFalha('confirmações', err);
    throw err;
  }
}

function verificarConfirmacoes() {
  const cal = CalendarApp.getDefaultCalendar();
  const agora = new Date();
  const ate = new Date(agora.getTime() + DIAS_ANTECEDENCIA * 24 * 60 * 60 * 1000);

  cal.getEvents(agora, ate).forEach(function (ev) {
    const titulo = ev.getTitle();
    if (!titulo.startsWith('[CONFIRMADO]') || ev.getTag('notificado') === '1') return;

    const cliente = ev.getTag('cliente_email');
    const nome = ev.getTag('cliente_nome');
    const quando = ev.getTag('data') + ' às ' + ev.getTag('horario');

    if (cliente) {
      MailApp.sendEmail(
        cliente,
        'Seu horário está confirmado na Yurt',
        'Olá, ' + nome + '!\n\nSeu horário está confirmado:\n' + ev.getTag('servicos') +
        '\n' + quando + '\n\nEndereço: ' + ENDERECO + '\nPagamento via Pix. Até lá!'
      );
    }
    MailApp.sendEmail(
      EMAIL_ROBERTO,
      'Confirmação enviada ao cliente: ' + nome,
      'A confirmação foi enviada para ' + (cliente || 'cliente sem e-mail') + ' (' + quando + ').'
    );
    ev.setTag('notificado', '1');
  });
}

// Limpeza (acionador diário): apaga da agenda os agendamentos do site com mais de DIAS_RETENCAO dias.
// Só mexe em eventos criados pelo site (título [PENDENTE], [CONFIRMADO] ou [CANCELADO]).
// Rode limparAntigos(true) no editor para só contar, sem apagar.
function limparAntigos(simular) {
  try {
    const cal = CalendarApp.getDefaultCalendar();
    const corte = new Date(new Date().getTime() - DIAS_RETENCAO * 24 * 60 * 60 * 1000);
    const desde = new Date(corte.getTime() - 365 * 24 * 60 * 60 * 1000);
    const antigos = cal.getEvents(desde, corte).filter(function (ev) {
      return /^\[(PENDENTE|CONFIRMADO|CANCELADO)\]/.test(ev.getTitle()) && ev.getEndTime() < corte;
    });
    if (simular === true) {
      Logger.log('Simulação: ' + antigos.length + ' agendamentos seriam apagados.');
      return antigos.length;
    }
    antigos.forEach(function (ev) { ev.deleteEvent(); });
    Logger.log('Apagados ' + antigos.length + ' agendamentos com mais de ' + DIAS_RETENCAO + ' dias.');
    return antigos.length;
  } catch (err) {
    avisarFalha('limpeza de agendamentos antigos', err);
    throw err;
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
