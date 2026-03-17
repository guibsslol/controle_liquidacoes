// ==========================================
// CONFIGURAÇÃO DO FIREBASE (NUVEM)
// ==========================================
const firebaseConfig = {
  apiKey: 'AIzaSyAIQrfYY0QvtyN7qj61uLhq6Xyb4Eyn3ZA',
  authDomain: 'controliqui-smebji.firebaseapp.com',
  databaseURL: 'https://controliqui-smebji-default-rtdb.firebaseio.com',
  projectId: 'controliqui-smebji',
  storageBucket: 'controliqui-smebji.firebasestorage.app',
  messagingSenderId: '659644181097',
  appId: '1:659644181097:web:82b55c5eba921bc06e10f8',
};

// Inicializa a Nuvem
firebase.initializeApp(firebaseConfig);
const database = firebase.database();

// ==========================================
// VARIÁVEIS GLOBAIS DO SISTEMA
// ==========================================
let dadosAbas = { 'Assunto Geral': { Janeiro: [] } };
let abaAtiva = 'Assunto Geral';
let mesAtivo = 'Janeiro';

let abaArrastada = null;
let subAbaArrastada = null;
let linhaEmEdicao = null;
let colunaSort = 'id';
let ordemSort = 'asc';
let IDsSelecionados = new Set();
let dadosRelatorioGeral = [];

let chartStatus = null;
let chartEmpresas = null;

// ==========================================
// TEMA E MEMÓRIA DE NAVEGAÇÃO
// ==========================================
function toggleTema() {
  const isDark = document.body.getAttribute('data-theme') === 'dark';
  if (isDark) {
    document.body.removeAttribute('data-theme');
    localStorage.setItem('tema', 'light');
  } else {
    document.body.setAttribute('data-theme', 'dark');
    localStorage.setItem('tema', 'dark');
  }
  if (document.getElementById('modal-dashboard').style.display === 'flex') atualizarGraficos();
}

// ==========================================
// MÁSCARA E TECLAS
// ==========================================
function mascaraProcesso(e) {
  let input = e.target;
  if (
    (input.value === 'BJI' || input.value === 'BJI-') &&
    e.inputType === 'deleteContentBackward'
  ) {
    input.value = '';
    return;
  }
  let val = input.value
    .toUpperCase()
    .replace('BJI-', '')
    .replace(/[^A-Z0-9]/g, '');
  if (val.length === 0 && input.value.length > 0) {
    input.value = 'BJI-';
    return;
  }
  let formatted = 'BJI-';
  if (val.length > 0) formatted += val.substring(0, 6);
  if (val.length > 6) formatted += '/' + val.substring(6, 12);
  if (val.length > 12) formatted += '/' + val.substring(12, 16);
  input.value = formatted;
}

document.addEventListener('keydown', function (e) {
  if (linhaEmEdicao !== null) {
    if (e.key === 'Escape') {
      e.preventDefault();
      cancelarEdicaoInline();
      return;
    }
    if (e.key === 'Enter') {
      if (e.target.classList.contains('input-inline')) {
        e.preventDefault();
        salvarEdicaoInline(linhaEmEdicao);
        return;
      }
    }
  }
  if (e.key === 'Enter') {
    const noFormulario = e.target.closest('.form-row') && !e.target.closest('.filtros-container');
    if (noFormulario) {
      e.preventDefault();
      adicionarRegistro();
    }
  }
});

// ==========================================
// 1. INICIALIZAÇÃO E COMUNICAÇÃO COM O FIREBASE
// ==========================================
window.onload = () => {
  if (localStorage.getItem('tema') === 'dark') document.body.setAttribute('data-theme', 'dark');

  // Ouve as mudanças no Firebase em Tempo Real
  const dbRef = database.ref('sistema');
  dbRef.on('value', (snapshot) => {
    const data = snapshot.val();

    if (data && data.abas) {
      dadosAbas = data.abas;

      let abaMemoria = localStorage.getItem('ultimaAba');
      abaAtiva =
        abaMemoria && dadosAbas[abaMemoria] ? abaMemoria : data.ativa || Object.keys(dadosAbas)[0];

      let mesMemoria = localStorage.getItem('ultimoMes');
      mesAtivo =
        mesMemoria && dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesMemoria]
          ? mesMemoria
          : Object.keys(dadosAbas[abaAtiva])[0] || 'Geral';
    } else {
      // Primeira vez abrindo o sistema ou banco totalmente vazio
      dadosAbas = { 'Assunto Geral': { Janeiro: [] } };
      abaAtiva = 'Assunto Geral';
      mesAtivo = 'Janeiro';
    }

    renderizarAbas();
    renderizarSubAbas();
    renderizarTabela();
    atualizarAutocompletarAba();

    const statusEl = document.getElementById('status-conexao');
    statusEl.innerHTML = '<i class="fa-solid fa-cloud"></i> Online (Tempo Real)';
    statusEl.style.color = '#27ae60';
    statusEl.style.borderColor = '#27ae60';
  });
};

function salvarArquivoAutomaticamente() {
  // Envia os dados atualizados para a Nuvem
  database
    .ref('sistema')
    .set({ abas: dadosAbas, ativa: abaAtiva })
    .catch((error) => {
      console.error('Erro ao salvar no Firebase:', error);
      Swal.fire(
        'Erro de Conexão',
        'Não foi possível salvar na nuvem. Verifique sua internet.',
        'error',
      );
    });
}

function fazerBackupSeguranca() {
  if (Object.keys(dadosAbas).length === 0) return;
  const dados = JSON.stringify({ abas: dadosAbas, ativa: abaAtiva }, null, 2);
  const blob = new Blob([dados], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Backup_Controle_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
  Swal.fire({
    icon: 'success',
    title: 'Backup Feito!',
    text: 'Guarde este arquivo num local seguro.',
    toast: true,
    position: 'top-end',
    timer: 3000,
  });
}

// NOVO: Restaurar os dados de um Backup Local direto para a Nuvem
function restaurarBackup(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const importado = JSON.parse(e.target.result);
      if (importado && importado.abas) {
        dadosAbas = importado.abas;
        abaAtiva = importado.ativa || Object.keys(dadosAbas)[0];
        salvarArquivoAutomaticamente();
        Swal.fire('Restaurado!', 'Seu backup foi enviado para a nuvem com sucesso!', 'success');
      } else {
        Swal.fire('Erro', 'O arquivo JSON não possui a estrutura correta.', 'error');
      }
    } catch (err) {
      Swal.fire('Erro', 'Arquivo JSON inválido ou corrompido.', 'error');
    }
    document.getElementById('input-restaurar-json').value = '';
  };
  reader.readAsText(file);
}

// ==========================================
// 2. SISTEMA DE ABAS E SUB-ABAS
// ==========================================
function renderizarAbas() {
  const listaAbas = document.getElementById('lista-abas');
  listaAbas.innerHTML = '';
  Object.keys(dadosAbas).forEach((nomeAba) => {
    const divAba = document.createElement('div');
    divAba.className = `aba ${nomeAba === abaAtiva ? 'ativa' : ''}`;
    divAba.setAttribute('draggable', true);
    divAba.innerHTML = `
        <span class="titulo-aba">${nomeAba}</span>
        <div class="aba-acoes">
            <button class="btn-aba-acao edit" onclick="event.stopPropagation(); editarNomeAba('${nomeAba}')" title="Renomear"><i class="fa-solid fa-pen"></i></button>
            <button class="btn-aba-acao" onclick="event.stopPropagation(); excluirAba('${nomeAba}')" title="Excluir"><i class="fa-solid fa-trash"></i></button>
        </div>`;

    divAba.onclick = () => {
      abaAtiva = nomeAba;
      mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
      linhaEmEdicao = null;
      IDsSelecionados.clear();
      localStorage.setItem('ultimaAba', abaAtiva);
      localStorage.setItem('ultimoMes', mesAtivo);
      salvarArquivoAutomaticamente();
      renderizarAbas();
      renderizarSubAbas();
      renderizarTabela();
    };

    divAba.addEventListener('dragstart', function (e) {
      abaArrastada = nomeAba;
      e.dataTransfer.effectAllowed = 'move';
      setTimeout(() => (this.style.opacity = '0.4'), 0);
    });
    divAba.addEventListener('dragend', function () {
      this.style.opacity = '1';
      document.querySelectorAll('.aba').forEach((a) => a.classList.remove('drag-over'));
      abaArrastada = null;
    });
    divAba.addEventListener('dragover', function (e) {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      return false;
    });
    divAba.addEventListener('dragenter', function (e) {
      if (nomeAba !== abaArrastada) this.classList.add('drag-over');
    });
    divAba.addEventListener('dragleave', function () {
      this.classList.remove('drag-over');
    });
    divAba.addEventListener('drop', function (e) {
      e.stopPropagation();
      this.classList.remove('drag-over');
      if (abaArrastada !== nomeAba) {
        const chaves = Object.keys(dadosAbas);
        chaves.splice(chaves.indexOf(abaArrastada), 1);
        chaves.splice(chaves.indexOf(nomeAba), 0, abaArrastada);
        const novo = {};
        chaves.forEach((c) => (novo[c] = dadosAbas[c]));
        dadosAbas = novo;
        salvarArquivoAutomaticamente();
        renderizarAbas();
      }
      return false;
    });
    listaAbas.appendChild(divAba);
  });
}

function renderizarSubAbas() {
  const listaMeses = document.getElementById('lista-sub-abas');
  listaMeses.innerHTML = '';
  if (!dadosAbas[abaAtiva]) return;

  Object.keys(dadosAbas[abaAtiva]).forEach((nomeMes) => {
    const divMes = document.createElement('div');
    divMes.className = `sub-aba ${nomeMes === mesAtivo ? 'ativa' : ''}`;
    divMes.setAttribute('draggable', true);
    divMes.innerHTML = `
            <span>${nomeMes}</span>
            <div class="aba-acoes" style="display: ${nomeMes === mesAtivo ? 'flex' : 'none'}">
                <button class="btn-aba-acao edit" onclick="event.stopPropagation(); editarNomeMes('${nomeMes}')" title="Renomear"><i class="fa-solid fa-pen"></i></button>
                <button class="btn-aba-acao" onclick="event.stopPropagation(); excluirMes('${nomeMes}')" title="Excluir"><i class="fa-solid fa-trash"></i></button>
            </div>`;

    divMes.onclick = () => {
      mesAtivo = nomeMes;
      linhaEmEdicao = null;
      IDsSelecionados.clear();
      localStorage.setItem('ultimoMes', mesAtivo);
      renderizarSubAbas();
      renderizarTabela();
    };

    divMes.addEventListener('dragstart', function (e) {
      subAbaArrastada = nomeMes;
      e.dataTransfer.effectAllowed = 'move';
      setTimeout(() => (this.style.opacity = '0.4'), 0);
    });
    divMes.addEventListener('dragend', function () {
      this.style.opacity = '1';
      document.querySelectorAll('.sub-aba').forEach((a) => a.classList.remove('drag-over'));
      subAbaArrastada = null;
    });
    divMes.addEventListener('dragover', function (e) {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      return false;
    });
    divMes.addEventListener('dragenter', function (e) {
      if (nomeMes !== subAbaArrastada) this.classList.add('drag-over');
    });
    divMes.addEventListener('dragleave', function () {
      this.classList.remove('drag-over');
    });
    divMes.addEventListener('drop', function (e) {
      e.stopPropagation();
      this.classList.remove('drag-over');
      if (subAbaArrastada !== nomeMes) reordenarSubAbas(subAbaArrastada, nomeMes);
      return false;
    });
    listaMeses.appendChild(divMes);
  });
}

function reordenarSubAbas(mesOrigem, mesDestino) {
  const chaves = Object.keys(dadosAbas[abaAtiva]);
  chaves.splice(chaves.indexOf(mesOrigem), 1);
  chaves.splice(chaves.indexOf(mesDestino), 0, mesOrigem);
  const novoObjetoMeses = {};
  chaves.forEach((chave) => {
    novoObjetoMeses[chave] = dadosAbas[abaAtiva][chave];
  });
  dadosAbas[abaAtiva] = novoObjetoMeses;
  salvarArquivoAutomaticamente();
  renderizarSubAbas();
}

async function duplicarMes() {
  const mesesDisponiveis = Object.keys(dadosAbas[abaAtiva]);
  if (mesesDisponiveis.length === 0) {
    Swal.fire('Erro', 'Não há meses para copiar.', 'error');
    return;
  }
  let optionsHtml = '';
  mesesDisponiveis.forEach((m) => {
    optionsHtml += `<option value="${m}" ${m === mesAtivo ? 'selected' : ''}>${m}</option>`;
  });

  const { value: formValues } = await Swal.fire({
    title: 'Copiar Mês',
    html: `
            <div style="text-align: left; font-size: 14px;">
                <label style="font-weight: bold; color: var(--text-main);">1. Qual mês deseja copiar?</label>
                <select id="swal-origem" class="swal2-select" style="width: 100%; margin: 5px 0 15px 0; padding: 5px;">${optionsHtml}</select>
                <label style="font-weight: bold; color: var(--text-main);">2. Nome do Novo Mês:</label>
                <input id="swal-novo-mes" class="swal2-input" placeholder="Ex: Março-26" style="width: 100%; margin: 5px 0 15px 0; box-sizing: border-box;">
                <label style="font-weight: bold; color: var(--text-main);">3. Quais colunas deseja importar?</label>
                <div style="margin-top: 5px; display: grid; grid-template-columns: 1fr 1fr; gap: 8px; background: var(--bg-header); padding: 10px; border-radius: 5px; border: 1px solid var(--border-color);">
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-proc" checked> Processo</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-emp" checked> Empresa</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-elem" checked> Elemento</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-empenho" checked> Empenho</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-liq"> Liquidação</label>
                </div>
            </div>`,
    focusConfirm: false,
    showCancelButton: true,
    confirmButtonText: 'Criar e Importar',
    cancelButtonText: 'Cancelar',
    preConfirm: () => {
      const origem = document.getElementById('swal-origem').value;
      const novoMes = document.getElementById('swal-novo-mes').value.trim();
      if (!novoMes) {
        Swal.showValidationMessage('Digite o nome do novo mês!');
        return false;
      }
      if (dadosAbas[abaAtiva][novoMes]) {
        Swal.showValidationMessage('Já existe um mês com este nome nesta aba!');
        return false;
      }
      return {
        origem,
        novoMes,
        importProc: document.getElementById('chk-proc').checked,
        importEmp: document.getElementById('chk-emp').checked,
        importElem: document.getElementById('chk-elem').checked,
        importEmpenho: document.getElementById('chk-empenho').checked,
        importLiq: document.getElementById('chk-liq').checked,
      };
    },
  });

  if (formValues) {
    const { origem, novoMes, importProc, importEmp, importElem, importEmpenho, importLiq } =
      formValues;
    dadosAbas[abaAtiva][novoMes] = [];
    dadosAbas[abaAtiva][origem].forEach((reg, index) => {
      dadosAbas[abaAtiva][novoMes].push({
        id: Date.now() + index,
        processo: importProc ? reg.processo : '',
        empresa: importEmp ? reg.empresa : '',
        elemento: importElem ? reg.elemento : '',
        empenho: importEmpenho ? reg.empenho : '',
        liquidacao: importLiq ? reg.liquidacao : '',
        status: 'Aguardando Pagamento',
        op: '',
      });
    });
    mesAtivo = novoMes;
    localStorage.setItem('ultimoMes', mesAtivo);
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
    Swal.fire(
      'Pronto!',
      `O mês "${novoMes}" foi criado com ${dadosAbas[abaAtiva][novoMes].length} processos!`,
      'success',
    );
  }
}

async function editarNomeAba(nomeAtual) {
  const { value: novoNome } = await Swal.fire({
    title: 'Renomear Assunto',
    input: 'text',
    inputValue: nomeAtual,
    showCancelButton: true,
  });
  if (novoNome && novoNome.trim() !== '' && novoNome !== nomeAtual) {
    if (!dadosAbas[novoNome]) {
      dadosAbas[novoNome] = dadosAbas[nomeAtual];
      delete dadosAbas[nomeAtual];
      if (abaAtiva === nomeAtual) {
        abaAtiva = novoNome;
        localStorage.setItem('ultimaAba', abaAtiva);
      }
      salvarArquivoAutomaticamente();
      renderizarAbas();
    } else {
      Swal.fire('Erro', 'Já existe um assunto com este nome.', 'error');
    }
  }
}

function excluirAba(nomeAba) {
  if (Object.keys(dadosAbas).length > 1) {
    Swal.fire({
      title: `Excluir "${nomeAba}"?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((result) => {
      if (result.isConfirmed) {
        delete dadosAbas[nomeAba];
        abaAtiva = Object.keys(dadosAbas)[0];
        mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
        localStorage.setItem('ultimaAba', abaAtiva);
        localStorage.setItem('ultimoMes', mesAtivo);
        salvarArquivoAutomaticamente();
        renderizarAbas();
        renderizarSubAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'Precisa de ter pelo menos uma aba de assunto.', 'info');
  }
}

async function criarNovaAba() {
  const { value: nome } = await Swal.fire({
    title: 'Novo Assunto',
    input: 'text',
    showCancelButton: true,
  });
  if (nome && nome.trim() !== '') {
    if (!dadosAbas[nome]) {
      dadosAbas[nome] = { Geral: [] };
      abaAtiva = nome;
      mesAtivo = 'Geral';
      localStorage.setItem('ultimaAba', abaAtiva);
      localStorage.setItem('ultimoMes', mesAtivo);
      salvarArquivoAutomaticamente();
      renderizarAbas();
      renderizarSubAbas();
      renderizarTabela();
    }
  }
}

async function editarNomeMes(nomeAtual) {
  const { value: novoNome } = await Swal.fire({
    title: 'Renomear Mês',
    input: 'text',
    inputValue: nomeAtual,
    showCancelButton: true,
  });
  if (novoNome && novoNome.trim() !== '' && novoNome !== nomeAtual) {
    if (!dadosAbas[abaAtiva][novoNome]) {
      dadosAbas[abaAtiva][novoNome] = dadosAbas[abaAtiva][nomeAtual];
      delete dadosAbas[abaAtiva][nomeAtual];
      if (mesAtivo === nomeAtual) {
        mesAtivo = novoNome;
        localStorage.setItem('ultimoMes', mesAtivo);
      }
      salvarArquivoAutomaticamente();
      renderizarSubAbas();
    } else {
      Swal.fire('Erro', 'Já existe um mês com este nome.', 'error');
    }
  }
}

function excluirMes(nomeMes) {
  if (Object.keys(dadosAbas[abaAtiva]).length > 1) {
    Swal.fire({
      title: `Excluir o mês "${nomeMes}"?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((result) => {
      if (result.isConfirmed) {
        delete dadosAbas[abaAtiva][nomeMes];
        mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
        localStorage.setItem('ultimoMes', mesAtivo);
        salvarArquivoAutomaticamente();
        renderizarSubAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'O assunto precisa de ter pelo menos um mês.', 'info');
  }
}

async function criarNovoMes() {
  const { value: nome } = await Swal.fire({
    title: 'Novo Mês',
    input: 'text',
    placeholder: 'Ex: Fevereiro',
    showCancelButton: true,
  });
  if (nome && nome.trim() !== '') {
    if (!dadosAbas[abaAtiva][nome]) {
      dadosAbas[abaAtiva][nome] = [];
      mesAtivo = nome;
      localStorage.setItem('ultimoMes', mesAtivo);
      salvarArquivoAutomaticamente();
      renderizarSubAbas();
      renderizarTabela();
    }
  }
}

async function agruparAbas() {
  const abasDisponiveis = Object.keys(dadosAbas);
  if (abasDisponiveis.length < 2) {
    Swal.fire('Atenção', 'Precisa de pelo menos 2 assuntos para os agrupar.', 'info');
    return;
  }
  let htmlCheckboxes =
    '<div style="text-align: left; max-height: 200px; overflow-y: auto; margin-top: 15px; padding: 10px; border: 1px solid var(--border-color); border-radius: 5px; background: var(--bg-header);">';
  abasDisponiveis.forEach((aba) => {
    htmlCheckboxes += `<label style="display: block; margin-bottom: 8px; cursor: pointer; font-size: 14px; color: var(--text-main);"><input type="checkbox" class="swal-aba-checkbox" value="${aba}" style="margin-right: 8px;"> ${aba}</label>`;
  });
  htmlCheckboxes += '</div>';

  const { value: formValues } = await Swal.fire({
    title: 'Agrupar Assuntos',
    html:
      `<div style="font-size: 14px; text-align: left; margin-bottom: 10px; color: var(--text-main);">Selecione os assuntos que deseja fundir e dê um nome para a nova Aba Principal:</div><input id="swal-input-novo-nome" class="swal2-input" placeholder="Nome do Novo Assunto" style="margin-top: 0;">` +
      htmlCheckboxes,
    focusConfirm: false,
    showCancelButton: true,
    confirmButtonText: 'Agrupar Agora',
    cancelButtonText: 'Cancelar',
    preConfirm: () => {
      const selecionados = Array.from(document.querySelectorAll('.swal-aba-checkbox:checked')).map(
        (cb) => cb.value,
      );
      const novoNome = document.getElementById('swal-input-novo-nome').value.trim();
      if (selecionados.length === 0) {
        Swal.showValidationMessage('Selecione pelo menos 1 assunto para agrupar!');
        return false;
      }
      if (!novoNome) {
        Swal.showValidationMessage('Digite o nome!');
        return false;
      }
      return { selecionados, novoNome };
    },
  });

  if (formValues) {
    const { selecionados, novoNome } = formValues;
    if (!dadosAbas[novoNome]) dadosAbas[novoNome] = {};
    selecionados.forEach((abaAntiga) => {
      Object.keys(dadosAbas[abaAntiga]).forEach((mesAntigo) => {
        let novoNomeMes = abaAntiga
          .replace('LIQ. ', '')
          .replace('LIQ.', '')
          .replace('Liq. ', '')
          .replace('LIQ ', '')
          .trim();
        if (dadosAbas[novoNome][novoNomeMes])
          dadosAbas[novoNome][novoNomeMes] = dadosAbas[novoNome][novoNomeMes].concat(
            dadosAbas[abaAntiga][mesAntigo],
          );
        else dadosAbas[novoNome][novoNomeMes] = dadosAbas[abaAntiga][mesAntigo];
      });
      if (abaAntiga !== novoNome) delete dadosAbas[abaAntiga];
    });
    abaAtiva = novoNome;
    mesAtivo = Object.keys(dadosAbas[novoNome])[0];
    localStorage.setItem('ultimaAba', abaAtiva);
    localStorage.setItem('ultimoMes', mesAtivo);
    salvarArquivoAutomaticamente();
    renderizarAbas();
    renderizarSubAbas();
    renderizarTabela();
    Swal.fire('Sucesso!', 'Assuntos agrupados!', 'success');
  }
}

async function moverProcessosLote() {
  const { value: formValues } = await Swal.fire({
    title: 'Transferência em Lote',
    html: `
        <div style="text-align: left; font-size: 14px;">
            <p style="margin-bottom:5px; font-weight:bold; color: #e67e22;">1. Procurar na ABA ATUAL processos onde:</p>
            <select id="swal-move-tipo" class="swal2-select" style="width:100%; padding:5px; margin-bottom:10px;">
                <option value="empresa">Empresa for igual a</option><option value="processo">Processo for igual a</option><option value="empenho">Empenho for igual a</option><option value="elemento">Elemento for igual a</option>
            </select>
            <input id="swal-move-valor" class="swal2-input" placeholder="Digite o valor exato..." style="width:100%; margin: 5px 0 15px 0; box-sizing: border-box;">
            <p style="margin-bottom:5px; font-weight:bold; color: #27ae60;">2. Mover estes processos para:</p>
            <select id="swal-move-aba" class="swal2-select" style="width:100%; padding:5px; margin-bottom:10px;">${Object.keys(
              dadosAbas,
            )
              .map((a) => `<option value="${a}">${a}</option>`)
              .join('')}</select>
            <input id="swal-move-mes" class="swal2-input" placeholder="Nome da Sub-Aba/Mês de Destino" style="width:100%; margin: 5px 0 15px 0; box-sizing: border-box;">
        </div>`,
    focusConfirm: false,
    showCancelButton: true,
    confirmButtonText: 'Mover Processos',
    cancelButtonText: 'Cancelar',
    preConfirm: () => {
      const tipo = document.getElementById('swal-move-tipo').value;
      const valor = document.getElementById('swal-move-valor').value.trim();
      const abaDestino = document.getElementById('swal-move-aba').value;
      const mesDestino = document.getElementById('swal-move-mes').value.trim();
      if (!valor || !mesDestino) {
        Swal.showValidationMessage('Preencha o valor e o mês!');
        return false;
      }
      return { tipo, valor, abaDestino, mesDestino };
    },
  });

  if (formValues) {
    const { tipo, valor, abaDestino, mesDestino } = formValues;
    let processosMovidos = [];
    let processosRestantes = [];
    dadosAbas[abaAtiva][mesAtivo].forEach((reg) => {
      if (reg[tipo] && reg[tipo].toString().toLowerCase() === valor.toLowerCase())
        processosMovidos.push(reg);
      else processosRestantes.push(reg);
    });
    if (processosMovidos.length === 0) {
      Swal.fire('Nenhum encontrado', `Nenhum processo coincide com a busca.`, 'info');
      return;
    }
    if (!dadosAbas[abaDestino][mesDestino]) dadosAbas[abaDestino][mesDestino] = [];
    dadosAbas[abaDestino][mesDestino] = dadosAbas[abaDestino][mesDestino].concat(processosMovidos);
    dadosAbas[abaAtiva][mesAtivo] = processosRestantes;
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
    Swal.fire('Sucesso!', `${processosMovidos.length} processos movidos!`, 'success');
  }
}

function atualizarSelecao(checkbox, id) {
  if (checkbox.checked) IDsSelecionados.add(id);
  else IDsSelecionados.delete(id);
  verificarBarraLote();
}
function toggleSelecionarTodos(checkboxCentral) {
  document.querySelectorAll('.chk-linha').forEach((chk) => {
    chk.checked = checkboxCentral.checked;
    const id = parseInt(chk.value);
    if (checkboxCentral.checked) IDsSelecionados.add(id);
    else IDsSelecionados.delete(id);
  });
  verificarBarraLote();
}
function verificarBarraLote() {
  const barra = document.getElementById('acoes-lote');
  const contador = document.getElementById('contador-selecionados');
  if (IDsSelecionados.size > 0) {
    barra.style.display = 'flex';
    contador.innerText = `${IDsSelecionados.size} processo(s) selecionado(s)`;
  } else {
    barra.style.display = 'none';
    const chk = document.getElementById('chk-todos');
    if (chk) chk.checked = false;
  }
}

function excluirSelecionados() {
  Swal.fire({
    title: `Excluir ${IDsSelecionados.size} processos?`,
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#d33',
    confirmButtonText: 'Sim, Excluir',
  }).then((result) => {
    if (result.isConfirmed) {
      dadosAbas[abaAtiva][mesAtivo] = dadosAbas[abaAtiva][mesAtivo].filter(
        (r) => !IDsSelecionados.has(r.id),
      );
      IDsSelecionados.clear();
      verificarBarraLote();
      salvarArquivoAutomaticamente();
      renderizarTabela();
      Swal.fire({
        icon: 'success',
        title: 'Excluídos!',
        toast: true,
        position: 'top-end',
        showConfirmButton: false,
        timer: 1500,
      });
    }
  });
}

async function moverSelecionados() {
  let abasOptions = Object.keys(dadosAbas)
    .map((a) => `<option value="${a}">${a}</option>`)
    .join('');
  const { value: formValues } = await Swal.fire({
    title: `Mover ${IDsSelecionados.size} Processos`,
    html: `
        <div style="text-align: left; font-size: 14px;">
            <label style="font-weight:bold; color: #2980b9;">Aba de Destino:</label>
            <select id="swal-move-sel-aba" class="swal2-select" style="width:100%; padding:5px; margin-bottom:10px;" onchange="window.atualizarMesesDestinoSel()"><option value="">Selecione a Aba...</option>${abasOptions}</select>
            <label style="font-weight:bold; color: #2980b9;">Mês de Destino:</label>
            <select id="swal-move-sel-mes" class="swal2-select" style="width:100%; padding:5px; margin-bottom:10px;"><option value="">Selecione primeiro a Aba acima</option></select>
            <div style="font-size:12px; margin-top:5px; color:var(--text-muted);">Ou digite para criar um Mês Novo no destino:</div>
            <input id="swal-move-sel-mes-novo" class="swal2-input" placeholder="Ex: Maio-26" style="width:100%; margin: 5px 0 0 0; box-sizing: border-box;">
        </div>`,
    didOpen: () => {
      window.atualizarMesesDestinoSel = () => {
        const aba = document.getElementById('swal-move-sel-aba').value;
        const selMes = document.getElementById('swal-move-sel-mes');
        selMes.innerHTML = '<option value="">Selecione o Mês...</option>';
        if (aba && dadosAbas[aba]) {
          Object.keys(dadosAbas[aba]).forEach((m) => {
            selMes.innerHTML += `<option value="${m}">${m}</option>`;
          });
        }
      };
    },
    focusConfirm: false,
    showCancelButton: true,
    confirmButtonText: 'Mover Processos',
    cancelButtonText: 'Cancelar',
    preConfirm: () => {
      const abaDestino = document.getElementById('swal-move-sel-aba').value;
      let mesDestino = document.getElementById('swal-move-sel-mes').value;
      const mesNovo = document.getElementById('swal-move-sel-mes-novo').value.trim();
      if (!abaDestino) {
        Swal.showValidationMessage('Selecione a aba de destino!');
        return false;
      }
      if (!mesDestino && !mesNovo) {
        Swal.showValidationMessage('Selecione ou digite o mês!');
        return false;
      }
      if (mesNovo) mesDestino = mesNovo;
      return { abaDestino, mesDestino };
    },
  });

  if (formValues) {
    const { abaDestino, mesDestino } = formValues;
    let processosMovidos = [];
    let processosRestantes = [];
    dadosAbas[abaAtiva][mesAtivo].forEach((reg) => {
      if (IDsSelecionados.has(reg.id)) processosMovidos.push(reg);
      else processosRestantes.push(reg);
    });
    if (!dadosAbas[abaDestino][mesDestino]) dadosAbas[abaDestino][mesDestino] = [];
    dadosAbas[abaDestino][mesDestino] = dadosAbas[abaDestino][mesDestino].concat(processosMovidos);
    dadosAbas[abaAtiva][mesAtivo] = processosRestantes;
    IDsSelecionados.clear();
    verificarBarraLote();
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
    Swal.fire('Sucesso!', `${processosMovidos.length} processos movidos!`, 'success');
  }
}

// ==========================================
// RELATÓRIO GLOBAL DE PENDÊNCIAS
// ==========================================
function abrirRelatorioPendencias() {
  dadosRelatorioGeral = [];
  for (let aba in dadosAbas) {
    for (let mes in dadosAbas[aba]) {
      dadosAbas[aba][mes].forEach((reg) => {
        if (reg.status === 'Aguardando Pagamento' || !reg.op) {
          dadosRelatorioGeral.push({ ...reg, abaLocal: aba, mesLocal: mes });
        }
      });
    }
  }

  const preencher = (id, field) => {
    const select = document.getElementById(id);
    const uniq = [...new Set(dadosRelatorioGeral.map((r) => r[field]).filter(Boolean))].sort();
    select.innerHTML =
      '<option value="">Todos</option>' +
      uniq.map((v) => `<option value="${v}">${v}</option>`).join('');
  };
  preencher('rel-filtro-processo', 'processo');
  preencher('rel-filtro-empresa', 'empresa');
  preencher('rel-filtro-elemento', 'elemento');
  document.getElementById('rel-filtro-status').value = '';
  document.getElementById('modal-relatorio').style.display = 'flex';
  filtrarRelatorio();
}

function fecharRelatorio() {
  document.getElementById('modal-relatorio').style.display = 'none';
}

function filtrarRelatorio() {
  const fProc = document.getElementById('rel-filtro-processo').value;
  const fEmp = document.getElementById('rel-filtro-empresa').value;
  const fElem = document.getElementById('rel-filtro-elemento').value;
  const fStatus = document.getElementById('rel-filtro-status').value;

  let filtrados = dadosRelatorioGeral.filter((reg) => {
    let matchStatus = true;
    if (fStatus === 'Aguardando') matchStatus = reg.status === 'Aguardando Pagamento';
    if (fStatus === 'SemOP') matchStatus = !reg.op && reg.status === 'Pago';
    return (
      (!fProc || reg.processo === fProc) &&
      (!fEmp || reg.empresa === fEmp) &&
      (!fElem || reg.elemento === fElem) &&
      matchStatus
    );
  });

  const tbody = document.getElementById('tabela-relatorio-corpo');
  tbody.innerHTML = '';

  if (filtrados.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 20px; color: #27ae60; font-weight:bold;"><i class="fa-solid fa-check-circle"></i> Parabéns! Nenhuma pendência encontrada.</td></tr>`;
    return;
  }

  filtrados.forEach((reg) => {
    let opVisivel = reg.op
      ? reg.op
      : '<span style="background:#fff3cd; color:#d35400; padding:2px 6px; border-radius:3px; font-weight:bold;">S/ OP</span>';
    let corStatus =
      reg.status === 'Aguardando Pagamento'
        ? 'color:#c0392b; font-weight:bold;'
        : 'color:#27ae60; font-weight:bold;';
    tbody.innerHTML += `
        <tr>
            <td>${reg.processo}</td><td>${reg.empresa}</td><td>${reg.elemento}</td><td>${reg.empenho}</td><td>${reg.liquidacao}</td>
            <td style="${corStatus}">${reg.status}</td><td>${opVisivel}</td>
            <td style="font-weight:bold; color:var(--text-muted);">${reg.abaLocal} <i class="fa-solid fa-angle-right"></i> ${reg.mesLocal}</td>
            <td class="coluna-acao"><button onclick="irParaProcesso('${reg.abaLocal}', '${reg.mesLocal}', ${reg.id})" style="padding:5px 10px; background:#3498db; color:white; border:none; border-radius:4px; cursor:pointer;" title="Ir para o processo"><i class="fa-solid fa-arrow-right"></i></button></td>
        </tr>`;
  });
}

window.irParaProcesso = function (aba, mes, id) {
  abaAtiva = aba;
  mesAtivo = mes;
  localStorage.setItem('ultimaAba', abaAtiva);
  localStorage.setItem('ultimoMes', mesAtivo);
  fecharRelatorio();
  IDsSelecionados.clear();
  salvarArquivoAutomaticamente();
  renderizarAbas();
  renderizarSubAbas();
  renderizarTabela();
  setTimeout(() => {
    const btnEdit = document.querySelector(`button[onclick="ativarEdicaoInline(${id})"]`);
    if (btnEdit) {
      const tr = btnEdit.closest('tr');
      tr.style.transition = 'background-color 1s';
      tr.style.backgroundColor = '#fff3cd';
      tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
      setTimeout(() => (tr.style.backgroundColor = 'transparent'), 2000);
    }
  }, 300);
};

// ==========================================
// DASHBOARD (GRÁFICOS)
// ==========================================
function abrirDashboard() {
  document.getElementById('modal-dashboard').style.display = 'flex';
  atualizarGraficos();
}
function fecharDashboard() {
  document.getElementById('modal-dashboard').style.display = 'none';
}

function atualizarGraficos() {
  const registros =
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo] ? dadosAbas[abaAtiva][mesAtivo] : [];
  const isDark = document.body.getAttribute('data-theme') === 'dark';
  const textColor = isDark ? '#e0e0e0' : '#333';

  let aguardando = 0,
    pagoComOp = 0,
    pagoSemOp = 0;
  let empresas = {};

  registros.forEach((r) => {
    if (r.status === 'Aguardando Pagamento') aguardando++;
    else if (r.status === 'Pago' && !r.op) pagoSemOp++;
    else pagoComOp++;
    if (r.empresa) empresas[r.empresa] = (empresas[r.empresa] || 0) + 1;
  });

  const empresasArray = Object.entries(empresas)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  const topEmpresasLabels = empresasArray.map((e) =>
    e[0].length > 20 ? e[0].substring(0, 20) + '...' : e[0],
  );
  const topEmpresasData = empresasArray.map((e) => e[1]);

  const ctxStatus = document.getElementById('graficoStatus').getContext('2d');
  if (chartStatus) chartStatus.destroy();
  chartStatus = new Chart(ctxStatus, {
    type: 'doughnut',
    data: {
      labels: ['Pagas com OP', 'Pagas Sem OP', 'Aguardando'],
      datasets: [
        {
          data: [pagoComOp, pagoSemOp, aguardando],
          backgroundColor: ['#2ecc71', '#f39c12', '#e74c3c'],
          borderWidth: isDark ? 2 : 1,
          borderColor: isDark ? '#1e1e1e' : '#fff',
        },
      ],
    },
    options: { plugins: { legend: { labels: { color: textColor } } } },
  });

  const ctxEmpresas = document.getElementById('graficoEmpresas').getContext('2d');
  if (chartEmpresas) chartEmpresas.destroy();
  chartEmpresas = new Chart(ctxEmpresas, {
    type: 'bar',
    data: {
      labels: topEmpresasLabels,
      datasets: [
        {
          label: 'Nº de Processos',
          data: topEmpresasData,
          backgroundColor: '#3498db',
          borderRadius: 5,
        },
      ],
    },
    options: {
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: textColor }, grid: { display: false } },
        y: { ticks: { color: textColor, stepSize: 1 }, grid: { color: isDark ? '#333' : '#eee' } },
      },
    },
  });
}

// ==========================================
// AUTOCOMPLETAR
// ==========================================
function atualizarAutocompletarAba() {
  let procs = new Set(),
    emps = new Set(),
    elems = new Set(),
    emps2 = new Set();
  if (dadosAbas[abaAtiva]) {
    Object.values(dadosAbas[abaAtiva]).forEach((mes) => {
      mes.forEach((reg) => {
        if (reg.processo) procs.add(reg.processo);
        if (reg.empresa) emps.add(reg.empresa);
        if (reg.elemento) elems.add(reg.elemento);
        if (reg.empenho) emps2.add(reg.empenho);
      });
    });
  }
  const preencherListas = (idSelect, items) => {
    const datalist = document.getElementById(idSelect);
    if (datalist) {
      datalist.innerHTML = '';
      Array.from(items)
        .sort()
        .forEach((item) => (datalist.innerHTML += `<option value="${item}">`));
    }
  };
  preencherListas('lista-processos', procs);
  preencherListas('lista-empresas', emps);
  preencherListas('lista-elementos', elems);
  preencherListas('lista-empenhos', emps2);
}

function ordenarTabela(coluna) {
  if (colunaSort === coluna) ordemSort = ordemSort === 'asc' ? 'desc' : 'asc';
  else {
    colunaSort = coluna;
    ordemSort = 'asc';
  }
  renderizarTabela();
}

// ==========================================
// TABELA E EDIÇÃO
// ==========================================
function adicionarRegistro() {
  const processo = document.getElementById('processo').value.trim();
  const empresa = document.getElementById('empresa').value.trim();
  const elemento = document.getElementById('elemento').value.trim();
  const empenho = document.getElementById('empenho').value.trim();
  const liquidacao = document.getElementById('liquidacao').value.trim();
  const status = document.getElementById('status_pagamento').value;
  const op = document.getElementById('op').value.trim();

  if (!empresa) {
    Swal.fire('Obrigatório', 'A Empresa é obrigatória!', 'warning');
    return;
  }

  let duplicado = null;
  for (let aba in dadosAbas) {
    for (let mes in dadosAbas[aba]) {
      let achou = dadosAbas[aba][mes].find(
        (r) =>
          r.processo === processo &&
          r.empresa === empresa &&
          r.elemento === elemento &&
          r.empenho === empenho &&
          r.liquidacao === liquidacao,
      );
      if (achou) {
        duplicado = { aba, mes };
        break;
      }
    }
    if (duplicado) break;
  }

  if (duplicado) {
    Swal.fire({
      icon: 'error',
      title: 'Processo Duplicado!',
      html: `Este lançamento já existe em:<br><br><b>Aba:</b> ${duplicado.aba}<br><b>Mês:</b> ${duplicado.mes}`,
    });
    return;
  }

  const novoReg = { id: Date.now(), processo, empresa, elemento, empenho, liquidacao, status, op };
  if (!dadosAbas[abaAtiva][mesAtivo]) dadosAbas[abaAtiva][mesAtivo] = [];
  dadosAbas[abaAtiva][mesAtivo].push(novoReg);

  Swal.fire({
    icon: 'success',
    title: 'Adicionado!',
    toast: true,
    position: 'top-end',
    showConfirmButton: false,
    timer: 1500,
  });
  salvarArquivoAutomaticamente();
  renderizarTabela();
  limparInputs();
}

function limparInputs() {
  document.querySelectorAll('.form-row input').forEach((i) => {
    if (!i.id.includes('filtro')) i.value = '';
  });
  document.getElementById('status_pagamento').value = 'Aguardando Pagamento';
  document.getElementById('processo').focus();
}

function ativarEdicaoInline(id) {
  linhaEmEdicao = id;
  renderizarTabela();
}

window.ativarEdicaoEFocus = function (id, campo) {
  ativarEdicaoInline(id);
  setTimeout(() => {
    const el = document.getElementById(`edit-${campo}-${id}`);
    if (el) el.focus();
  }, 50);
};

window.atualizarStatusDireto = function (id, novoStatus) {
  const index = dadosAbas[abaAtiva][mesAtivo].findIndex((r) => r.id === id);
  if (index !== -1) {
    dadosAbas[abaAtiva][mesAtivo][index].status = novoStatus;
    salvarArquivoAutomaticamente();
    renderizarTabela();
  }
};

function cancelarEdicaoInline() {
  linhaEmEdicao = null;
  renderizarTabela();
}

function salvarEdicaoInline(id) {
  const index = dadosAbas[abaAtiva][mesAtivo].findIndex((r) => r.id === id);
  if (index !== -1) {
    dadosAbas[abaAtiva][mesAtivo][index].processo = document
      .getElementById(`edit-processo-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][mesAtivo][index].empresa = document
      .getElementById(`edit-empresa-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][mesAtivo][index].elemento = document
      .getElementById(`edit-elemento-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][mesAtivo][index].empenho = document
      .getElementById(`edit-empenho-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][mesAtivo][index].liquidacao = document
      .getElementById(`edit-liquidacao-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][mesAtivo][index].status = document.getElementById(
      `edit-status-${id}`,
    ).value;
    dadosAbas[abaAtiva][mesAtivo][index].op = document.getElementById(`edit-op-${id}`).value.trim();
    salvarArquivoAutomaticamente();
  }
  linhaEmEdicao = null;
  renderizarTabela();
  Swal.fire({
    icon: 'success',
    title: 'Atualizado!',
    toast: true,
    position: 'top-end',
    showConfirmButton: false,
    timer: 1500,
  });
}

function atualizarFiltrosDinâmicosDaTela(registrosDaTela) {
  const preencherSelect = (idSelect, valores) => {
    const select = document.getElementById(idSelect);
    const valorAtual = select.value;
    select.innerHTML = `<option value="">Todos</option>`;
    Array.from(valores)
      .sort()
      .forEach((v) => (select.innerHTML += `<option value="${v}">${v}</option>`));
    select.value = valorAtual;
  };
  const registrosBrutos =
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo] ? dadosAbas[abaAtiva][mesAtivo] : [];
  preencherSelect(
    'filtro-processo',
    new Set(registrosBrutos.map((r) => r.processo).filter(Boolean)),
  );
  preencherSelect('filtro-empresa', new Set(registrosBrutos.map((r) => r.empresa).filter(Boolean)));
  preencherSelect(
    'filtro-elemento',
    new Set(registrosBrutos.map((r) => r.elemento).filter(Boolean)),
  );
  preencherSelect('filtro-empenho', new Set(registrosBrutos.map((r) => r.empenho).filter(Boolean)));
}

function renderizarTabela() {
  const tbody = document.getElementById('tabela-corpo');
  tbody.innerHTML = '';
  let registros =
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo] ? [...dadosAbas[abaAtiva][mesAtivo]] : [];

  document.querySelectorAll('.th-sortable i').forEach((icon) => {
    icon.className = 'fa-solid fa-sort';
    icon.parentElement.classList.remove('sorted-asc', 'sorted-desc');
  });
  const currentIcon = document.getElementById(`sort-icon-${colunaSort}`);
  if (currentIcon) {
    currentIcon.className = ordemSort === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';
    currentIcon.parentElement.classList.add(`sorted-${ordemSort}`);
  }

  registros.sort((a, b) => {
    let valA = (a[colunaSort] || '').toString().toLowerCase();
    let valB = (b[colunaSort] || '').toString().toLowerCase();
    if (valA < valB) return ordemSort === 'asc' ? -1 : 1;
    if (valA > valB) return ordemSort === 'asc' ? 1 : -1;
    return 0;
  });

  const fProcesso = document.getElementById('filtro-processo').value;
  const fEmpresa = document.getElementById('filtro-empresa').value;
  const fElemento = document.getElementById('filtro-elemento').value;
  const fEmpenho = document.getElementById('filtro-empenho').value;
  const fStatus = document.getElementById('filtro-status').value;

  registros = registros.filter((reg) => {
    return (
      (fProcesso === '' || reg.processo === fProcesso) &&
      (fEmpresa === '' || reg.empresa === fEmpresa) &&
      (fElemento === '' || reg.elemento === fElemento) &&
      (fEmpenho === '' || reg.empenho === fEmpenho) &&
      (fStatus === '' || reg.status === fStatus)
    );
  });

  const cardTotal = document.getElementById('resumo-total');
  if (cardTotal) {
    const total = registros.length;
    const pagasFull = registros.filter((r) => r.status === 'Pago').length;

    cardTotal.innerText = total;
    document.getElementById('resumo-pendente').innerText = registros.filter(
      (r) => r.status === 'Aguardando Pagamento',
    ).length;
    document.getElementById('resumo-pago').innerText = registros.filter(
      (r) => r.status === 'Pago' && r.op !== '',
    ).length;
    document.getElementById('resumo-sem-op').innerText = registros.filter(
      (r) => r.status === 'Pago' && r.op === '',
    ).length;

    const percent = total > 0 ? (pagasFull / total) * 100 : 0;
    document.getElementById('barra-progresso').style.width = `${percent}%`;
  }

  const chkTodos = document.getElementById('chk-todos');
  if (chkTodos) {
    if (registros.length > 0 && registros.every((r) => IDsSelecionados.has(r.id)))
      chkTodos.checked = true;
    else chkTodos.checked = false;
  }

  const chkAgrupar = document.getElementById('chk-agrupar-empresa');
  const modoAgrupado = chkAgrupar && chkAgrupar.checked;

  if (modoAgrupado) {
    const grupos = {};
    registros.forEach((r) => {
      const emp = r.empresa || 'Sem Empresa';
      if (!grupos[emp]) grupos[emp] = [];
      grupos[emp].push(r);
    });
    Object.keys(grupos)
      .sort()
      .forEach((empresa) => {
        const trGroup = document.createElement('tr');
        trGroup.className = 'row-group-header';
        trGroup.innerHTML = `<td colspan="9" style="padding: 10px;"><i class="fa-solid fa-building" style="color: var(--text-muted); margin-right:5px;"></i> ${empresa} <span style="float:right; font-size:12px; font-weight:normal; color:var(--text-muted); padding-top:2px;">${grupos[empresa].length} processo(s)</span></td>`;
        tbody.appendChild(trGroup);
        grupos[empresa].forEach((reg) => {
          tbody.appendChild(criarElementoTR(reg));
        });
      });
  } else {
    registros.forEach((reg) => {
      tbody.appendChild(criarElementoTR(reg));
    });
  }
  atualizarFiltrosDinâmicosDaTela(registros);
}

function criarElementoTR(reg) {
  const tr = document.createElement('tr');
  const isChecked = IDsSelecionados.has(reg.id) ? 'checked' : '';

  if (linhaEmEdicao === reg.id) {
    tr.innerHTML = `
        <td></td>
        <td><input type="text" id="edit-processo-${reg.id}" class="input-inline" value="${reg.processo}" oninput="mascaraProcesso(event)"></td>
        <td><input type="text" id="edit-empresa-${reg.id}" class="input-inline" value="${reg.empresa}"></td>
        <td><input type="text" id="edit-elemento-${reg.id}" class="input-inline" value="${reg.elemento}"></td>
        <td><input type="text" id="edit-empenho-${reg.id}" class="input-inline" value="${reg.empenho}"></td>
        <td><input type="text" id="edit-liquidacao-${reg.id}" class="input-inline" value="${reg.liquidacao}"></td>
        <td><select id="edit-status-${reg.id}" class="input-inline" style="padding: 5px;"><option value="Aguardando Pagamento" ${reg.status === 'Aguardando Pagamento' ? 'selected' : ''}>Aguardando Pagamento</option><option value="Pago" ${reg.status === 'Pago' ? 'selected' : ''}>Pago</option></select></td>
        <td><input type="text" id="edit-op-${reg.id}" class="input-inline" value="${reg.op}"></td>
        <td class="coluna-acao"><button class="btn-edit" onclick="salvarEdicaoInline(${reg.id})" title="Confirmar" style="background-color: #27ae60;"><i class="fa-solid fa-check"></i></button><button class="btn-delete" onclick="cancelarEdicaoInline()" title="Cancelar" style="background-color: #95a5a6;"><i class="fa-solid fa-xmark"></i></button></td>
      `;
    tr.style.backgroundColor = 'var(--bg-header)';
    tr.style.boxShadow = 'inset 0 0 5px rgba(52, 152, 219, 0.3)';
  } else {
    let classEmpenhoLiq = '',
      classOP = '';
    if (reg.empenho !== '' && reg.liquidacao !== '') classEmpenhoLiq = 'bg-bege';
    if (reg.liquidacao !== '') {
      if (reg.status === 'Aguardando Pagamento' || reg.op === '') classOP = 'bg-vermelho';
      else if (reg.status === 'Pago' && reg.op !== '') classOP = 'bg-verde';
    }

    tr.innerHTML = `
        <td style="text-align: center;"><input type="checkbox" class="chk-linha" value="${reg.id}" ${isChecked} onchange="atualizarSelecao(this, ${reg.id})" style="cursor:pointer;"></td>
        <td ondblclick="ativarEdicaoEFocus(${reg.id}, 'processo')" style="cursor:pointer;" title="Duplo clique para editar">${reg.processo}</td>
        <td ondblclick="ativarEdicaoEFocus(${reg.id}, 'empresa')" style="cursor:pointer;" title="Duplo clique para editar">${reg.empresa}</td>
        <td ondblclick="ativarEdicaoEFocus(${reg.id}, 'elemento')" style="cursor:pointer;" title="Duplo clique para editar">${reg.elemento}</td>
        <td class="${classEmpenhoLiq}" ondblclick="ativarEdicaoEFocus(${reg.id}, 'empenho')" style="cursor:pointer;" title="Duplo clique para editar">${reg.empenho}</td>
        <td class="${classEmpenhoLiq}" ondblclick="ativarEdicaoEFocus(${reg.id}, 'liquidacao')" style="cursor:pointer;" title="Duplo clique para editar">${reg.liquidacao}</td>
        <td class="${classOP}"><select class="status-dropdown" onchange="atualizarStatusDireto(${reg.id}, this.value)"><option value="Aguardando Pagamento" ${reg.status === 'Aguardando Pagamento' ? 'selected' : ''}>Aguardando Pagamento</option><option value="Pago" ${reg.status === 'Pago' ? 'selected' : ''}>Pago</option></select></td>
        <td class="${classOP}" ondblclick="ativarEdicaoEFocus(${reg.id}, 'op')" style="cursor:pointer;" title="Duplo clique para editar">${reg.op}</td>
        <td class="coluna-acao"><button class="btn-edit" onclick="ativarEdicaoInline(${reg.id})" title="Editar na Linha"><i class="fa-solid fa-pen"></i></button><button class="btn-delete" onclick="apagarLinha(${reg.id})" title="Excluir"><i class="fa-solid fa-trash"></i></button></td>
      `;
    if (isChecked) tr.style.backgroundColor = 'var(--bg-lote)';
  }
  return tr;
}

function apagarLinha(id) {
  Swal.fire({
    title: 'Excluir?',
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#d33',
    confirmButtonText: 'Excluir',
  }).then((result) => {
    if (result.isConfirmed) {
      dadosAbas[abaAtiva][mesAtivo] = dadosAbas[abaAtiva][mesAtivo].filter((r) => r.id !== id);
      IDsSelecionados.delete(id);
      verificarBarraLote();
      salvarArquivoAutomaticamente();
      renderizarTabela();
    }
  });
}

function exportarParaExcel() {
  let tabela = document.getElementById('tabela-processos');
  let nomeArquivo = `${abaAtiva}_${mesAtivo}`.replace(/[^a-z0-9]/gi, '_');
  let workbook = XLSX.utils.table_to_book(tabela, { sheet: 'Plan1' });
  XLSX.writeFile(workbook, `Controle_${nomeArquivo}.xlsx`);
}

function processarImportacaoExcel(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    Swal.fire({
      title: 'Importar Planilha Completa?',
      text: `Encontramos ${workbook.SheetNames.length} páginas no Excel. Serão importadas como NOVOS ASSUNTOS.`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Sim, importar tudo!',
      cancelButtonText: 'Cancelar',
    }).then((result) => {
      if (result.isConfirmed) {
        let totalImportados = 0;
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          if (!dadosAbas[sheetName]) dadosAbas[sheetName] = { Geral: [] };
          else if (!dadosAbas[sheetName]['Geral']) dadosAbas[sheetName]['Geral'] = [];
          for (let i = 1; i < json.length; i++) {
            const row = json[i];
            if (!row || row.length === 0) continue;
            const proc = row[1] ? String(row[1]).trim() : '';
            const emp = row[2] ? String(row[2]).trim() : '';
            if (!emp) continue;
            const elem = row[3] ? String(row[3]).trim() : '';
            const empen = row[4] ? String(row[4]).trim() : '';
            const liq = row[5] ? String(row[5]).trim() : '';
            const numOp = row[6] ? String(row[6]).trim() : '';
            const statusPag = numOp !== '' ? 'Pago' : 'Aguardando Pagamento';
            let procFormatado = proc;
            if (procFormatado && !procFormatado.toUpperCase().startsWith('BJI-'))
              procFormatado = 'BJI-' + procFormatado;
            const novoReg = {
              id: Date.now() + Math.floor(Math.random() * 100000),
              processo: procFormatado,
              empresa: emp,
              elemento: elem,
              empenho: empen,
              liquidacao: liq,
              status: statusPag,
              op: numOp,
            };
            dadosAbas[sheetName]['Geral'].push(novoReg);
            totalImportados++;
          }
        });
        if (workbook.SheetNames.length > 0) {
          abaAtiva = workbook.SheetNames[0];
          mesAtivo = 'Geral';
          localStorage.setItem('ultimaAba', abaAtiva);
          localStorage.setItem('ultimoMes', mesAtivo);
        }
        salvarArquivoAutomaticamente();
        renderizarAbas();
        renderizarSubAbas();
        renderizarTabela();
        Swal.fire('Sucesso!', `${totalImportados} processos importados.`, 'success');
      }
      document.getElementById('input-importar-excel').value = '';
    });
  };
  reader.readAsArrayBuffer(file);
}
