let dadosAbas = { 'Assunto Geral': { Janeiro: [] } };
let abaAtiva = 'Assunto Geral';
let mesAtivo = 'Janeiro';

let arquivoHandle = null;
let abaArrastada = null;
let linhaEmEdicao = null;

// Variáveis para a ordenação (Sort)
let colunaSort = 'id';
let ordemSort = 'asc';

// ==========================================
// MÁSCARA DO PROCESSO (BJI-) E TECLA ENTER
// ==========================================
function mascaraProcesso(e) {
  let input = e.target;
  // Se o usuário apagar tudo e sobrar só o prefixo, limpa o campo
  if (
    (input.value === 'BJI' || input.value === 'BJI-') &&
    e.inputType === 'deleteContentBackward'
  ) {
    input.value = '';
    return;
  }

  let val = input.value.toUpperCase().replace('BJI-', '');
  val = val.replace(/[^A-Z0-9]/g, ''); // Remove tudo que não for letra ou número

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

// Atalho do "Enter" para salvar o formulário
document.addEventListener('keydown', function (e) {
  if (e.key === 'Enter') {
    // Verifica se o usuário está digitando nos campos principais (e não nos filtros)
    const noFormulario = e.target.closest('.form-row') && !e.target.closest('.filtros-container');
    if (noFormulario) {
      e.preventDefault(); // Impede o Enter de fazer coisas indesejadas
      if (linhaEmEdicao) {
        salvarEdicaoInline(linhaEmEdicao);
      } else {
        adicionarRegistro();
      }
    }
  }
});

// ==========================================
// 1. SISTEMA DE ARQUIVO E MIGRAÇÃO
// ==========================================
window.onload = async () => {
  try {
    if (typeof idbKeyval !== 'undefined') {
      const savedHandle = await idbKeyval.get('meuArquivoProcessos');
      if (savedHandle) {
        arquivoHandle = savedHandle;
        const btn = document.getElementById('btn-conectar');
        btn.innerHTML = '<i class="fa-solid fa-unlock"></i> Retomar Sessão';
        btn.style.backgroundColor = '#2980b9';
        btn.onclick = retomarSessao;
        Swal.fire({
          icon: 'info',
          title: 'Sessão Encontrada',
          text: 'Clique em "Retomar Sessão" para continuar.',
          toast: true,
          position: 'top-end',
          timer: 3000,
        });
      }
    }
  } catch (e) {
    console.log('Nenhuma sessão anterior encontrada.');
  }
};

async function retomarSessao() {
  try {
    const options = { mode: 'readwrite' };
    if ((await arquivoHandle.requestPermission(options)) === 'granted') {
      await lerDadosDoArquivo();
    } else {
      Swal.fire('Acesso Negado', 'Você precisa permitir a edição.', 'warning');
    }
  } catch (erro) {
    Swal.fire('Erro', 'O arquivo original foi movido ou excluído. Conecte novamente.', 'error');
    desconectarArquivo();
  }
}

async function conectarNovoArquivo() {
  try {
    [arquivoHandle] = await window.showOpenFilePicker({
      types: [{ description: 'Banco de Dados', accept: { 'application/json': ['.json'] } }],
      multiple: false,
    });
  } catch (erro) {
    return;
  }

  try {
    if (typeof idbKeyval !== 'undefined') await idbKeyval.set('meuArquivoProcessos', arquivoHandle);
    await lerDadosDoArquivo();
  } catch (erro) {
    Swal.fire('Erro ao Conectar', 'Motivo: ' + erro.message, 'error');
  }
}

async function lerDadosDoArquivo() {
  const file = await arquivoHandle.getFile();
  const contents = await file.text();
  let importado = {};

  try {
    if (contents.trim() !== '') importado = JSON.parse(contents);
  } catch (e) {}

  if (importado && importado.abas) {
    dadosAbas = importado.abas;
    abaAtiva = importado.ativa || Object.keys(dadosAbas)[0];

    // MÁGICA DA MIGRAÇÃO: Se os dados antigos não têm meses, ele cria o mês "Geral" para não quebrar.
    let precisaSalvar = false;
    Object.keys(dadosAbas).forEach((aba) => {
      if (Array.isArray(dadosAbas[aba])) {
        dadosAbas[aba] = { Geral: dadosAbas[aba] };
        precisaSalvar = true;
      }
    });

    mesAtivo = Object.keys(dadosAbas[abaAtiva])[0] || 'Geral';
    if (precisaSalvar) await salvarArquivoAutomaticamente();
  } else {
    await salvarArquivoAutomaticamente();
  }

  renderizarAbas();
  renderizarSubAbas();
  renderizarTabela();

  document.getElementById('status-conexao').innerHTML =
    '<i class="fa-solid fa-circle-check"></i> Conectado e Salvando';
  document.getElementById('status-conexao').style.color = '#27ae60';
  document.getElementById('btn-conectar').style.display = 'none';
  document.getElementById('btn-desconectar').style.display = 'block';

  Swal.fire({
    icon: 'success',
    title: 'Banco Conectado!',
    toast: true,
    position: 'top-end',
    timer: 2000,
  });
}

async function desconectarArquivo() {
  if (typeof idbKeyval !== 'undefined') await idbKeyval.del('meuArquivoProcessos');
  arquivoHandle = null;
  dadosAbas = { 'Assunto Geral': { Janeiro: [] } };

  document.getElementById('status-conexao').innerHTML =
    '<i class="fa-solid fa-circle-xmark"></i> Desconectado';
  document.getElementById('status-conexao').style.color = '#e74c3c';
  const btn = document.getElementById('btn-conectar');
  btn.style.display = 'block';
  btn.innerHTML = '<i class="fa-solid fa-link"></i> Conectar Banco';
  btn.style.backgroundColor = '#8e44ad';
  btn.onclick = conectarNovoArquivo;
  document.getElementById('btn-desconectar').style.display = 'none';

  renderizarAbas();
  renderizarSubAbas();
  renderizarTabela();
  Swal.fire('Desconectado', 'O banco fechou com segurança.', 'info');
}

async function salvarArquivoAutomaticamente() {
  if (!arquivoHandle) return;
  try {
    const writable = await arquivoHandle.createWritable();
    await writable.write(JSON.stringify({ abas: dadosAbas, ativa: abaAtiva }));
    await writable.close();
    atualizarAutocompletarGlobal(); // Atualiza a memória de autocompletar ao salvar
  } catch (erro) {
    console.error('Erro ao salvar:', erro);
  }
}

// ==========================================
// 2. SISTEMA DE ABAS E SUB-ABAS (MESES)
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
        </div>
    `;

    divAba.onclick = () => {
      abaAtiva = nomeAba;
      mesAtivo = Object.keys(dadosAbas[abaAtiva])[0]; // Pula para o primeiro mês da aba
      linhaEmEdicao = null;
      salvarArquivoAutomaticamente();
      renderizarAbas();
      renderizarSubAbas();
      renderizarTabela();
    };

    // Drag and Drop
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

    divMes.innerHTML = `
            <span>${nomeMes}</span>
            <div class="aba-acoes" style="display: ${nomeMes === mesAtivo ? 'flex' : 'none'}">
                <button class="btn-aba-acao edit" onclick="event.stopPropagation(); editarNomeMes('${nomeMes}')" title="Renomear"><i class="fa-solid fa-pen"></i></button>
                <button class="btn-aba-acao" onclick="event.stopPropagation(); excluirMes('${nomeMes}')" title="Excluir"><i class="fa-solid fa-trash"></i></button>
            </div>
        `;

    divMes.onclick = () => {
      mesAtivo = nomeMes;
      linhaEmEdicao = null;
      renderizarSubAbas();
      renderizarTabela();
    };

    listaMeses.appendChild(divMes);
  });
}

// Funções de CRUD das Abas e Meses (Renomear, Excluir, Criar)
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
      if (abaAtiva === nomeAtual) abaAtiva = novoNome;
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
      text: 'Todos os meses e dados serão perdidos!',
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((result) => {
      if (result.isConfirmed) {
        delete dadosAbas[nomeAba];
        abaAtiva = Object.keys(dadosAbas)[0];
        mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
        salvarArquivoAutomaticamente();
        renderizarAbas();
        renderizarSubAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'Você precisa ter pelo menos uma aba de assunto.', 'info');
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
      dadosAbas[nome] = { Geral: [] }; // Cria o assunto já com um mês padrão
      abaAtiva = nome;
      mesAtivo = 'Geral';
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
      if (mesAtivo === nomeAtual) mesAtivo = novoNome;
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
      text: 'Os dados do mês serão perdidos!',
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((result) => {
      if (result.isConfirmed) {
        delete dadosAbas[abaAtiva][nomeMes];
        mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
        salvarArquivoAutomaticamente();
        renderizarSubAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'O assunto precisa ter pelo menos um mês.', 'info');
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
      salvarArquivoAutomaticamente();
      renderizarSubAbas();
      renderizarTabela();
    }
  }
}

// ==========================================
// 3. AUTOCOMPLETAR GLOBAL & ORDENAÇÃO
// ==========================================

function atualizarAutocompletarGlobal() {
  // Escaneia TODOS os processos de TODAS as abas e meses para criar a memória perfeita
  let procs = new Set(),
    emps = new Set(),
    elems = new Set(),
    emps2 = new Set();

  Object.values(dadosAbas).forEach((aba) => {
    Object.values(aba).forEach((mes) => {
      mes.forEach((reg) => {
        if (reg.processo) procs.add(reg.processo);
        if (reg.empresa) emps.add(reg.empresa);
        if (reg.elemento) elems.add(reg.elemento);
        if (reg.empenho) emps2.add(reg.empenho);
      });
    });
  });

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
  if (colunaSort === coluna) {
    ordemSort = ordemSort === 'asc' ? 'desc' : 'asc';
  } else {
    colunaSort = coluna;
    ordemSort = 'asc';
  }
  renderizarTabela();
}

// ==========================================
// 4. TABELA E EDIÇÃO INLINE
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

  const novoReg = { id: Date.now(), processo, empresa, elemento, empenho, liquidacao, status, op };

  if (!dadosAbas[abaAtiva][mesAtivo]) dadosAbas[abaAtiva][mesAtivo] = []; // Prevenção
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
  document.getElementById('processo').focus(); // Foca no primeiro campo após limpar
}

function ativarEdicaoInline(id) {
  linhaEmEdicao = id;
  renderizarTabela();
}
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

function renderizarTabela() {
  const tbody = document.getElementById('tabela-corpo');
  tbody.innerHTML = '';

  let registros =
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo] ? [...dadosAbas[abaAtiva][mesAtivo]] : [];

  // Atualiza os ícones de ordenação no cabeçalho
  document.querySelectorAll('.th-sortable i').forEach((icon) => {
    icon.className = 'fa-solid fa-sort'; // reseta todos
    icon.parentElement.classList.remove('sorted-asc', 'sorted-desc');
  });
  const currentIcon = document.getElementById(`sort-icon-${colunaSort}`);
  if (currentIcon) {
    currentIcon.className = ordemSort === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';
    currentIcon.parentElement.classList.add(`sorted-${ordemSort}`);
  }

  // Aplica a Ordenação (A-Z / Z-A)
  registros.sort((a, b) => {
    let valA = (a[colunaSort] || '').toString().toLowerCase();
    let valB = (b[colunaSort] || '').toString().toLowerCase();
    if (valA < valB) return ordemSort === 'asc' ? -1 : 1;
    if (valA > valB) return ordemSort === 'asc' ? 1 : -1;
    return 0;
  });

  // Aplica os Filtros
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
    cardTotal.innerText = registros.length;
    document.getElementById('resumo-pendente').innerText = registros.filter(
      (r) => r.status === 'Aguardando Pagamento',
    ).length;
    document.getElementById('resumo-pago').innerText = registros.filter(
      (r) => r.status === 'Pago',
    ).length;
  }

  registros.forEach((reg) => {
    const tr = document.createElement('tr');

    if (linhaEmEdicao === reg.id) {
      tr.innerHTML = `
        <td><input type="text" id="edit-processo-${reg.id}" class="input-inline" value="${reg.processo}" oninput="mascaraProcesso(event)"></td>
        <td><input type="text" id="edit-empresa-${reg.id}" class="input-inline" value="${reg.empresa}"></td>
        <td><input type="text" id="edit-elemento-${reg.id}" class="input-inline" value="${reg.elemento}"></td>
        <td><input type="text" id="edit-empenho-${reg.id}" class="input-inline" value="${reg.empenho}"></td>
        <td><input type="text" id="edit-liquidacao-${reg.id}" class="input-inline" value="${reg.liquidacao}"></td>
        <td>
            <select id="edit-status-${reg.id}" class="input-inline" style="padding: 5px;">
                <option value="Aguardando Pagamento" ${reg.status === 'Aguardando Pagamento' ? 'selected' : ''}>Aguardando Pagamento</option>
                <option value="Pago" ${reg.status === 'Pago' ? 'selected' : ''}>Pago</option>
            </select>
        </td>
        <td><input type="text" id="edit-op-${reg.id}" class="input-inline" value="${reg.op}"></td>
        <td class="coluna-acao">
            <button class="btn-edit" onclick="salvarEdicaoInline(${reg.id})" title="Confirmar" style="background-color: #27ae60;"><i class="fa-solid fa-check"></i></button>
            <button class="btn-delete" onclick="cancelarEdicaoInline()" title="Cancelar" style="background-color: #95a5a6;"><i class="fa-solid fa-xmark"></i></button>
        </td>
      `;
      tr.style.backgroundColor = '#fdfefe';
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
        <td>${reg.processo}</td>
        <td>${reg.empresa}</td>
        <td>${reg.elemento}</td>
        <td class="${classEmpenhoLiq}">${reg.empenho}</td>
        <td class="${classEmpenhoLiq}">${reg.liquidacao}</td>
        <td>${reg.status}</td>
        <td class="${classOP}">${reg.op}</td>
        <td class="coluna-acao">
            <button class="btn-edit" onclick="ativarEdicaoInline(${reg.id})" title="Editar na Linha"><i class="fa-solid fa-pen"></i></button>
            <button class="btn-delete" onclick="apagarLinha(${reg.id})" title="Excluir"><i class="fa-solid fa-trash"></i></button>
        </td>
      `;
    }
    tbody.appendChild(tr);
  });

  // Atualiza as opções dos filtros dinamicamente após renderizar as linhas
  atualizarFiltrosDinâmicosDaTela(registros);
}

// Uma função separada apenas para os filtros da tela atual não perderem as opções
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

atualizarAutocompletarGlobal(); // Chamada inicial
renderizarAbas();
renderizarSubAbas();
renderizarTabela();
