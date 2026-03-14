let dadosAbas = { 'Mês Atual': [] };
let abaAtiva = 'Mês Atual';
let cadastros = { processos: [], empresas: [], elementos: [], empenhos: [] };
let arquivoHandle = null;

let abaArrastada = null; // Controla qual aba está sendo movida

// NOVA VARIÁVEL: Controla qual linha da tabela está virando formulário
let linhaEmEdicao = null;

// ==========================================
// 1. SISTEMA DE ARQUIVO
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
      Swal.fire('Acesso Negado', 'Você precisa permitir a edição para continuar.', 'warning');
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
    if (typeof idbKeyval !== 'undefined') {
      await idbKeyval.set('meuArquivoProcessos', arquivoHandle);
    }
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
  } else {
    await salvarArquivoAutomaticamente();
  }

  renderizarAbas();
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
  dadosAbas = { 'Mês Atual': [] };

  document.getElementById('status-conexao').innerHTML =
    '<i class="fa-solid fa-circle-xmark"></i> Desconectado';
  document.getElementById('status-conexao').style.color = '#e74c3c';

  const btnConectar = document.getElementById('btn-conectar');
  btnConectar.style.display = 'block';
  btnConectar.innerHTML = '<i class="fa-solid fa-link"></i> Conectar Banco';
  btnConectar.style.backgroundColor = '#8e44ad';
  btnConectar.onclick = conectarNovoArquivo;

  document.getElementById('btn-desconectar').style.display = 'none';

  renderizarAbas();
  renderizarTabela();
  Swal.fire('Desconectado', 'O banco de dados foi fechado com segurança.', 'info');
}

async function salvarArquivoAutomaticamente() {
  if (!arquivoHandle) return;
  try {
    const writable = await arquivoHandle.createWritable();
    await writable.write(JSON.stringify({ abas: dadosAbas, ativa: abaAtiva }));
    await writable.close();
  } catch (erro) {
    console.error('Erro ao salvar:', erro);
    Swal.fire(
      'Aviso',
      'Erro ao salvar. Verifique se o arquivo está aberto em outro programa.',
      'warning',
    );
  }
}

// ==========================================
// 2. SISTEMA DE ABAS
// ==========================================

function renderizarAbas() {
  const listaAbas = document.getElementById('lista-abas');
  listaAbas.innerHTML = '';

  Object.keys(dadosAbas).forEach((nomeAba) => {
    const divAba = document.createElement('div');
    divAba.className = `aba ${nomeAba === abaAtiva ? 'ativa' : ''}`;

    // Habilita o elemento para ser arrastado
    divAba.setAttribute('draggable', true);

    divAba.innerHTML = `
            <span class="titulo-aba">${nomeAba}</span>
            <div class="aba-acoes">
                <button class="btn-aba-acao edit" onclick="event.stopPropagation(); editarNomeAba('${nomeAba}')" title="Renomear"><i class="fa-solid fa-pen"></i></button>
                <button class="btn-aba-acao" onclick="event.stopPropagation(); excluirAba('${nomeAba}')" title="Excluir"><i class="fa-solid fa-trash"></i></button>
            </div>
        `;

    // Clique normal para abrir a aba
    divAba.onclick = () => {
      abaAtiva = nomeAba;
      linhaEmEdicao = null;
      salvarArquivoAutomaticamente();
      renderizarAbas();
      renderizarTabela();
    };

    // ==========================================
    // EVENTOS DE ARRASTAR E SOLTAR (DRAG & DROP)
    // ==========================================

    // Quando começa a arrastar
    divAba.addEventListener('dragstart', function (e) {
      abaArrastada = nomeAba;
      e.dataTransfer.effectAllowed = 'move';
      setTimeout(() => (this.style.opacity = '0.4'), 0); // Deixa a aba original meio transparente
    });

    // Quando termina de arrastar (solta o mouse)
    divAba.addEventListener('dragend', function () {
      this.style.opacity = '1';
      document.querySelectorAll('.aba').forEach((a) => a.classList.remove('drag-over'));
      abaArrastada = null;
    });

    // Permite que a aba seja solta em cima desta
    divAba.addEventListener('dragover', function (e) {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      return false;
    });

    // Efeito visual ao passar por cima
    divAba.addEventListener('dragenter', function (e) {
      if (nomeAba !== abaArrastada) this.classList.add('drag-over');
    });

    // Remove efeito visual ao sair de cima
    divAba.addEventListener('dragleave', function () {
      this.classList.remove('drag-over');
    });

    // AÇÃO FINAL: Ao soltar a aba
    divAba.addEventListener('drop', function (e) {
      e.stopPropagation();
      this.classList.remove('drag-over');

      if (abaArrastada !== nomeAba) {
        reordenarAbas(abaArrastada, nomeAba);
      }
      return false;
    });

    listaAbas.appendChild(divAba);
  });
}

// Nova função que reescreve a ordem no banco de dados
function reordenarAbas(abaOrigem, abaDestino) {
  const chaves = Object.keys(dadosAbas);
  const indexOrigem = chaves.indexOf(abaOrigem);
  const indexDestino = chaves.indexOf(abaDestino);

  // Remove a aba da posição antiga e insere na nova posição
  chaves.splice(indexOrigem, 1);
  chaves.splice(indexDestino, 0, abaOrigem);

  // Recria o objeto de dados mantendo a nova ordem das chaves
  const novoDadosAbas = {};
  chaves.forEach((chave) => {
    novoDadosAbas[chave] = dadosAbas[chave];
  });

  dadosAbas = novoDadosAbas; // Substitui o antigo pelo novo
  salvarArquivoAutomaticamente();
  renderizarAbas();
}

async function editarNomeAba(nomeAtual) {
  const { value: novoNome } = await Swal.fire({
    title: 'Renomear Aba',
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
      Swal.fire('Erro', 'Já existe uma aba com este nome.', 'error');
    }
  }
}

function excluirAba(nomeAba) {
  if (Object.keys(dadosAbas).length > 1) {
    Swal.fire({
      title: `Excluir a aba "${nomeAba}"?`,
      text: 'Os dados serão perdidos!',
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((result) => {
      if (result.isConfirmed) {
        delete dadosAbas[nomeAba];
        abaAtiva = Object.keys(dadosAbas)[0];
        salvarArquivoAutomaticamente();
        renderizarAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'Você precisa ter pelo menos uma aba.', 'info');
  }
}

async function criarNovaAba() {
  const { value: nome } = await Swal.fire({
    title: 'Nova Aba',
    input: 'text',
    showCancelButton: true,
  });
  if (nome && nome.trim() !== '') {
    if (!dadosAbas[nome]) {
      dadosAbas[nome] = [];
      abaAtiva = nome;
      salvarArquivoAutomaticamente();
      renderizarAbas();
      renderizarTabela();
    }
  }
}

// ==========================================
// 3. TABELA E EDIÇÃO INLINE
// ==========================================

function atualizarFiltrosDinamicos() {
  const registrosAba = dadosAbas[abaAtiva] || [];
  const processosUnicos = [...new Set(registrosAba.map((r) => r.processo).filter((p) => p !== ''))];
  const empresasUnicas = [...new Set(registrosAba.map((r) => r.empresa).filter((e) => e !== ''))];
  const elementosUnicos = [...new Set(registrosAba.map((r) => r.elemento).filter((e) => e !== ''))];
  const empenhosUnicos = [...new Set(registrosAba.map((r) => r.empenho).filter((e) => e !== ''))];

  const preencherSelect = (idSelect, lista, textoPadrao) => {
    const select = document.getElementById(idSelect);
    const valorAtual = select.value;
    select.innerHTML = `<option value="">${textoPadrao}</option>`;
    lista
      .sort()
      .forEach((item) => (select.innerHTML += `<option value="${item}">${item}</option>`));
    select.value = valorAtual;
  };

  preencherSelect('filtro-processo', processosUnicos, 'Todos os Processos');
  preencherSelect('filtro-empresa', empresasUnicas, 'Todas as Empresas');
  preencherSelect('filtro-elemento', elementosUnicos, 'Todos os Elementos');
  preencherSelect('filtro-empenho', empenhosUnicos, 'Todos os Empenhos');
}

// O formulário do topo agora serve APENAS para adicionar coisas novas
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
  dadosAbas[abaAtiva].push(novoReg);

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
  document.querySelectorAll('.form-row input').forEach((i) => (i.value = ''));
  document.getElementById('status_pagamento').value = 'Aguardando Pagamento';
}

// Funções para controle da Edição na Linha (Inline)
function ativarEdicaoInline(id) {
  linhaEmEdicao = id;
  renderizarTabela();
}

function cancelarEdicaoInline() {
  linhaEmEdicao = null;
  renderizarTabela();
}

function salvarEdicaoInline(id) {
  const index = dadosAbas[abaAtiva].findIndex((r) => r.id === id);
  if (index !== -1) {
    dadosAbas[abaAtiva][index].processo = document
      .getElementById(`edit-processo-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][index].empresa = document.getElementById(`edit-empresa-${id}`).value.trim();
    dadosAbas[abaAtiva][index].elemento = document
      .getElementById(`edit-elemento-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][index].empenho = document.getElementById(`edit-empenho-${id}`).value.trim();
    dadosAbas[abaAtiva][index].liquidacao = document
      .getElementById(`edit-liquidacao-${id}`)
      .value.trim();
    dadosAbas[abaAtiva][index].status = document.getElementById(`edit-status-${id}`).value;
    dadosAbas[abaAtiva][index].op = document.getElementById(`edit-op-${id}`).value.trim();

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

  atualizarFiltrosDinamicos();

  let registros = dadosAbas[abaAtiva] || [];
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

    // Verifica se esta é a linha que o usuário quer editar
    if (linhaEmEdicao === reg.id) {
      // Renderiza inputs dentro das células da tabela
      tr.innerHTML = `
                <td><input type="text" id="edit-processo-${reg.id}" class="input-inline" value="${reg.processo}"></td>
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
      // Deixa a linha de edição destacada
      tr.style.backgroundColor = '#fdfefe';
      tr.style.boxShadow = 'inset 0 0 5px rgba(52, 152, 219, 0.3)';
    } else {
      // Renderiza a linha normal de visualização
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
      dadosAbas[abaAtiva] = dadosAbas[abaAtiva].filter((r) => r.id !== id);
      salvarArquivoAutomaticamente();
      renderizarTabela();
    }
  });
}

function exportarParaExcel() {
  let tabela = document.getElementById('tabela-processos');
  let workbook = XLSX.utils.table_to_book(tabela, { sheet: abaAtiva });
  XLSX.writeFile(workbook, `Controle_Processos_${abaAtiva}.xlsx`);
}

renderizarAbas();
renderizarTabela();
