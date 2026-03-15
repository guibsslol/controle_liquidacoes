let dadosAbas = { 'Assunto Geral': { Janeiro: [] } };
let abaAtiva = 'Assunto Geral';
let mesAtivo = 'Janeiro';

let arquivoHandle = null;
let abaArrastada = null;
let subAbaArrastada = null; // NOVO: Controla qual sub-aba (mês) está sendo movida
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

// ==========================================
// FUNÇÃO NOVA: DUPLICAR MÊS (COM ESCOLHA DE CAMPOS)
// ==========================================
async function duplicarMes() {
  const mesesDisponiveis = Object.keys(dadosAbas[abaAtiva]);
  if (mesesDisponiveis.length === 0) {
    Swal.fire('Erro', 'Não há meses para copiar.', 'error');
    return;
  }

  // Cria as opções de meses para o menu suspenso (selecionando o mês atual por padrão)
  let optionsHtml = '';
  mesesDisponiveis.forEach((m) => {
    optionsHtml += `<option value="${m}" ${m === mesAtivo ? 'selected' : ''}>${m}</option>`;
  });

  // Abre a janela perguntando o que deve ser copiado
  const { value: formValues } = await Swal.fire({
    title: 'Copiar Mês',
    html: `
            <div style="text-align: left; font-size: 14px;">
                <label style="font-weight: bold; color: #2c3e50;">1. Qual mês deseja copiar?</label>
                <select id="swal-origem" class="swal2-select" style="width: 100%; margin: 5px 0 15px 0; padding: 5px;">
                    ${optionsHtml}
                </select>
                
                <label style="font-weight: bold; color: #2c3e50;">2. Nome do Novo Mês:</label>
                <input id="swal-novo-mes" class="swal2-input" placeholder="Ex: Março-26" style="width: 100%; margin: 5px 0 15px 0; box-sizing: border-box;">
                
                <label style="font-weight: bold; color: #2c3e50;">3. Quais colunas deseja importar?</label>
                <div style="margin-top: 5px; display: grid; grid-template-columns: 1fr 1fr; gap: 8px; background: #f8f9fa; padding: 10px; border-radius: 5px; border: 1px solid #eee;">
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-proc" checked> Processo</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-emp" checked> Empresa</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-elem" checked> Elemento</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-empenho" checked> Empenho</label>
                    <label style="cursor: pointer;"><input type="checkbox" id="chk-liq"> Liquidação</label>
                </div>
                <div style="margin-top: 12px; font-size: 11px; color: #e74c3c; font-weight: bold;">
                    <i class="fa-solid fa-circle-info"></i> Atenção: O Status será resetado para "Aguardando Pagamento" e a OP ficará em branco.
                </div>
            </div>
        `,
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

    // 1. Cria o novo mês no banco de dados
    dadosAbas[abaAtiva][novoMes] = [];

    // 2. Varre o mês antigo e copia apenas as colunas que o usuário marcou
    dadosAbas[abaAtiva][origem].forEach((reg, index) => {
      dadosAbas[abaAtiva][novoMes].push({
        id: Date.now() + index, // ID único novo
        processo: importProc ? reg.processo : '',
        empresa: importEmp ? reg.empresa : '',
        elemento: importElem ? reg.elemento : '',
        empenho: importEmpenho ? reg.empenho : '',
        liquidacao: importLiq ? reg.liquidacao : '',
        status: 'Aguardando Pagamento', // Resetado
        op: '', // Resetado
      });
    });

    // 3. Muda a tela para o novo mês recém-criado
    mesAtivo = novoMes;

    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();

    Swal.fire(
      'Pronto!',
      `O mês "${novoMes}" foi criado e ${dadosAbas[abaAtiva][novoMes].length} processos foram importados!`,
      'success',
    );
  }
}

function renderizarSubAbas() {
  const listaMeses = document.getElementById('lista-sub-abas');
  listaMeses.innerHTML = '';

  if (!dadosAbas[abaAtiva]) return;

  Object.keys(dadosAbas[abaAtiva]).forEach((nomeMes) => {
    const divMes = document.createElement('div');
    divMes.className = `sub-aba ${nomeMes === mesAtivo ? 'ativa' : ''}`;

    // NOVO: Habilita o elemento para ser arrastado
    divMes.setAttribute('draggable', true);

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

    // ==========================================
    // EVENTOS DE ARRASTAR E SOLTAR (SUB-ABAS)
    // ==========================================

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

    // AÇÃO FINAL: Ao soltar a sub-aba
    divMes.addEventListener('drop', function (e) {
      e.stopPropagation();
      this.classList.remove('drag-over');

      if (subAbaArrastada !== nomeMes) {
        reordenarSubAbas(subAbaArrastada, nomeMes);
      }
      return false;
    });

    listaMeses.appendChild(divMes);
  });
}

// Nova função que reescreve a ordem dos meses no banco de dados
function reordenarSubAbas(mesOrigem, mesDestino) {
  const chaves = Object.keys(dadosAbas[abaAtiva]);
  const indexOrigem = chaves.indexOf(mesOrigem);
  const indexDestino = chaves.indexOf(mesDestino);

  // Remove do local antigo e insere no novo
  chaves.splice(indexOrigem, 1);
  chaves.splice(indexDestino, 0, mesOrigem);

  // Recria a estrutura do mês mantendo a nova ordem
  const novoObjetoMeses = {};
  chaves.forEach((chave) => {
    novoObjetoMeses[chave] = dadosAbas[abaAtiva][chave];
  });

  // Atualiza os dados, salva no arquivo e atualiza a tela
  dadosAbas[abaAtiva] = novoObjetoMeses;
  salvarArquivoAutomaticamente();
  renderizarSubAbas();
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

// ==========================================
// FUNÇÃO NOVA: AGRUPAR ABAS SOLTAS EM SUB-ABAS
// ==========================================
async function agruparAbas() {
  const abasDisponiveis = Object.keys(dadosAbas);
  if (abasDisponiveis.length < 2) {
    Swal.fire('Atenção', 'Você precisa de pelo menos 2 assuntos para agrupá-los.', 'info');
    return;
  }

  // Cria a lista de caixinhas (checkboxes) com as abas atuais
  let htmlCheckboxes =
    '<div style="text-align: left; max-height: 200px; overflow-y: auto; margin-top: 15px; padding: 10px; border: 1px solid #ccc; border-radius: 5px; background: #f9f9f9;">';
  abasDisponiveis.forEach((aba) => {
    htmlCheckboxes += `<label style="display: block; margin-bottom: 8px; cursor: pointer; font-size: 14px;"><input type="checkbox" class="swal-aba-checkbox" value="${aba}" style="margin-right: 8px;"> ${aba}</label>`;
  });
  htmlCheckboxes += '</div>';

  const { value: formValues } = await Swal.fire({
    title: 'Agrupar Assuntos',
    html:
      `<div style="font-size: 14px; text-align: left; margin-bottom: 10px;">Selecione os assuntos que deseja fundir e dê um nome para a nova Aba Principal:</div>` +
      `<input id="swal-input-novo-nome" class="swal2-input" placeholder="Nome do Novo Assunto (ex: Liquidações)" style="margin-top: 0;">` +
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
        Swal.showValidationMessage('Digite o nome do novo assunto principal!');
        return false;
      }
      return { selecionados, novoNome };
    },
  });

  if (formValues) {
    const { selecionados, novoNome } = formValues;

    // Se a nova aba principal ainda não existe, cria ela
    if (!dadosAbas[novoNome]) {
      dadosAbas[novoNome] = {};
    }

    selecionados.forEach((abaAntiga) => {
      // Varre os meses da aba antiga (normalmente só vai ter o "Geral" da importação)
      Object.keys(dadosAbas[abaAntiga]).forEach((mesAntigo) => {
        let novoNomeMes = abaAntiga;

        // MÁGICA: Limpa automaticamente o prefixo "LIQ." para o nome da sub-aba ficar bonito
        novoNomeMes = novoNomeMes
          .replace('LIQ. ', '')
          .replace('LIQ.', '')
          .replace('Liq. ', '')
          .replace('LIQ ', '')
          .trim();

        // Transfere os dados para o novo endereço
        if (dadosAbas[novoNome][novoNomeMes]) {
          dadosAbas[novoNome][novoNomeMes] = dadosAbas[novoNome][novoNomeMes].concat(
            dadosAbas[abaAntiga][mesAntigo],
          );
        } else {
          dadosAbas[novoNome][novoNomeMes] = dadosAbas[abaAntiga][mesAntigo];
        }
      });

      // Deleta a aba antiga principal da raiz do sistema
      if (abaAntiga !== novoNome) {
        delete dadosAbas[abaAntiga];
      }
    });

    // Aponta a tela para a nova aba criada
    abaAtiva = novoNome;
    mesAtivo = Object.keys(dadosAbas[novoNome])[0];

    salvarArquivoAutomaticamente();
    renderizarAbas();
    renderizarSubAbas();
    renderizarTabela();

    Swal.fire('Sucesso!', 'Os assuntos foram agrupados em Sub-Abas perfeitamente!', 'success');
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

// ==========================================
// 5. IMPORTAÇÃO DE EXCEL (CRIANDO ABAS PRINCIPAIS)
// ==========================================

function processarImportacaoExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    Swal.fire({
      title: 'Importar Planilha Completa?',
      text: `Encontramos ${workbook.SheetNames.length} páginas no Excel. Elas serão importadas como NOVOS ASSUNTOS (Abas Principais) lá no rodapé do sistema.`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Sim, importar tudo!',
      cancelButtonText: 'Cancelar',
    }).then((result) => {
      if (result.isConfirmed) {
        let totalImportados = 0;

        // Faz um loop passando por TODAS as abas do arquivo Excel
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          // A MÁGICA: Cria uma NOVA ABA PRINCIPAL com o nome da planilha do Excel
          // E cria um sub-mês padrão chamado "Geral" para guardar os dados
          if (!dadosAbas[sheetName]) {
            dadosAbas[sheetName] = { Geral: [] };
          } else if (!dadosAbas[sheetName]['Geral']) {
            dadosAbas[sheetName]['Geral'] = []; // Prevenção caso a aba já exista sem o mês Geral
          }

          // Começa do index 1 para pular a linha de cabeçalho do Excel
          for (let i = 1; i < json.length; i++) {
            const row = json[i];

            // Ignora linhas vazias
            if (!row || row.length === 0) continue;

            const proc = row[1] ? String(row[1]).trim() : '';
            const emp = row[2] ? String(row[2]).trim() : '';

            // Ignora a linha se a coluna Empresa estiver em branco
            if (!emp) continue;

            const elem = row[3] ? String(row[3]).trim() : '';
            const empen = row[4] ? String(row[4]).trim() : '';
            const liq = row[5] ? String(row[5]).trim() : '';
            const numOp = row[6] ? String(row[6]).trim() : '';

            const statusPag = numOp !== '' ? 'Pago' : 'Aguardando Pagamento';

            // Garante que os processos importados tenham o prefixo BJI-
            let procFormatado = proc;
            if (procFormatado && !procFormatado.toUpperCase().startsWith('BJI-')) {
              procFormatado = 'BJI-' + procFormatado;
            }

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

            // Salva os dados na aba principal nova, dentro do mês "Geral"
            dadosAbas[sheetName]['Geral'].push(novoReg);
            totalImportados++;
          }
        });

        // Muda a tela automaticamente para a primeira aba principal importada
        if (workbook.SheetNames.length > 0) {
          abaAtiva = workbook.SheetNames[0];
          mesAtivo = 'Geral';
        }

        salvarArquivoAutomaticamente();

        // Agora o sistema manda redesenhar as abas de baixo também!
        renderizarAbas();
        renderizarSubAbas();
        renderizarTabela();
        atualizarAutocompletarGlobal();

        Swal.fire(
          'Sucesso!',
          `${totalImportados} processos importados e organizados em ${workbook.SheetNames.length} novos Assuntos!`,
          'success',
        );
      }

      document.getElementById('input-importar-excel').value = '';
    });
  };
  reader.readAsArrayBuffer(file);
}

atualizarAutocompletarGlobal(); // Chamada inicial
renderizarAbas();
renderizarSubAbas();
renderizarTabela();
