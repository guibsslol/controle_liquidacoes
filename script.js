const firebaseConfig = {
  apiKey: 'AIzaSyAIQrfYY0QvtyN7qj61uLhq6Xyb4Eyn3ZA',
  authDomain: 'controliqui-smebji.firebaseapp.com',
  databaseURL: 'https://controliqui-smebji-default-rtdb.firebaseio.com',
  projectId: 'controliqui-smebji',
  storageBucket: 'controliqui-smebji.firebasestorage.app',
  messagingSenderId: '659644181097',
  appId: '1:659644181097:web:82b55c5eba921bc06e10f8',
};

firebase.initializeApp(firebaseConfig);
const database = firebase.database();
const auth = firebase.auth();
const adminAuthApp = firebase.initializeApp(firebaseConfig, 'AdminAuthApp');

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

let currentUser = null;
let currentRole = 'guest';
let currentUserName = '';

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
  if (e.ctrlKey && e.shiftKey && e.key === 'P') {
    e.preventDefault();
    abrirPainelAdmin();
  }
});

window.onload = () => {
  if (localStorage.getItem('tema') === 'dark') document.body.setAttribute('data-theme', 'dark');
  auth.onAuthStateChanged((user) => {
    if (user) {
      currentUser = user;
      database
        .ref('usuarios/' + user.uid)
        .once('value')
        .then((snap) => {
          if (snap.exists()) {
            currentRole = snap.val().cargo;
            currentUserName = snap.val().nome;
            aplicarInterfaceUsuario();
          }
        });
    } else {
      currentUser = null;
      currentRole = 'guest';
      currentUserName = '';
      aplicarInterfaceUsuario();
    }
  });

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
      dadosAbas = { 'Assunto Geral': { Janeiro: [] } };
      abaAtiva = 'Assunto Geral';
      mesAtivo = 'Janeiro';
    }
    renderizarAbas();
    renderizarSubAbas();
    renderizarTabela();
    atualizarAutocompletarAba();
    const statusEl = document.getElementById('status-conexao');
    statusEl.innerHTML = '<i class="fa-solid fa-cloud"></i> Online';
    statusEl.style.color = '#27ae60';
    statusEl.style.borderColor = '#27ae60';
  });
};

function aplicarInterfaceUsuario() {
  document.body.setAttribute('data-role', currentRole);
  if (currentRole !== 'guest') {
    document.getElementById('nome-usuario-logado').style.display = 'inline';
    document.getElementById('nome-usuario-logado').innerText =
      `Olá, ${currentUserName.split(' ')[0]}`;
    document.getElementById('btn-login').style.display = 'none';
    document.getElementById('btn-logout').style.display = 'inline-block';
  } else {
    document.getElementById('nome-usuario-logado').style.display = 'none';
    document.getElementById('btn-login').style.display = 'inline-block';
    document.getElementById('btn-logout').style.display = 'none';
  }
  if (currentRole === 'admin_geral' || currentRole === 'admin_comum') {
    document.querySelectorAll('.admin-only').forEach((el) => (el.style.display = 'inline-block'));
  } else {
    document.querySelectorAll('.admin-only').forEach((el) => (el.style.display = 'none'));
  }
}

function registrarLog(acao, detalhes) {
  if (currentRole === 'guest') return;
  database.ref('logs').push({
    data: new Date().toISOString(),
    usuario: currentUserName,
    acao: acao,
    detalhes: detalhes,
    local: `${abaAtiva} > ${mesAtivo}`,
  });
}

function salvarArquivoAutomaticamente() {
  if (currentRole === 'guest') return;

  // A MÁGICA: Evita que o Firebase apague as abas vazias injetando um Fantasma (Dummy)
  for (let aba in dadosAbas) {
    for (let mes in dadosAbas[aba]) {
      let reais = dadosAbas[aba][mes] ? dadosAbas[aba][mes].filter((r) => !r.isDummy) : [];
      if (reais.length === 0) {
        dadosAbas[aba][mes] = [
          {
            id: 'dummy',
            isDummy: true,
            processo: '',
            empresa: 'Aba Vazia',
            elemento: '',
            empenho: '',
            liquidacao: '',
            status: 'Aguardando Pagamento',
            op: '',
          },
        ];
      } else {
        dadosAbas[aba][mes] = reais;
      }
    }
  }

  database
    .ref('sistema')
    .set({ abas: dadosAbas, ativa: abaAtiva })
    .catch((e) => console.log(e));
}

// ==========================================
// FUNÇÕES DE LOGIN E ADMIN
// ==========================================
function abrirModalLogin() {
  document.getElementById('modal-login').style.display = 'flex';
  const savedUser = localStorage.getItem('savedUsername');
  if (savedUser) {
    document.getElementById('login-username').value = savedUser;
    document.getElementById('login-lembrar').checked = true;
    document.getElementById('login-senha').focus();
  } else {
    document.getElementById('login-username').focus();
  }
}
function fecharModalLogin() {
  document.getElementById('modal-login').style.display = 'none';
}

function fazerLogin() {
  const user = document.getElementById('login-username').value.trim().toLowerCase();
  const senha = document.getElementById('login-senha').value;
  const lembrar = document.getElementById('login-lembrar').checked;

  if (!user || !senha) return Swal.fire('Aviso', 'Preencha o usuário e a senha', 'warning');
  const emailFake = user + '@bji.local';

  auth
    .signInWithEmailAndPassword(emailFake, senha)
    .then(() => {
      if (lembrar) localStorage.setItem('savedUsername', user);
      else localStorage.removeItem('savedUsername');
      fecharModalLogin();
      document.getElementById('login-senha').value = '';
      Swal.fire({
        icon: 'success',
        title: 'Login efetuado!',
        toast: true,
        position: 'top-end',
        timer: 2000,
        showConfirmButton: false,
      });
    })
    .catch((error) => Swal.fire('Erro', 'Credenciais inválidas.', 'error'));
}

function fazerLogout() {
  auth.signOut();
  Swal.fire({
    icon: 'info',
    title: 'Sessão terminada.',
    toast: true,
    position: 'top-end',
    timer: 2000,
    showConfirmButton: false,
  });
}

function abrirPainelAdmin() {
  database
    .ref('usuarios')
    .once('value')
    .then((snap) => {
      if (!snap.exists() || currentRole === 'admin_geral' || currentRole === 'admin_comum') {
        document.getElementById('modal-admin').style.display = 'flex';
        mudarAbaAdmin('usuarios');
      } else {
        Swal.fire('Acesso Negado', 'Apenas administradores podem acessar.', 'error');
      }
    });
}
function fecharPainelAdmin() {
  document.getElementById('modal-admin').style.display = 'none';
}

function mudarAbaAdmin(aba) {
  document.getElementById('admin-tab-usuarios').style.display =
    aba === 'usuarios' ? 'block' : 'none';
  document.getElementById('admin-tab-logs').style.display = aba === 'logs' ? 'block' : 'none';
  document.getElementById('tab-usuarios-btn').style.borderBottomColor =
    aba === 'usuarios' ? '#3498db' : 'transparent';
  document.getElementById('tab-logs-btn').style.borderBottomColor =
    aba === 'logs' ? '#3498db' : 'transparent';
  if (aba === 'usuarios') renderizarUsuarios();
  if (aba === 'logs') renderizarLogs();
}

function renderizarUsuarios() {
  database
    .ref('usuarios')
    .once('value')
    .then((snap) => {
      const tbody = document.getElementById('tabela-usuarios-corpo');
      tbody.innerHTML = '';
      if (!snap.exists()) return;
      const formatarCargo = (c) =>
        c === 'admin_geral'
          ? '<b style="color:#8e44ad">Admin Geral</b>'
          : c === 'admin_comum'
            ? '<b style="color:#2980b9">Admin Comum</b>'
            : 'Utilizador Comum';
      snap.forEach((child) => {
        const u = child.val();
        let tr = document.createElement('tr');
        let btnApagar = `<button onclick="apagarUsuario('${child.key}', '${u.cargo}')" style="padding:4px 8px; background:#e74c3c; color:white; border:none; border-radius:4px;"><i class="fa-solid fa-trash"></i></button>`;
        if (u.cargo === 'admin_geral' && currentRole !== 'admin_geral') btnApagar = '';
        if (currentUser && child.key === currentUser.uid)
          btnApagar =
            '<span style="color:#27ae60; font-size:11px; font-weight:bold;">(Você)</span>';
        tr.innerHTML = `<td>${u.nome}</td><td>${u.username}</td><td>${formatarCargo(u.cargo)}</td><td style="text-align:center;">${btnApagar}</td>`;
        tbody.appendChild(tr);
      });
    });
}

function criarNovoUsuario() {
  const nome = document.getElementById('novo-user-nome').value.trim();
  const username = document.getElementById('novo-user-username').value.trim().toLowerCase();
  const senha = document.getElementById('novo-user-senha').value;
  let cargo = document.getElementById('novo-user-cargo').value;

  if (!nome || !username || senha.length < 6)
    return Swal.fire('Aviso', 'Preencha tudo corretamente. Senha mínima 6 caracteres.', 'warning');
  const emailFake = username + '@bji.local';

  adminAuthApp
    .auth()
    .createUserWithEmailAndPassword(emailFake, senha)
    .then((userCredential) => {
      const uid = userCredential.user.uid;
      database
        .ref('usuarios')
        .once('value')
        .then((snap) => {
          if (!snap.exists()) {
            cargo = 'admin_geral';
            Swal.fire(
              'Bem-vindo, Chefe!',
              'O primeiro utilizador é sempre Admin Geral.',
              'success',
            );
          } else {
            Swal.fire('Sucesso!', 'Utilizador criado.', 'success');
          }

          database
            .ref('usuarios/' + uid)
            .set({ nome: nome, username: username, cargo: cargo })
            .then(() => {
              registrarLog('Criação de Utilizador', `Criou a conta de ${nome} (${cargo})`);
              document.getElementById('novo-user-nome').value = '';
              document.getElementById('novo-user-username').value = '';
              document.getElementById('novo-user-senha').value = '';
              renderizarUsuarios();
              adminAuthApp.auth().signOut();
            });
        });
    })
    .catch((e) => Swal.fire('Erro', e.message, 'error'));
}

function apagarUsuario(uid, cargoAlvo) {
  if (cargoAlvo === 'admin_geral' && currentRole !== 'admin_geral')
    return Swal.fire('Acesso Negado', 'Não tem permissão para apagar o Admin Geral.', 'error');
  Swal.fire({
    title: 'Apagar utilizador?',
    text: 'Ele perderá o acesso.',
    icon: 'warning',
    showCancelButton: true,
  }).then((r) => {
    if (r.isConfirmed) {
      database
        .ref('usuarios/' + uid)
        .remove()
        .then(() => {
          registrarLog('Exclusão de Utilizador', `Apagou o utilizador UID: ${uid}`);
          renderizarUsuarios();
          Swal.fire('Apagado!', '', 'success');
        });
    }
  });
}

function renderizarLogs() {
  const filtro = document.getElementById('filtro-log-texto').value.toLowerCase();
  database
    .ref('logs')
    .orderByChild('data')
    .limitToLast(100)
    .once('value')
    .then((snap) => {
      const tbody = document.getElementById('tabela-logs-corpo');
      tbody.innerHTML = '';
      if (!snap.exists()) return;
      let logs = [];
      snap.forEach((c) => {
        logs.push(c.val());
      });
      logs.reverse();
      logs.forEach((log) => {
        if (
          filtro &&
          !log.usuario.toLowerCase().includes(filtro) &&
          !log.detalhes.toLowerCase().includes(filtro) &&
          !log.acao.toLowerCase().includes(filtro)
        )
          return;
        let dataFormatada = new Date(log.data).toLocaleString('pt-BR');
        let corAcao =
          log.acao.includes('Exclusão') || log.acao.includes('Apagou')
            ? '#e74c3c'
            : log.acao.includes('Adição') || log.acao.includes('Novo')
              ? '#27ae60'
              : '#2980b9';
        tbody.innerHTML += `<tr><td style="color:#7f8c8d;">${dataFormatada}</td><td style="font-weight:bold;">${log.usuario}</td><td><span style="background:${corAcao}; color:white; padding:2px 6px; border-radius:3px; font-size:10px;">${log.acao}</span></td><td>${log.detalhes} <br><span style="font-size:10px; color:#95a5a6;">Em: ${log.local}</span></td></tr>`;
      });
    });
}

function limparLogs() {
  if (currentRole !== 'admin_geral')
    return Swal.fire('Acesso Negado', 'Só o Admin Geral pode limpar o histórico.', 'error');
  Swal.fire({
    title: 'Limpar todo o histórico?',
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#d33',
  }).then((r) => {
    if (r.isConfirmed) {
      database
        .ref('logs')
        .remove()
        .then(() => {
          renderizarLogs();
          Swal.fire('Limpo', '', 'success');
        });
    }
  });
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
    divAba.setAttribute('draggable', currentRole !== 'guest');
    divAba.innerHTML = `<span class="titulo-aba">${nomeAba}</span><div class="aba-acoes"><button class="btn-aba-acao edit" onclick="event.stopPropagation(); editarNomeAba('${nomeAba}')"><i class="fa-solid fa-pen"></i></button><button class="btn-aba-acao" onclick="event.stopPropagation(); excluirAba('${nomeAba}')"><i class="fa-solid fa-trash"></i></button></div>`;

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

    if (currentRole !== 'guest') {
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
          registrarLog('Ordenação', `Reordenou o assunto ${abaArrastada}`);
          salvarArquivoAutomaticamente();
          renderizarAbas();
        }
        return false;
      });
    }
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
    divMes.setAttribute('draggable', currentRole !== 'guest');
    divMes.innerHTML = `<span>${nomeMes}</span><div class="aba-acoes" style="display: ${nomeMes === mesAtivo ? 'flex' : 'none'}"><button class="btn-aba-acao edit" onclick="event.stopPropagation(); editarNomeMes('${nomeMes}')"><i class="fa-solid fa-pen"></i></button><button class="btn-aba-acao" onclick="event.stopPropagation(); excluirMes('${nomeMes}')"><i class="fa-solid fa-trash"></i></button></div>`;

    divMes.onclick = () => {
      mesAtivo = nomeMes;
      linhaEmEdicao = null;
      IDsSelecionados.clear();
      localStorage.setItem('ultimoMes', mesAtivo);
      renderizarSubAbas();
      renderizarTabela();
    };

    if (currentRole !== 'guest') {
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
        if (subAbaArrastada !== nomeMes) {
          const chaves = Object.keys(dadosAbas[abaAtiva]);
          chaves.splice(chaves.indexOf(subAbaArrastada), 1);
          chaves.splice(chaves.indexOf(nomeMes), 0, subAbaArrastada);
          const novoObjetoMeses = {};
          chaves.forEach((chave) => {
            novoObjetoMeses[chave] = dadosAbas[abaAtiva][chave];
          });
          dadosAbas[abaAtiva] = novoObjetoMeses;
          registrarLog('Ordenação', `Reordenou o mês ${subAbaArrastada}`);
          salvarArquivoAutomaticamente();
          renderizarSubAbas();
        }
        return false;
      });
    }
    listaMeses.appendChild(divMes);
  });
}

async function duplicarMes() {
  if (currentRole === 'guest') return;
  const mesesDisponiveis = Object.keys(dadosAbas[abaAtiva]);
  if (mesesDisponiveis.length === 0) return Swal.fire('Erro', 'Não há meses para copiar.', 'error');
  let optionsHtml = '';
  mesesDisponiveis.forEach((m) => {
    optionsHtml += `<option value="${m}" ${m === mesAtivo ? 'selected' : ''}>${m}</option>`;
  });

  const { value: formValues } = await Swal.fire({
    title: 'Copiar Mês',
    html: `<div style="text-align: left; font-size: 14px;"><label>Origem:</label><select id="swal-origem" class="swal2-select" style="width: 100%; margin: 5px 0 15px 0;">${optionsHtml}</select><label>Novo Mês:</label><input id="swal-novo-mes" class="swal2-input" style="width: 100%; margin: 5px 0 15px 0;"><label>Importar:</label><div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;"><label><input type="checkbox" id="chk-proc" checked> Processo</label><label><input type="checkbox" id="chk-emp" checked> Empresa</label><label><input type="checkbox" id="chk-elem" checked> Elemento</label><label><input type="checkbox" id="chk-empenho" checked> Empenho</label><label><input type="checkbox" id="chk-liq"> Liquidação</label></div></div>`,
    focusConfirm: false,
    showCancelButton: true,
    preConfirm: () => {
      const o = document.getElementById('swal-origem').value;
      const n = document.getElementById('swal-novo-mes').value.trim();
      if (!n || dadosAbas[abaAtiva][n])
        return Swal.showValidationMessage('Nome inválido ou existente!');
      return {
        origem: o,
        novoMes: n,
        importProc: document.getElementById('chk-proc').checked,
        importEmp: document.getElementById('chk-emp').checked,
        importElem: document.getElementById('chk-elem').checked,
        importEmpenho: document.getElementById('chk-empenho').checked,
        importLiq: document.getElementById('chk-liq').checked,
      };
    },
  });

  if (formValues) {
    dadosAbas[abaAtiva][formValues.novoMes] = [];
    dadosAbas[abaAtiva][formValues.origem].forEach((reg, index) => {
      if (reg.isDummy) return;
      dadosAbas[abaAtiva][formValues.novoMes].push({
        id: Date.now() + index,
        processo: formValues.importProc ? reg.processo : '',
        empresa: formValues.importEmp ? reg.empresa : '',
        elemento: formValues.importElem ? reg.elemento : '',
        empenho: formValues.importEmpenho ? reg.empenho : '',
        liquidacao: formValues.importLiq ? reg.liquidacao : '',
        status: 'Aguardando Pagamento',
        op: '',
      });
    });
    mesAtivo = formValues.novoMes;
    localStorage.setItem('ultimoMes', mesAtivo);
    registrarLog('Cópia de Mês', `Copiou o mês ${formValues.origem} gerando ${formValues.novoMes}`);
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
  }
}

async function editarNomeAba(nomeAtual) {
  if (currentRole === 'guest') return;
  const { value: novoNome } = await Swal.fire({
    title: 'Renomear',
    input: 'text',
    inputValue: nomeAtual,
    showCancelButton: true,
  });
  if (novoNome && novoNome.trim() !== '' && novoNome !== nomeAtual && !dadosAbas[novoNome]) {
    dadosAbas[novoNome] = dadosAbas[nomeAtual];
    delete dadosAbas[nomeAtual];
    if (abaAtiva === nomeAtual) {
      abaAtiva = novoNome;
      localStorage.setItem('ultimaAba', abaAtiva);
    }
    registrarLog('Edição de Assunto', `Renomeou ${nomeAtual} para ${novoNome}`);
    salvarArquivoAutomaticamente();
    renderizarAbas();
  }
}
function excluirAba(nomeAba) {
  if (currentRole === 'guest') return;
  if (Object.keys(dadosAbas).length > 1) {
    Swal.fire({
      title: `Excluir "${nomeAba}"?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((r) => {
      if (r.isConfirmed) {
        delete dadosAbas[nomeAba];
        abaAtiva = Object.keys(dadosAbas)[0];
        mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
        registrarLog('Exclusão de Assunto', `Apagou o assunto ${nomeAba} inteiro`);
        salvarArquivoAutomaticamente();
        renderizarAbas();
        renderizarSubAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'Mínimo de 1 aba.', 'info');
  }
}
async function criarNovaAba() {
  if (currentRole === 'guest') return;
  const { value: nome } = await Swal.fire({
    title: 'Novo Assunto',
    input: 'text',
    showCancelButton: true,
  });
  if (nome && nome.trim() !== '' && !dadosAbas[nome]) {
    dadosAbas[nome] = { Geral: [] };
    abaAtiva = nome;
    mesAtivo = 'Geral';
    registrarLog('Novo Assunto', `Criou o assunto ${nome}`);
    salvarArquivoAutomaticamente();
    renderizarAbas();
    renderizarSubAbas();
    renderizarTabela();
  }
}
async function editarNomeMes(nomeAtual) {
  if (currentRole === 'guest') return;
  const { value: novoNome } = await Swal.fire({
    title: 'Renomear Mês',
    input: 'text',
    inputValue: nomeAtual,
    showCancelButton: true,
  });
  if (
    novoNome &&
    novoNome.trim() !== '' &&
    novoNome !== nomeAtual &&
    !dadosAbas[abaAtiva][novoNome]
  ) {
    dadosAbas[abaAtiva][novoNome] = dadosAbas[abaAtiva][nomeAtual];
    delete dadosAbas[abaAtiva][nomeAtual];
    if (mesAtivo === nomeAtual) {
      mesAtivo = novoNome;
    }
    registrarLog('Edição de Mês', `Renomeou mês ${nomeAtual} para ${novoNome}`);
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
  }
}
function excluirMes(nomeMes) {
  if (currentRole === 'guest') return;
  if (Object.keys(dadosAbas[abaAtiva]).length > 1) {
    Swal.fire({
      title: `Excluir o mês "${nomeMes}"?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonColor: '#d33',
    }).then((r) => {
      if (r.isConfirmed) {
        delete dadosAbas[abaAtiva][nomeMes];
        mesAtivo = Object.keys(dadosAbas[abaAtiva])[0];
        registrarLog('Exclusão de Mês', `Apagou o mês ${nomeMes}`);
        salvarArquivoAutomaticamente();
        renderizarSubAbas();
        renderizarTabela();
      }
    });
  } else {
    Swal.fire('Atenção', 'Mínimo de 1 mês.', 'info');
  }
}
async function criarNovoMes() {
  if (currentRole === 'guest') return;
  const { value: nome } = await Swal.fire({
    title: 'Novo Mês',
    input: 'text',
    showCancelButton: true,
  });
  if (nome && nome.trim() !== '' && !dadosAbas[abaAtiva][nome]) {
    dadosAbas[abaAtiva][nome] = [];
    mesAtivo = nome;
    registrarLog('Novo Mês', `Criou o mês ${nome}`);
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
  }
}

async function agruparAbas() {
  if (currentRole === 'guest') return;
  const abasDisponiveis = Object.keys(dadosAbas);
  if (abasDisponiveis.length < 2) return;
  let htmlCheckboxes =
    '<div style="text-align: left; max-height: 200px; overflow-y: auto; padding: 10px; border: 1px solid var(--border-color); border-radius: 5px; background: var(--bg-header);">';
  abasDisponiveis.forEach((aba) => {
    htmlCheckboxes += `<label style="display: block; margin-bottom: 8px; cursor: pointer;"><input type="checkbox" class="swal-aba-checkbox" value="${aba}"> ${aba}</label>`;
  });
  htmlCheckboxes += '</div>';

  const { value: formValues } = await Swal.fire({
    title: 'Agrupar Assuntos',
    html:
      `<input id="swal-input-novo-nome" class="swal2-input" placeholder="Nome do Novo Assunto">` +
      htmlCheckboxes,
    focusConfirm: false,
    showCancelButton: true,
    preConfirm: () => {
      const selecionados = Array.from(document.querySelectorAll('.swal-aba-checkbox:checked')).map(
        (cb) => cb.value,
      );
      const novoNome = document.getElementById('swal-input-novo-nome').value.trim();
      if (selecionados.length === 0 || !novoNome)
        return Swal.showValidationMessage('Preencha os dados!');
      return { selecionados, novoNome };
    },
  });

  if (formValues) {
    const { selecionados, novoNome } = formValues;
    if (!dadosAbas[novoNome]) dadosAbas[novoNome] = {};
    selecionados.forEach((abaAntiga) => {
      Object.keys(dadosAbas[abaAntiga]).forEach((mesAntigo) => {
        let novoNomeMes = abaAntiga.replace('LIQ. ', '').replace('LIQ.', '').trim();
        let registosReais = dadosAbas[abaAntiga][mesAntigo].filter((r) => !r.isDummy);
        if (dadosAbas[novoNome][novoNomeMes])
          dadosAbas[novoNome][novoNomeMes] = dadosAbas[novoNome][novoNomeMes].concat(registosReais);
        else dadosAbas[novoNome][novoNomeMes] = registosReais;
      });
      if (abaAntiga !== novoNome) delete dadosAbas[abaAntiga];
    });
    abaAtiva = novoNome;
    mesAtivo = Object.keys(dadosAbas[novoNome])[0];
    registrarLog('Agrupamento', `Agrupou ${selecionados.length} assuntos em ${novoNome}`);
    salvarArquivoAutomaticamente();
    renderizarAbas();
    renderizarSubAbas();
    renderizarTabela();
  }
}

async function moverProcessosLote() {
  if (currentRole === 'guest') return;
  const { value: formValues } = await Swal.fire({
    title: 'Transferência em Lote',
    width: '500px',
    html: `
        <div style="text-align: left; font-size: 14px; overflow: hidden; padding: 5px;">
            <select id="swal-move-tipo" class="swal2-select" style="width:100%; box-sizing:border-box; margin:0 0 10px 0;">
                <option value="empresa">Empresa igual a</option><option value="processo">Processo igual a</option><option value="empenho">Empenho igual a</option><option value="elemento">Elemento igual a</option>
            </select>
            <input id="swal-move-valor" class="swal2-input" placeholder="Valor exato..." style="width:100%; box-sizing:border-box; margin:0 0 20px 0;">
            <p style="margin-bottom:5px; font-weight:bold; color:var(--text-main);">Mover para:</p>
            <select id="swal-move-aba" class="swal2-select" style="width:100%; box-sizing:border-box; margin:0 0 10px 0;">${Object.keys(
              dadosAbas,
            )
              .map((a) => `<option value="${a}">${a}</option>`)
              .join('')}</select>
            <input id="swal-move-mes" class="swal2-input" placeholder="Mês Destino" style="width:100%; box-sizing:border-box; margin:0;">
        </div>`,
    focusConfirm: false,
    showCancelButton: true,
    preConfirm: () => {
      const tipo = document.getElementById('swal-move-tipo').value;
      const valor = document.getElementById('swal-move-valor').value.trim();
      const abaDestino = document.getElementById('swal-move-aba').value;
      const mesDestino = document.getElementById('swal-move-mes').value.trim();
      if (!valor || !mesDestino) return Swal.showValidationMessage('Preencha tudo!');
      return { tipo, valor, abaDestino, mesDestino };
    },
  });

  if (formValues) {
    const { tipo, valor, abaDestino, mesDestino } = formValues;
    let processosMovidos = [];
    let processosRestantes = [];
    dadosAbas[abaAtiva][mesAtivo].forEach((reg) => {
      if (reg.isDummy) return;
      if (reg[tipo] && reg[tipo].toString().toLowerCase() === valor.toLowerCase())
        processosMovidos.push(reg);
      else processosRestantes.push(reg);
    });
    if (processosMovidos.length === 0)
      return Swal.fire('Nenhum encontrado', `Nada coincide com a busca.`, 'info');
    if (!dadosAbas[abaDestino][mesDestino]) dadosAbas[abaDestino][mesDestino] = [];
    dadosAbas[abaDestino][mesDestino] = dadosAbas[abaDestino][mesDestino].concat(processosMovidos);
    dadosAbas[abaAtiva][mesAtivo] = processosRestantes;
    registrarLog(
      'Transferência Lote',
      `Moveu ${processosMovidos.length} registos (${tipo}:${valor}) para ${abaDestino}>${mesDestino}`,
    );
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
  }
}

function atualizarSelecao(checkbox, id) {
  if (currentRole === 'guest') return;
  if (checkbox.checked) IDsSelecionados.add(id);
  else IDsSelecionados.delete(id);
  verificarBarraLote();
}
function toggleSelecionarTodos(checkboxCentral) {
  if (currentRole === 'guest') return;
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
  if (currentRole === 'guest') return;
  Swal.fire({
    title: `Excluir ${IDsSelecionados.size} processos?`,
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#d33',
  }).then((result) => {
    if (result.isConfirmed) {
      dadosAbas[abaAtiva][mesAtivo] = dadosAbas[abaAtiva][mesAtivo].filter(
        (r) => !IDsSelecionados.has(r.id),
      );
      registrarLog('Exclusão Lote', `Apagou ${IDsSelecionados.size} processos`);
      IDsSelecionados.clear();
      verificarBarraLote();
      salvarArquivoAutomaticamente();
      renderizarTabela();
    }
  });
}

async function moverSelecionados() {
  if (currentRole === 'guest') return;
  let abasOptions = Object.keys(dadosAbas)
    .map((a) => `<option value="${a}">${a}</option>`)
    .join('');
  const { value: formValues } = await Swal.fire({
    title: `Mover ${IDsSelecionados.size} Processos`,
    html: `
        <div style="text-align: left; font-size: 14px;">
            <label>Aba de Destino:</label><select id="swal-move-sel-aba" class="swal2-select" style="width:100%;" onchange="window.atualizarMesesDestinoSel()"><option value="">Selecione...</option>${abasOptions}</select>
            <label>Mês de Destino (ou digite novo):</label><select id="swal-move-sel-mes" class="swal2-select" style="width:100%;"><option value="">Selecione Acima</option></select>
            <input id="swal-move-sel-mes-novo" class="swal2-input" placeholder="Novo Mês..." style="width:100%;">
        </div>`,
    didOpen: () => {
      window.atualizarMesesDestinoSel = () => {
        const aba = document.getElementById('swal-move-sel-aba').value;
        const selMes = document.getElementById('swal-move-sel-mes');
        selMes.innerHTML = '<option value="">Selecione...</option>';
        if (aba && dadosAbas[aba]) {
          Object.keys(dadosAbas[aba]).forEach((m) => {
            selMes.innerHTML += `<option value="${m}">${m}</option>`;
          });
        }
      };
    },
    focusConfirm: false,
    showCancelButton: true,
    preConfirm: () => {
      const abaDestino = document.getElementById('swal-move-sel-aba').value;
      let mesDestino = document.getElementById('swal-move-sel-mes').value;
      const mesNovo = document.getElementById('swal-move-sel-mes-novo').value.trim();
      if (!abaDestino) return Swal.showValidationMessage('Selecione aba!');
      if (!mesDestino && !mesNovo) return Swal.showValidationMessage('Selecione mês!');
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
    registrarLog(
      'Transferência',
      `Moveu ${processosMovidos.length} processos para ${abaDestino} > ${mesDestino}`,
    );
    IDsSelecionados.clear();
    verificarBarraLote();
    salvarArquivoAutomaticamente();
    renderizarSubAbas();
    renderizarTabela();
  }
}

// ==========================================
// AUTOCOMPLETAR E ORDENAÇÃO
// ==========================================
function atualizarAutocompletarAba() {
  let procs = new Set(),
    emps = new Set(),
    elems = new Set(),
    emps2 = new Set();
  if (dadosAbas[abaAtiva]) {
    Object.values(dadosAbas[abaAtiva]).forEach((mes) => {
      mes.forEach((reg) => {
        if (reg.isDummy) return;
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
  if (currentRole === 'guest') return;
  const processo = document.getElementById('processo').value.trim();
  const empresa = document.getElementById('empresa').value.trim();
  const elemento = document.getElementById('elemento').value.trim();
  const empenho = document.getElementById('empenho').value.trim();
  const liquidacao = document.getElementById('liquidacao').value.trim();
  const status = document.getElementById('status_pagamento').value;
  const op = document.getElementById('op').value.trim();

  if (!empresa) return Swal.fire('Obrigatório', 'A Empresa é obrigatória!', 'warning');
  let duplicado = null;
  for (let aba in dadosAbas) {
    for (let mes in dadosAbas[aba]) {
      let achou = dadosAbas[aba][mes].find(
        (r) =>
          !r.isDummy &&
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
  if (duplicado)
    return Swal.fire({
      icon: 'error',
      title: 'Duplicado!',
      html: `Lançamento já existe em:<br><b>Aba:</b> ${duplicado.aba}<br><b>Mês:</b> ${duplicado.mes}`,
    });

  const novoReg = { id: Date.now(), processo, empresa, elemento, empenho, liquidacao, status, op };
  if (!dadosAbas[abaAtiva][mesAtivo]) dadosAbas[abaAtiva][mesAtivo] = [];
  dadosAbas[abaAtiva][mesAtivo].push(novoReg);

  registrarLog('Adição', `Processo ${processo || 'Sem Nro'} - Empresa: ${empresa}`);
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
  if (currentRole === 'guest') return;
  linhaEmEdicao = id;
  renderizarTabela();
}
window.ativarEdicaoEFocus = function (id, campo) {
  if (currentRole === 'guest') return;
  ativarEdicaoInline(id);
  setTimeout(() => {
    const el = document.getElementById(`edit-${campo}-${id}`);
    if (el) el.focus();
  }, 50);
};

window.atualizarStatusDireto = function (id, novoStatus) {
  if (currentRole === 'guest') return;
  const index = dadosAbas[abaAtiva][mesAtivo].findIndex((r) => r.id === id);
  if (index !== -1) {
    let nomeEmp = dadosAbas[abaAtiva][mesAtivo][index].empresa;
    dadosAbas[abaAtiva][mesAtivo][index].status = novoStatus;
    registrarLog('Edição Rápida', `Alterou estado para ${novoStatus} da empresa ${nomeEmp}`);
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

    registrarLog(
      'Edição',
      `Editou os dados da empresa ${dadosAbas[abaAtiva][mesAtivo][index].empresa}`,
    );
    salvarArquivoAutomaticamente();
  }
  linhaEmEdicao = null;
  renderizarTabela();
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
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo]
      ? dadosAbas[abaAtiva][mesAtivo].filter((r) => !r.isDummy)
      : [];
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
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo]
      ? [...dadosAbas[abaAtiva][mesAtivo]].filter((r) => !r.isDummy)
      : [];

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
    document.getElementById('barra-progresso').style.width =
      `${total > 0 ? (pagasFull / total) * 100 : 0}%`;
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
        <td class="coluna-acao"></td>
        <td><input type="text" id="edit-processo-${reg.id}" class="input-inline" value="${reg.processo}" oninput="mascaraProcesso(event)"></td>
        <td><input type="text" id="edit-empresa-${reg.id}" class="input-inline" value="${reg.empresa}"></td>
        <td><input type="text" id="edit-elemento-${reg.id}" class="input-inline" value="${reg.elemento}"></td>
        <td><input type="text" id="edit-empenho-${reg.id}" class="input-inline" value="${reg.empenho}"></td>
        <td><input type="text" id="edit-liquidacao-${reg.id}" class="input-inline" value="${reg.liquidacao}"></td>
        <td><select id="edit-status-${reg.id}" class="input-inline" style="padding: 5px;"><option value="Aguardando Pagamento" ${reg.status === 'Aguardando Pagamento' ? 'selected' : ''}>Aguardando Pagamento</option><option value="Pago" ${reg.status === 'Pago' ? 'selected' : ''}>Pago</option></select></td>
        <td><input type="text" id="edit-op-${reg.id}" class="input-inline" value="${reg.op}"></td>
        <td class="coluna-acao" style="white-space: nowrap;"><button class="btn-edit" onclick="salvarEdicaoInline(${reg.id})"><i class="fa-solid fa-check"></i></button><button class="btn-delete" onclick="cancelarEdicaoInline()"><i class="fa-solid fa-xmark"></i></button></td>
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
        <td class="coluna-acao" style="text-align: center;"><input type="checkbox" class="chk-linha" value="${reg.id}" ${isChecked} onchange="atualizarSelecao(this, ${reg.id})" style="cursor:pointer;"></td>
        <td ondblclick="ativarEdicaoEFocus(${reg.id}, 'processo')" style="cursor:pointer;" title="Duplo clique para editar">${reg.processo}</td>
        <td ondblclick="ativarEdicaoEFocus(${reg.id}, 'empresa')" style="cursor:pointer;" title="Duplo clique para editar">${reg.empresa}</td>
        <td ondblclick="ativarEdicaoEFocus(${reg.id}, 'elemento')" style="cursor:pointer;" title="Duplo clique para editar">${reg.elemento}</td>
        <td class="${classEmpenhoLiq}" ondblclick="ativarEdicaoEFocus(${reg.id}, 'empenho')" style="cursor:pointer;" title="Duplo clique para editar">${reg.empenho}</td>
        <td class="${classEmpenhoLiq}" ondblclick="ativarEdicaoEFocus(${reg.id}, 'liquidacao')" style="cursor:pointer;" title="Duplo clique para editar">${reg.liquidacao}</td>
        <td class="${classOP}"><select class="status-dropdown" onchange="atualizarStatusDireto(${reg.id}, this.value)"><option value="Aguardando Pagamento" ${reg.status === 'Aguardando Pagamento' ? 'selected' : ''}>Aguardando Pagamento</option><option value="Pago" ${reg.status === 'Pago' ? 'selected' : ''}>Pago</option></select></td>
        <td class="${classOP}" ondblclick="ativarEdicaoEFocus(${reg.id}, 'op')" style="cursor:pointer;" title="Duplo clique para editar">${reg.op}</td>
        <td class="coluna-acao" style="white-space: nowrap;"><button class="btn-edit" onclick="ativarEdicaoInline(${reg.id})"><i class="fa-solid fa-pen"></i></button><button class="btn-delete" onclick="apagarLinha(${reg.id})"><i class="fa-solid fa-trash"></i></button></td>
      `;
    if (isChecked) tr.style.backgroundColor = 'var(--bg-lote)';
  }
  return tr;
}

function apagarLinha(id) {
  if (currentRole === 'guest') return;
  Swal.fire({
    title: 'Excluir?',
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#d33',
    confirmButtonText: 'Excluir',
  }).then((result) => {
    if (result.isConfirmed) {
      let empresaNome = dadosAbas[abaAtiva][mesAtivo].find((r) => r.id === id).empresa;
      dadosAbas[abaAtiva][mesAtivo] = dadosAbas[abaAtiva][mesAtivo].filter((r) => r.id !== id);
      IDsSelecionados.delete(id);
      verificarBarraLote();
      registrarLog('Exclusão', `Apagou o registo da empresa ${empresaNome}`);
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
  if (currentRole === 'guest') return;
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    Swal.fire({
      title: 'Importar Planilha?',
      text: `Serão criados ${workbook.SheetNames.length} novos Assuntos.`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'Sim, importar!',
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
            const statusPag = row[6] ? 'Pago' : 'Aguardando Pagamento';
            let procFormatado = proc;
            if (procFormatado && !procFormatado.toUpperCase().startsWith('BJI-'))
              procFormatado = 'BJI-' + procFormatado;
            dadosAbas[sheetName]['Geral'].push({
              id: Date.now() + Math.floor(Math.random() * 100000),
              processo: procFormatado,
              empresa: emp,
              elemento: row[3] ? String(row[3]).trim() : '',
              empenho: row[4] ? String(row[4]).trim() : '',
              liquidacao: row[5] ? String(row[5]).trim() : '',
              status: statusPag,
              op: row[6] ? String(row[6]).trim() : '',
            });
            totalImportados++;
          }
        });
        abaAtiva = workbook.SheetNames[0];
        mesAtivo = 'Geral';
        registrarLog('Importação', `Importou ${totalImportados} registos de Excel`);
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

// Dashboards e Relatorios
function abrirDashboard() {
  document.getElementById('modal-dashboard').style.display = 'flex';
  atualizarGraficos();
}
function fecharDashboard() {
  document.getElementById('modal-dashboard').style.display = 'none';
}
function atualizarGraficos() {
  const registros =
    dadosAbas[abaAtiva] && dadosAbas[abaAtiva][mesAtivo]
      ? dadosAbas[abaAtiva][mesAtivo].filter((r) => !r.isDummy)
      : [];
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

function abrirRelatorioPendencias() {
  dadosRelatorioGeral = [];
  for (let aba in dadosAbas) {
    for (let mes in dadosAbas[aba]) {
      dadosAbas[aba][mes].forEach((reg) => {
        if (!reg.isDummy && (reg.status === 'Aguardando Pagamento' || !reg.op)) {
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
    tbody.innerHTML += ` <tr> <td>${reg.processo}</td><td>${reg.empresa}</td><td>${reg.elemento}</td><td>${reg.empenho}</td><td>${reg.liquidacao}</td> <td style="${corStatus}">${reg.status}</td><td>${opVisivel}</td> <td style="font-weight:bold; color:var(--text-muted);">${reg.abaLocal} <i class="fa-solid fa-angle-right"></i> ${reg.mesLocal}</td> <td class="coluna-acao" style="white-space: nowrap;"><button onclick="irParaProcesso('${reg.abaLocal}', '${reg.mesLocal}', ${reg.id})" style="padding:5px 10px; background:#3498db; color:white; border:none; border-radius:4px; cursor:pointer;" title="Ir para o processo"><i class="fa-solid fa-arrow-right"></i></button></td> </tr>`;
  });
}
window.irParaProcesso = function (aba, mes, id) {
  abaAtiva = aba;
  mesAtivo = mes;
  localStorage.setItem('ultimaAba', abaAtiva);
  localStorage.setItem('ultimoMes', mesAtivo);
  fecharRelatorio();
  IDsSelecionados.clear();
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
