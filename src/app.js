// ==========================================
// 1. CONFIGURAÇÃO E SETUP
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 ESCOLA BÍBLICA')
    .addItem('🛠️ Instalar/Restaurar Sistema', 'setupInicialDoSistema')
    .addItem('🔑 Resetar Senha Admin', 'resetarSenhaAdmin') // Nova função de emergência
    .addItem('🎓 Painel Admin', 'abrirPainelAdmin')
    .addToUi();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('Escola Bíblica IEQ Sede Região 1065')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// FUNÇÃO PARA OBTER A LOGO EM BASE64 (CORREÇÃO DA IMAGEM)
function getLogoBase64() {
  try {
    // ID do arquivo no Google Drive (extraído do seu link)
    const fileId = "1UnwLnvFkwA44pusS_bF0BPIYRjK7JQnN";
    
    // URL correta para baixar o arquivo do Google Drive
    const url = `https://drive.google.com/uc?id=${fileId}&export=download`;
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      return "data:" + blob.getContentType() + ";base64," + base64;
    }
    
    // Fallback: retornar uma imagem padrão em base64
    return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAIGNIUk0AAHolAACAgwAA+f8AAIDpAAB1MAAA6mAAADqYAAAXb5JfxUYAAAGASURBVHja7JdNSsNAFIZfQbqQ3oAbN15AQXAn3oCi+6IXEFy4deUJXHgE8QJupOAJxIUgKMGFghsRCl0I/Rl8XUzSJpNJJ5NOoQsfeZk3X76Zd75JIAgC1tm2rPkOCoJAEIQAhOECgKqqQ/wZQEAQhEwDQC6b0b6+gQoQAe4CQFEULMu6KgD9TUA7C3YAOp0OjuNcFUCpFQfYbrfTNO0jfPnwPO9sR/D9z6QAp9NpWq3Wv67s+/65jgDQ6/VQVTX2CObnAK8vz2xtb1Mul2MACwsL6Hrj3wFc1+Xl+QlJknBdNxWYz+fjCDKZTCoAVVUXQqSUzl8wNE2jXq+nChBCpO8DURTjRdT++IwsSyiKMrX9aDT6y4GJ4xqPxziOg2EY0fl3XXfG5W63S7PZxDTN3x8ghJi5PJ1OJwaQJOl3DqTT6VCtVrEsi3w+jyzLDAYDTNOcacJhGPLx/kYmk0mUjRqNRjJ6YDAYoCgKxWKRfD6PJEmYpkmj0Zhpwn6/z9PjA5ubm7G5oVAoJOuDer3OxsYG3W4X27aRZRnbtqlUKjNN2O/3ebi/Q5bl2NxQqVSS9cG0SZKEEAJFUWaan2EY3N3eIMtybG4olUrJADRNm4lof38f27a5vbme++1PT09omoZt2/Obi+M4mKY5U7Qsy0q90fzLAwBJ4QfE2NNEECg8kQAAAABJRU5ErkJggg==";
    
  } catch(e) {
    console.error("Erro ao carregar logo:", e);
    // Fallback seguro
    return "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgdmlld0JveD0iMCAwIDEwMCAxMDAiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CjxyZWN0IHdpZHRoPSIxMDAiIGhlaWdodD0iMTAwIiByeD0iMTUiIGZpbGw9IiMxQTIzN0UiLz4KPHRleHQgeD0iNTAiIHk9IjU1IiBmb250LXNpemU9IjE0IiBmb250LWZhbWlseT0iQXJpYWwiIGZpbGw9IndoaXRlIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBkb21pbmFudC1iYXNlbGluZT0ibWlkZGxlIj5FQjwvdGV4dD4KPC9zdmc+";
  }
}

// FUNÇÃO PARA CRIAR AS ABAS CORRETAS AUTOMATICAMENTE
function setupInicialDoSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const estrutura = {
    'USUARIOS': ['ID', 'Nome', 'Email', 'SenhaHash', 'Tipo', 'Token', 'TokenExpira', 'Status'],
    'ALUNOS': ['ID', 'Data Cadastro', 'Nome Completo', 'Data Nascimento', 'Email', 'Telefone', 'Foto URL', 'Célula', 'Ministério', 'Tipo Escola', 'Batizado', 'Data Batismo', 'Igreja Batismo', 'Interesse Servir', 'Observações', 'QR Code', 'User ID', 'Status'],
    'PROFESSORES': ['ID', 'Nome', 'Email', 'Telefone', 'Foto URL', 'Ministerio', 'Escola', 'Biografia', 'Status'],
    'AULAS': ['ID', 'Data', 'Horário', 'Disciplina', 'Tema', 'Descrição', 'Professor ID', 'Professor Nome', 'Tipo Escola', 'Status'],
    'MATERIAIS': ['ID', 'Titulo', 'Descricao', 'Tipo', 'Disciplina', 'ArquivoURL', 'Tamanho', 'Autor', 'DataUpload', 'Status'],
    'FREQUENCIAS': ['ID', 'Data', 'Aluno ID', 'Aluno Nome', 'Presente', 'Horário', 'Localizacao', 'Foto URL', 'Dispositivo', 'Aula ID', 'Escola'],
    'NOTIFICACOES': ['ID', 'UsuarioID', 'Titulo', 'Mensagem', 'Tipo', 'Lida', 'DataCriacao', 'DataLeitura', 'AcaoURL'],
    'CONFIG': ['Chave', 'Valor'],
    'CELULAS': ['ID', 'Nome', 'Líder', 'Dia', 'Horário', 'Endereço'],
    'ESCOLAS': ['ID', 'Nome', 'Descrição', 'Dias']
  };

  for (let aba in estrutura) {
    let sheet = ss.getSheetByName(aba);
    if (!sheet) {
      sheet = ss.insertSheet(aba);
      sheet.appendRow(estrutura[aba]);
      sheet.getRange(1, 1, 1, estrutura[aba].length).setFontWeight('bold').setBackground('#E0E0E0');
      sheet.setFrozenRows(1);
    }
  }
  
  // Criar admin padrão se não existir
  const sheetUsers = ss.getSheetByName('USUARIOS');
  if (sheetUsers.getLastRow() === 1) {
    const adminHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'admin123'));
    sheetUsers.appendRow(['ADM001', 'Administrador', 'admin@igreja.com', adminHash, 'admin', '', '', 'ativo']);
  }
  
  // Criar escolas padrão se não existir
  const sheetEscolas = ss.getSheetByName('ESCOLAS');
  if (sheetEscolas.getLastRow() === 1) {
    sheetEscolas.appendRow(['ESCO01', 'Escola de Discípulos', 'Formação de novos convertidos e discipulado básico', 'Segundas e Domingos']);
    sheetEscolas.appendRow(['ESCO02', 'Escola de Líderes', 'Treinamento para liderança e ministérios', 'Segundas e Domingos']);
    sheetEscolas.appendRow(['ESCO03', 'Escola Bíblica Continuada', 'Estudo aprofundado da Bíblia e teologia', 'Sextas e Sábados']);
  }
  
  SpreadsheetApp.getUi().alert('Sistema verificado com sucesso!\n\nLogin Admin:\nEmail: admin@igreja.com\nSenha: admin123');
}

// FUNÇÃO PARA RESETAR A SENHA DO ADMIN (CASO O LOGIN FALHE)
function resetarSenhaAdmin() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USUARIOS');
  const data = sheet.getDataRange().getValues();
  let found = false;
  
  // Hash da senha "admin123"
  const novaHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'admin123'));
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'admin@igreja.com') { // Coluna Email
      sheet.getRange(i + 1, 4).setValue(novaHash); // Coluna SenhaHash
      sheet.getRange(i + 1, 8).setValue('ativo'); // Garantir que está ativo
      found = true;
      break;
    }
  }
  
  if (!found) {
    sheet.appendRow(['ADM001', 'Administrador', 'admin@igreja.com', novaHash, 'admin', '', '', 'ativo']);
    SpreadsheetApp.getUi().alert('Usuário admin recriado.\nEmail: admin@igreja.com\nSenha: admin123');
  } else {
    SpreadsheetApp.getUi().alert('Senha do admin resetada para: admin123');
  }
}

// ==========================================
// 2. SISTEMA DE AUTENTICAÇÃO
// ==========================================

function loginUsuario(email, senha) {
  try {
    // Normalizar entrada (remover espaços extras)
    email = email.toString().trim();
    senha = senha.toString(); 
    
    console.log("Tentativa de login: " + email);

    const usuarios = getDadosDaPlanilha('USUARIOS');
    // Busca exata
    const usuario = usuarios.find(u => u.Email && u.Email.toString().trim() === email);
    
    if (!usuario) {
      console.log("Email não encontrado.");
      return { success: false, message: 'Usuário não encontrado!' };
    }
    
    const senhaHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha));
    
    if (senhaHash !== usuario.SenhaHash) {
      console.log("Senha incorreta.");
      return { success: false, message: 'Senha incorreta!' };
    }
    
    if (usuario.Status && usuario.Status.toLowerCase() !== 'ativo') {
      return { success: false, message: 'Usuário inativo!' };
    }
    
    const token = Utilities.getUuid();
    
    // Atualizar token na planilha
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USUARIOS');
    const data = sheet.getDataRange().getValues();
    
    // Encontrar a linha pelo ID para ser mais seguro
    let rowIndex = -1;
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == usuario.ID) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 6).setValue(token);
      sheet.getRange(rowIndex, 7).setValue(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000));
    }

    return {
      success: true,
      token: token,
      usuario: { 
        id: usuario.ID, 
        nome: usuario.Nome, 
        email: usuario.Email, 
        tipo: usuario.Tipo 
      }
    };
  } catch(e) {
    console.error('Erro no login:', e);
    return { success: false, message: 'Erro: ' + e.message };
  }
}

/* --- FUNÇÃO PRINCIPAL DE REGISTRO (COM UPLOAD DE FOTO) --- */
function registrarUsuario(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('USUARIOS');
    
    // Verifica email duplicado
    const usuarios = getDadosDaPlanilha('USUARIOS');
    const emailExistente = usuarios.find(u => String(u.Email).trim().toLowerCase() === String(dados.email).trim().toLowerCase());
    
    if (emailExistente) return { success: false, message: 'Email já cadastrado!' };
    
    // Gera IDs e Hash
    const prefixo = dados.tipo === 'professor' ? 'PROF' : dados.tipo === 'admin' ? 'ADM' : 'ALU';
    const userId = prefixo + Utilities.getUuid().slice(0, 8).toUpperCase();
    const senhaHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, dados.senha));
    
    // --- LÓGICA DA FOTO ---
    let fotoUrlFinal = '';
    // Se veio foto base64 do formulário, faz upload
    if (dados.fotoBase64 && dados.fotoBase64.length > 10) {
       // Usa a função auxiliar de upload (certifique-se que ela está no seu código)
       fotoUrlFinal = uploadImagemParaDrive(dados.fotoBase64, "perfil_" + userId + ".jpg");
    }
    dados.fotoURL = fotoUrlFinal; // Guarda a URL gerada
    // ---------------------

    // Salva na aba USUARIOS (Login)
    sheet.appendRow([userId, dados.nome, dados.email, senhaHash, dados.tipo || 'aluno', '', '', 'ativo']);
    
    // Distribui para a aba específica
    if (dados.tipo === 'aluno') {
      cadastrarAlunoAutomatico(userId, dados);
    } else if (dados.tipo === 'professor') {
      cadastrarProfessorAutomatico(userId, dados);
    }
    
    return { success: true, message: 'Cadastro realizado!', userId: userId };
  } catch(e) {
    console.error('Erro no registro:', e);
    return { success: false, message: 'Erro: ' + e.message };
  }
}

/* --- FUNÇÃO DE CADASTRO NA ABA ALUNOS (ORDEM CORRIGIDA) --- */
function cadastrarAlunoAutomatico(userId, dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALUNOS');
    
    const alunoId = 'ALU' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const dataCadastro = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    // ORDEM EXATA DAS COLUNAS:
    // 1.ID | 2.Cadastro | 3.Nome | 4.Nasc | 5.Email | 6.Tel | 7.Foto | 8.Célula | 9.Ministério(Atual) | 
    // 10.Escola | 11.Batizado | 12.DtBatismo | 13.IgrejaBatismo | 14.INTERESSE SERVIR | 15.Obs | 16.QR | 17.UserID | 18.Status

    const linha = [
      alunoId,                        // 1. ID
      dataCadastro,                   // 2. Data Cadastro
      dados.nome,                     // 3. Nome Completo
      dados.dataNascimento || '',     // 4. Data nascimento
      dados.email,                    // 5. Email
      dados.telefone || '',           // 6. Telefone
      dados.fotoURL || '',            // 7. Foto URL
      dados.celula || '-',            // 8. Célula (Vem do Dropdown Dinâmico)
      '-',                            // 9. Ministério ATUAL (Vazio no cadastro)
      dados.escola || '-',            // 10. Tipo Escola
      dados.batizado || 'Não',        // 11. Batizado
      '',                             // 12. Data Batismo
      '',                             // 13. Igreja Batismo
      dados.ministerio || '-',        // 14. INTERESSE SERVIR (Aqui entra o dado do form!)
      '',                             // 15. Observações
      '',                             // 16. Qr code
      userId,                         // 17. User ID
      'ativo'                         // 18. Status
    ];

    sheet.appendRow(linha);
    return alunoId;
  } catch(e) {
    console.error('Erro ao cadastrar aluno automático:', e);
    return null;
  }
}
/* --- FUNÇÃO CORRIGIDA: CADASTRAR PROFESSOR (ORDEM EXATA) --- */
function cadastrarProfessorAutomatico(userId, dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROFESSORES');
    const profId = userId; 
    
    // MONTAGEM DA LINHA NA ORDEM PEDIDA:
    // ['ID', 'Nome', 'Email', 'Telefone', 'Foto URL', 'Ministerio', 'Escola', 'Biografia', 'Status']
    
    const linha = [
      profId,                       // 1. ID
      dados.nome,                   // 2. Nome
      dados.email,                  // 3. Email
      dados.telefone || '',         // 4. Telefone
      dados.fotoURL || '',          // 5. Foto URL (Agora na posição correta)
      dados.ministerio || 'Geral',  // 6. Ministerio
      dados.escola || '',           // 7. Escola (ADICIONADO: Agora salva a escola!)
      dados.biografia || '',        // 8. Biografia
      'ativo'                       // 9. Status
    ];

    sheet.appendRow(linha);
    return profId;
  } catch(e) {
    console.error('Erro ao cadastrar professor automático:', e);
    return null;
  }
}

function verificarToken(token) {
  try {
    const usuarios = getDadosDaPlanilha('USUARIOS');
    const usuario = usuarios.find(u => u.Token === token);
    if (!usuario) return { valido: false };
    return { 
      valido: true, 
      usuario: { 
        id: usuario.ID, 
        nome: usuario.Nome, 
        email: usuario.Email, 
        tipo: usuario.Tipo 
      } 
    };
  } catch(e) {
    console.error('Erro ao verificar token:', e);
    return { valido: false };
  }
}

// ==========================================
// 3. FUNÇÕES DE DADOS (GETTERS)
// ==========================================

function getDadosDaPlanilha(nomeAba) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);
    
    return rows.map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let value = row[i];
        
        // --- CORREÇÃO DA DATA VS HORA ---
        if (value instanceof Date) {
          const ano = value.getFullYear();
          
          // Se o ano for 1899, o Google Sheets está dizendo que é APENAS HORA
          if (ano === 1899) {
             const h = String(value.getHours()).padStart(2, '0');
             const m = String(value.getMinutes()).padStart(2, '0');
             value = `${h}:${m}`; // Retorna "19:30"
          } 
          // Se for ano normal, é data (aniversário, data da aula)
          else {
             const mes = String(value.getMonth() + 1).padStart(2, '0');
             const dia = String(value.getDate()).padStart(2, '0');
             value = `${ano}-${mes}-${dia}`; // Retorna "2026-02-02"
          }
        } 
        // Se já vier como texto, limpa quebras de linha
        else if (typeof value === 'string') {
           value = value.trim();
        }
        
        obj[header] = value;
      });
      return obj;
    });
  } catch(e) {
    console.error('Erro ao buscar dados:', e);
    return [];
  }
}

function getAlunos() { return getDadosDaPlanilha('ALUNOS'); }
function getProfessores() { return getDadosDaPlanilha('PROFESSORES'); }
function getAulas() { return getDadosDaPlanilha('AULAS'); }
function getMateriais() { return getDadosDaPlanilha('MATERIAIS'); }
function getEscolas() { return getDadosDaPlanilha('ESCOLAS'); }

function getAlunoPorUserId(userId) {
  const alunos = getDadosDaPlanilha('ALUNOS');
  
  // Busca flexível (ignora maiúsculas e espaços)
  const aluno = alunos.find(a => 
    String(a['User ID'] || '').toUpperCase().trim() === String(userId).toUpperCase().trim() ||
    String(a['ID'] || '').toUpperCase().trim() === String(userId).toUpperCase().trim()
  );
  
  if (aluno) {
    // --- CORREÇÃO DA DATA DE NASCIMENTO (Fuso Horário) ---
    let dataNascFormatada = '';
    const rawData = aluno['Data Nascimento'];
    
    if (rawData instanceof Date) {
      // Se a planilha entregar como Objeto Data, forçamos o formato GMT para não subtrair o dia
      dataNascFormatada = Utilities.formatDate(rawData, "GMT", "yyyy-MM-dd");
    } else if (rawData && String(rawData).length >= 10) {
      // Se vier como texto (ex: 1999-05-10), pegamos apenas os 10 primeiros caracteres
      dataNascFormatada = String(rawData).substring(0, 10);
    }

    // --- CORREÇÃO DO MINISTÉRIO ---
    // Tenta pegar o Nome (Coluna J), se não tiver, pega o Código (Coluna I)
    let nomeMinisterio = aluno['ID Ministério'];
    if (!nomeMinisterio || nomeMinisterio === "" || nomeMinisterio === "#N/A") {
        nomeMinisterio = aluno['Ministério']; 
    }

    return {
      'ID': aluno['ID'] || aluno['Matricula'] || '',
      'User ID': aluno['User ID'],
      'Nome Completo': aluno['Nome Completo'] || 'Nome não informado',
      'Email': aluno['Email'] || '',
      'Telefone': aluno['Telefone'] || '',
      
      // Aqui vai a data corrigida
      'Data Nascimento': dataNascFormatada,
      
      'Data Cadastro': aluno['Data Cadastro'] || '',
      'Escola': aluno['Escola'] || aluno['Tipo Escola'] || '-',
      'Célula': aluno['Célula'] || '-',
      'Ministério': nomeMinisterio || '-',
      'Batizado': aluno['Batizado'] || 'Não',
      'Foto URL': aluno['Foto URL'] || '', 
      'Observações': aluno['Observações'] || '',
      'Status': aluno['Status'] || 'ativo'
    };
  }
  return null;
}

function getProfessorPorUserId(userId) {
  const profs = getProfessores();
  const professor = profs.find(p => p.ID === userId);
  
  if (professor) {
    professor['Foto URL'] = professor['Foto URL'] || '';
    professor['Ministerio'] = professor['Ministerio'] || '';
    professor['Escola'] = professor['Escola'] || '';
    professor['Biografia'] = professor['Biografia'] || '';
    professor['Telefone'] = professor['Telefone'] || '';
  }
  
  return professor || null;
}

function getProximasAulas(escola = null) {
  let aulas = getAulas();
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); 
  
  aulas = aulas.filter(a => {
    if (!a.Data) return false;
    try {
      const dataAula = new Date(a.Data);
      // Ajuste de fuso horário simples para comparação de datas
      dataAula.setHours(dataAula.getHours() + 3); 
      return dataAula >= hoje;
    } catch(e) {
      return false;
    }
  });
  
  if (escola) {
    aulas = aulas.filter(a => a['Tipo Escola'] === escola);
  }
  
  return aulas.sort((a,b) => {
    if (!a.Data || !b.Data) return 0;
    return new Date(a.Data) - new Date(b.Data);
  });
}
/* --- FUNÇÃO CORRIGIDA: BUSCAR AULAS DO PROFESSOR (FORMATADA) --- */
function getAulasProfessor(professorId) {
  const aulas = getDadosDaPlanilha('AULAS');
  const professores = getDadosDaPlanilha('PROFESSORES');
  
  // 1. Filtra apenas as aulas deste professor
  let minhasAulas = aulas.filter(a => String(a['Professor ID']) === String(professorId));

  // 2. Ordena por Data (Próximas primeiro)
  minhasAulas.sort((a, b) => {
    if (!a.Data) return 1;
    if (!b.Data) return -1;
    return new Date(a.Data) - new Date(b.Data);
  });

  return minhasAulas.map(aula => {
    // --- A. BUSCA DADOS DO PROFESSOR (FOTO E CARGO) ---
    let fotoProf = '';
    let nomeProf = 'Você';
    let cargoProf = 'Docente';

    // Procura os dados do próprio professor na planilha para garantir que estão atualizados
    const prof = professores.find(p => String(p.ID) === String(professorId));
    
    if (prof) {
       nomeProf = prof.Nome;
       cargoProf = prof.Ministerio || 'Professor';
       
       // Tratamento da Foto
       if (prof['Foto URL'] && prof['Foto URL'].length > 10) {
         let url = prof['Foto URL'];
         try {
           let id = null;
           if (url.indexOf('id=') > -1) id = url.split('id=')[1].split('&')[0];
           else if (url.indexOf('/d/') > -1) id = url.split('/d/')[1].split('/')[0];
           if (id) fotoProf = "https://lh3.googleusercontent.com/d/" + id;
         } catch(e) {}
       }
    } else {
       // Se não achar na planilha de professores, usa o que está na aula
       nomeProf = aula['Professor Nome'] || 'Você';
    }

    // --- B. FORMATAÇÃO DE DATA ---
    let dataF = '-';
    if (aula.Data) {
      try {
        dataF = Utilities.formatDate(new Date(aula.Data), "GMT", "dd/MM/yyyy");
      } catch(e) { dataF = String(aula.Data); }
    }

    // --- C. FORMATAÇÃO DE HORA ---
    let horaF = aula.Horário || '00:00';
    // Limpa se vier data junto (bug do 1899)
    if (String(horaF).length > 5) horaF = String(horaF).substring(0, 5);

    // --- RETORNO COM AS CHAVES QUE O HTML ESPERA ---
    return {
      ...aula,
      'DataFormatada': dataF,      // Correção do "undefined" na data
      'HoraFormatada': horaF,      // Correção do "undefined" na hora
      'FotoProfessor': fotoProf,
      'ProfessorNome': nomeProf,   // Correção do "undefined" no nome
      'CargoProfessor': cargoProf, // Correção do "undefined" no cargo
      'Descricao': aula['Descrição'] || '',
      'DataOriginal': aula.Data
    };
  });
}
function getMateriaisPorEscola(escola) {
  const materiais = getMateriais();
  if (!escola) return materiais;
  return materiais.filter(m => m.Disciplina === escola || m.Disciplina === 'Geral');
}

function getFrequenciaAluno(alunoId) {
  const freqs = getDadosDaPlanilha('FREQUENCIAS');
  return freqs.filter(f => f['Aluno ID'] === alunoId);
}

function getFrequenciaPorEscola(escola) {
  const freqs = getDadosDaPlanilha('FREQUENCIAS');
  if (!escola) return freqs;
  return freqs.filter(f => f.Escola === escola);
}

function getFrequenciaPorEscolaProfessor(professorId) {
  try {
    const professor = getProfessorPorUserId(professorId);
    if (!professor || !professor.Escola) return [];
    
    return getFrequenciaPorEscola(professor.Escola);
  } catch(e) {
    console.error('Erro ao buscar frequência por escola do professor:', e);
    return [];
  }
}

function getEstatisticas() {
  return {
    totalAlunos: getAlunos().length,
    totalProfessores: getProfessores().length,
    totalAulas: getAulas().length,
    totalMateriais: getMateriais().length,
    totalCheckins: getDadosDaPlanilha('FREQUENCIAS').length
  };
}

// ==========================================
// 4. AÇÕES (UPDATES/INSERTS)
// ==========================================

function criarAula(dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AULAS');
    const id = 'AULA' + Date.now();
    
    sheet.appendRow([
      id, 
      dados.data, 
      dados.horario, 
      dados.disciplina, 
      dados.tema, 
      dados.descricao, 
      dados.professorId, 
      dados.professorNome, 
      dados.escola, 
      'agendada'
    ]);
    
    return { success: true, message: 'Aula criada com sucesso!', id: id };
  } catch(e) {
    console.error('Erro ao criar aula:', e);
    return { success: false, message: e.message };
  }
}

function adicionarMaterial(dados) {
  try {
    console.log('🚀 INICIANDO UPLOAD (MODO HÍBRIDO)');
    console.log('📄 Arquivo:', dados.nomeArquivo);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('MATERIAIS') || ss.getSheets()[0];
    let arquivoURL = '';
    
    // --- 1. PROCESSAMENTO DO ARQUIVO ---
    if (dados.arquivoBase64 && dados.arquivoBase64.length > 0) {
      try {
        // 1. Limpeza da string Base64 (Remove cabeçalhos e espaços)
        let base64Limpo = dados.arquivoBase64;
        if (base64Limpo.indexOf(',') > -1) {
             base64Limpo = base64Limpo.split(',')[1];
        }

        console.log('🔄 Decodificando bytes...');
        const decodedBytes = Utilities.base64Decode(base64Limpo);
        
        console.log('💾 Preparando blob...');
        const blob = Utilities.newBlob(decodedBytes, dados.tipoMime || 'application/pdf', dados.nomeArquivo);
        
        let fileId;
        
        // 2. TENTA UPLOAD VIA API (Para suportar arquivos grandes > 5MB)
        try {
          const fileResource = {
            title: dados.nomeArquivo,
            mimeType: dados.tipoMime || 'application/pdf'
          };
          // Criação via API (Serviço Avançado)
          const fileApi = Drive.Files.insert(fileResource, blob);
          fileId = fileApi.id;
          console.log('✅ Arquivo criado via API. ID:', fileId);
          
        } catch (eApi) {
          console.log('⚠️ API falhou ou não ativada. Usando método simples...');
          // Fallback: Criação via DriveApp (funciona para arquivos menores)
          const fileApp = DriveApp.createFile(blob);
          fileId = fileApp.getId();
        }

        // 3. CONFIGURAÇÃO FINAL VIA DriveApp (Para garantir o link correto)
        // Usamos o DriveApp aqui porque ele gera o link mais compatível com Apps móveis
        const fileFinal = DriveApp.getFileById(fileId);
        
        // Tenta liberar acesso público ("Qualquer pessoa com o link")
        try {
          fileFinal.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (ePerm) {
          console.log('⚠️ Aviso: Permissão pública bloqueada pelo domínio da conta.');
        }
        
        // Pega o URL nativo (Geralmente termina em ?usp=drivesdk)
        // Esse é o link que funcionava antes para visualização
        arquivoURL = fileFinal.getUrl();

        console.log('✅ Link gerado:', arquivoURL);
        
      } catch (uploadError) {
        console.error('❌ ERRO NO UPLOAD:', uploadError.toString());
        arquivoURL = 'ERRO_UPLOAD: ' + uploadError.message;
      }
    } else {
      arquivoURL = 'SEM_ARQUIVO';
    }
    
    // --- 2. REGISTRO NA PLANILHA ---
    const materialId = 'MAT' + Utilities.getUuid().slice(0, 8).toUpperCase();
    
    sheet.appendRow([
      materialId,
      dados.titulo || 'Sem título',
      dados.descricao || '',
      'pdf',
      dados.disciplina || 'Geral',
      arquivoURL,
      dados.tamanhoFormatado || '0 KB',
      dados.autor || 'Desconhecido',
      new Date(),
      'ativo'
    ]);
    
    return {
      success: true,
      id: materialId,
      url: arquivoURL,
      message: arquivoURL.includes('ERRO') ? 'Erro no envio do arquivo.' : 'Material salvo com sucesso!',
      warning: arquivoURL.includes('ERRO')
    };
    
  } catch (finalError) {
    console.error('💥 ERRO FATAL:', finalError.toString());
    return {
      success: false,
      message: 'Erro no servidor: ' + finalError.message
    };
  }
}

function registrarFrequencia(dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FREQUENCIAS');
    const agora = new Date();
    const hojeData = agora.toISOString().split('T')[0]; // Renomeei para hojeData para evitar conflito
    
    // Buscar dados do aluno para pegar a escola
    const alunos = getDadosDaPlanilha('ALUNOS');
    const aluno = alunos.find(a => a['ID'] === dados.alunoId || a['User ID'] === dados.alunoId);
    
    // Buscar aula atual mais próxima para vincular
    const aulas = getDadosDaPlanilha('AULAS');
    
    // Procurar aula do dia atual na mesma escola do aluno
    let aulaId = '';
    let escola = '';
    
    if (aluno) {
      escola = aluno['Tipo Escola'] || aluno['Escola'] || dados.escola || '';
      const aulaHoje = aulas.find(a => {
        if (!a.Data) return false;
        
        try {
          const dataAula = new Date(a.Data);
          const dataAulaFormatada = dataAula.toISOString().split('T')[0];
          const mesmaData = dataAulaFormatada === hojeData;
          const mesmaEscola = a['Tipo Escola'] === escola;
          const statusValido = a.Status === 'agendada' || a.Status === 'realizada';
          
          return mesmaData && mesmaEscola && statusValido;
        } catch(e) {
          console.error('Erro ao processar data da aula:', e, a.Data);
          return false;
        }
      });
      
      if (aulaHoje) {
        aulaId = aulaHoje.ID;
      }
    }
    
    // Verificar duplicidade (agora considerando também a aula)
    const freqs = getDadosDaPlanilha('FREQUENCIAS');
    const jaTem = freqs.some(f => {
      try {
        const dataFreq = f.Data ? new Date(f.Data).toISOString().split('T')[0] : '';
        const mesmoAluno = f['Aluno ID'] === dados.alunoId;
        const mesmaAula = f['Aula ID'] === aulaId;
        const mesmoDia = dataFreq === hojeData;
        
        // Verifica se já tem na mesma aula OU no mesmo dia
        return mesmoAluno && (mesmaAula || (mesmoDia && aulaId === ''));
      } catch(e) {
        console.error('Erro ao verificar duplicidade:', e);
        return false;
      }
    });
    
    if (jaTem) return { 
      success: false, 
      jaRegistrado: true, 
      message: 'Já registrado para esta aula/hoje!' 
    };

    // Processar foto (se houver)
    let fotoURL = '';
    if (dados.fotoData) {
      try {
        let encoded = dados.fotoData;
        if (encoded.includes(',')) {
          encoded = encoded.split(',')[1];
        }
        const blob = Utilities.newBlob(
          Utilities.base64Decode(encoded), 
          'image/jpeg', 
          'checkin_' + dados.alunoId + '_' + Date.now() + '.jpg'
        );
        const file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoURL = file.getUrl();
      } catch(uploadError) {
        console.error('Erro no upload da foto:', uploadError);
      }
    }
    
    // Gerar link do Google Maps
    const mapsLink = `https://www.google.com/maps/search/?api=1&query=${dados.latitude || 0},${dados.longitude || 0}`;
    const registroId = 'FREQ' + Utilities.getUuid().slice(0, 8).toUpperCase();

    // Se não encontrou escola no aluno, usa a enviada nos dados
    if (!escola) {
      escola = dados.escola || '';
    }

    // Registrar na planilha COM AULA ID e ESCOLA
    sheet.appendRow([
      registroId,                    // ID
      agora,                         // Data
      dados.alunoId || '',           // Aluno ID
      dados.alunoNome,               // Aluno Nome
      'Sim',                         // Presente
      agora.toLocaleTimeString('pt-BR'), // Horário
      mapsLink,                      // Localização
      fotoURL,                       // Foto URL
      dados.dispositivo || 'Web App', // Dispositivo
      aulaId,                        // Aula ID (AGORA PREENCHIDO)
      escola                         // Escola (AGORA PREENCHIDO)
    ]);
    
    return { 
      success: true, 
      message: 'Frequência registrada com sucesso!',
      fotoURL: fotoURL, 
      localizacaoURL: mapsLink,
      aulaId: aulaId,
      escola: escola,
      id: registroId
    };
    
  } catch(e) {
    console.error('Erro ao registrar frequência:', e);
    return { success: false, message: e.message };
  }
}
function atualizarAluno(dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ALUNOS');
    const alunos = getAlunos();
    const index = alunos.findIndex(a => a['User ID'] === dados.userId);
    
    if (index === -1) return { success: false, message: 'Aluno não encontrado' };
    const row = index + 2;
    
    if (dados.nome) sheet.getRange(row, 3).setValue(dados.nome);
    if (dados.email) sheet.getRange(row, 5).setValue(dados.email);
    if (dados.telefone) sheet.getRange(row, 6).setValue(dados.telefone);
    if (dados.dataNascimento) sheet.getRange(row, 4).setValue(dados.dataNascimento);
    if (dados.celula) sheet.getRange(row, 8).setValue(dados.celula);
    if (dados.ministerio) sheet.getRange(row, 9).setValue(dados.ministerio);
    if (dados.batizado) sheet.getRange(row, 11).setValue(dados.batizado);
    if (dados.observacoes) sheet.getRange(row, 15).setValue(dados.observacoes);

    if (dados.fotoBase64) {
      try {
        let encoded = dados.fotoBase64;
        if (encoded.includes(',')) {
          encoded = encoded.split(',')[1];
        }
        const blob = Utilities.newBlob(
          Utilities.base64Decode(encoded), 
          'image/jpeg', 
          'aluno_' + dados.userId + '.jpg'
        );
        const file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);// CORREÇÃO: Usar link direto
        const linkDireto = "https://drive.google.com/uc?export=view&id=" + file.getId();
        sheet.getRange(row, 7).setValue(linkDireto);
        } catch(uploadError) {
        console.error('Erro ao atualizar foto:', uploadError);
      }
    }

    if (dados.email) {
      const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USUARIOS');
      const usuarios = getDadosDaPlanilha('USUARIOS');
      const userIndex = usuarios.findIndex(u => u.ID === dados.userId);
      if (userIndex !== -1) {
        userSheet.getRange(userIndex + 2, 3).setValue(dados.email);
        if (dados.nome) userSheet.getRange(userIndex + 2, 2).setValue(dados.nome);
      }
    }

    return { success: true, message: 'Perfil atualizado!' };
  } catch(e) {
    console.error('Erro ao atualizar aluno:', e);
    return { success: false, message: e.message };
  }
}

function atualizarProfessor(dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROFESSORES');
    const profs = getProfessores();
    const index = profs.findIndex(p => p.ID === dados.userId);
    
    if (index === -1) return { success: false, message: 'Professor não encontrado' };
    const row = index + 2;
    
    if (dados.nome) sheet.getRange(row, 2).setValue(dados.nome);
    if (dados.email) sheet.getRange(row, 3).setValue(dados.email);
    if (dados.telefone) sheet.getRange(row, 4).setValue(dados.telefone);
    if (dados.ministerio) sheet.getRange(row, 6).setValue(dados.ministerio);
    if (dados.biografia) sheet.getRange(row, 8).setValue(dados.biografia);
    
    if (dados.fotoBase64) {
      try {
        let encoded = dados.fotoBase64;
        if (encoded.includes(',')) {
          encoded = encoded.split(',')[1];
        }
        const blob = Utilities.newBlob(
          Utilities.base64Decode(encoded), 
          'image/jpeg', 
          'professor_' + dados.userId + '.jpg'
        );
        const file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        sheet.getRange(row, 5).setValue(file.getUrl());
      } catch(uploadError) {
        console.error('Erro ao atualizar foto do professor:', uploadError);
      }
    }
    
    if (dados.email || dados.nome) {
      const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USUARIOS');
      const usuarios = getDadosDaPlanilha('USUARIOS');
      const userIndex = usuarios.findIndex(u => u.ID === dados.userId);
      if (userIndex !== -1) {
        if (dados.nome) userSheet.getRange(userIndex + 2, 2).setValue(dados.nome);
        if (dados.email) userSheet.getRange(userIndex + 2, 3).setValue(dados.email);
      }
    }

    return { success: true, message: 'Perfil atualizado!' };
  } catch(e) {
    console.error('Erro ao atualizar professor:', e);
    return { success: false, message: e.message };
  }
}

function alterarSenhaUsuario(userId, senhaAtual, novaSenha) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USUARIOS');
    const usuarios = getDadosDaPlanilha('USUARIOS');
    const index = usuarios.findIndex(u => u.ID === userId);
    
    if (index === -1) return { success: false, message: 'Usuário não encontrado' };
    
    const usuario = usuarios[index];
    const senhaAtualHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senhaAtual));
    
    if (senhaAtualHash !== usuario.SenhaHash) {
      return { success: false, message: 'Senha atual incorreta!' };
    }
    
    const novaSenhaHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, novaSenha));
    sheet.getRange(index + 2, 4).setValue(novaSenhaHash);
    
    return { success: true, message: 'Senha alterada com sucesso!' };
    
  } catch(e) {
    console.error('Erro ao alterar senha:', e);
    return { success: false, message: 'Erro: ' + e.message };
  }
}

// ==========================================
// 5. BÍBLIA E VERSÍCULOS - API ATUALIZADA
// ==========================================

// ==========================================
// 5. BÍBLIA E VERSÍCULOS - API ATUALIZADA
// ==========================================

function getVersiculoDia() {
  try {
    // Usando a nova API bible-api.com com versão em português (ARA - Almeida Revista e Atualizada)
    const url = 'https://bible-api.com/?random=verse&translation=almeida';
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' }
    });
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      return { 
        sucesso: true, 
        versiculo: { 
          texto: json.text, 
          referencia: json.reference, 
          versao: json.translation_name || 'Almeida', 
          data: new Date().toISOString().split('T')[0] 
        } 
      };
    }
  } catch(e) {
    console.error('Erro ao buscar versículo do dia:', e);
  }
  
  // Fallback em português
  return { 
    sucesso: true, 
    versiculo: { 
      texto: "Tudo posso naquele que me fortalece.", 
      referencia: "Filipenses 4:13", 
      versao: "Almeida Corrigida Fiel", 
      data: new Date().toISOString().split('T')[0] 
    } 
  };
}

function buscarVersiculos(pesquisa) {
  try {
    // Usando a nova API bible-api.com
    const url = `https://bible-api.com/${encodeURIComponent(pesquisa)}?translation=almeida`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' }
    });
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      
      if (json && json.verses && json.verses.length > 0) {
        const resultados = json.verses.map(v => ({
          referencia: `${v.book_name} ${v.chapter}:${v.verse}`,
          texto: v.text,
          versao: json.translation_name || 'Almeida Corrigida Fiel'
        }));
        
        return {
          sucesso: true,
          resultados: resultados.slice(0, 10),
          total: resultados.length,
          pesquisa: pesquisa
        };
      }
    }
    
    return {
      sucesso: false,
      resultados: [],
      total: 0,
      pesquisa: pesquisa,
      mensagem: 'Nenhum versículo encontrado.'
    };
    
  } catch(e) {
    console.error('Erro na busca de versículos:', e);
    return {
      sucesso: false,
      resultados: [],
      total: 0,
      pesquisa: pesquisa,
      mensagem: 'Erro na busca. Tente novamente.'
    };
  }
}

function getLivrosBiblia() {
  const livros = [
    'Gênesis', 'Êxodo', 'Levítico', 'Números', 'Deuteronômio',
    'Josué', 'Juízes', 'Rute', '1 Samuel', '2 Samuel',
    '1 Reis', '2 Reis', '1 Crônicas', '2 Crônicas', 'Esdras',
    'Neemias', 'Ester', 'Jó', 'Salmos', 'Provérbios',
    'Eclesiastes', 'Cânticos', 'Isaías', 'Jeremias', 'Lamentações',
    'Ezequiel', 'Daniel', 'Oséias', 'Joel', 'Amós',
    'Obadias', 'Jonas', 'Miquéias', 'Naum', 'Habacuque',
    'Sofonias', 'Ageu', 'Zacarias', 'Malaquias', 'Mateus',
    'Marcos', 'Lucas', 'João', 'Atos', 'Romanos',
    '1 Coríntios', '2 Coríntios', 'Gálatas', 'Efésios', 'Filipenses',
    'Colossenses', '1 Tessalonicenses', '2 Tessalonicenses', '1 Timóteo',
    '2 Timóteo', 'Tito', 'Filemom', 'Hebreus', 'Tiago',
    '1 Pedro', '2 Pedro', '1 João', '2 João', '3 João',
    'Judas', 'Apocalipse'
  ];
  
  return livros.map(livro => ({ nome: livro }));
}

function getVersiculos(livro, capitulo) {
  try {
    // Usando a nova API bible-api.com
    const url = `https://bible-api.com/${encodeURIComponent(livro)}%20${capitulo}?translation=almeida`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' }
    });
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      
      if (json.verses && json.verses.length > 0) {
        return {
          sucesso: true,
          livro: livro,
          capitulo: capitulo,
          versiculos: json.verses.map(v => ({
            versiculo: v.verse,
            texto: v.text
          })),
          referencia: json.reference,
          versao: json.translation_name || 'Almeida Corrigida Fiel'
        };
      }
    }
    
    return { sucesso: false, mensagem: 'Capítulo não encontrado' };
    
  } catch(e) {
    console.error('Erro ao buscar versículos:', e);
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

// Função para buscar capítulos específicos - usado pela interface web
function getCapituloBiblia(livro, capitulo) {
  try {
    // Formata o livro para a API (remove espaços)
    const livroFormatado = livro.replace(/ /g, '%20');
    const url = `https://bible-api.com/${livroFormatado}%20${capitulo}?translation=almeida`;
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' }
    });
    
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      return {
        success: true,
        data: json
      };
    } else {
      return {
        success: false,
        message: 'Erro ao buscar capítulo'
      };
    }
  } catch(e) {
    console.error('Erro ao buscar capítulo da Bíblia:', e);
    return {
      success: false,
      message: 'Erro: ' + e.message
    };
  }
}

// ==========================================
// 6. FUNÇÕES DE RELATÓRIOS E UTILITÁRIOS
// ==========================================

function gerarRelatorioFrequencia(dados) {
  try {
    let frequencias = getDadosDaPlanilha('FREQUENCIAS');
    
    if (dados.escola) {
      frequencias = frequencias.filter(f => f.Escola === dados.escola);
    }
    
    if (dados.dataInicio && dados.dataFim) {
      const inicio = new Date(dados.dataInicio);
      const fim = new Date(dados.dataFim);
      
      frequencias = frequencias.filter(f => {
        if (!f.Data) return false;
        const data = new Date(f.Data);
        return data >= inicio && data <= fim;
      });
    }
    
    return {
      success: true,
      data: frequencias,
      total: frequencias.length
    };
    
  } catch(e) {
    console.error('Erro ao gerar relatório de frequência:', e);
    return { success: false, message: 'Erro: ' + e.message };
  }
}

function gerarRelatorioAlunos(dados) {
  try {
    let alunos = getAlunos();
    
    if (dados.escola) {
      alunos = alunos.filter(a => a['Tipo Escola'] === dados.escola);
    }
    
    if (dados.status) {
      alunos = alunos.filter(a => a.Status === dados.status);
    }
    
    return {
      success: true,
      data: alunos,
      total: alunos.length
    };
    
  } catch(e) {
    return { success: false, message: 'Erro: ' + e.message };
  }
}

function gerarRelatorioAulas(dados) {
  try {
    let aulas = getAulas();
    
    if (dados.escola) {
      aulas = aulas.filter(a => a['Tipo Escola'] === dados.escola);
    }
    
    if (dados.dataInicio && dados.dataFim) {
      const inicio = new Date(dados.dataInicio);
      const fim = new Date(dados.dataFim);
      
      aulas = aulas.filter(a => {
        if (!a.Data) return false;
        const data = new Date(a.Data);
        return data >= inicio && data <= fim;
      });
    }
    
    return {
      success: true,
      data: aulas,
      total: aulas.length
    };
    
  } catch(e) {
    return { success: false, message: 'Erro: ' + e.message };
  }
}

function gerarRelatorioProfessores(dados) {
  try {
    let professores = getProfessores();
    if (dados.status) {
      professores = professores.filter(p => p.Status === dados.status);
    }
    return { success: true, data: professores, total: professores.length };
  } catch(e) {
    return { success: false, message: 'Erro: ' + e.message };
  }
}

function exportarParaExcel(dados, tipo) {
  try {
    if (!dados || dados.length === 0) return { success: false, message: 'Nenhum dado para exportar' };
    
    const nomePlanilha = `Relatorio_${tipo}_${Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd_HHmmss')}`;
    const novaPlanilha = SpreadsheetApp.create(nomePlanilha);
    const sheet = novaPlanilha.getActiveSheet();
    
    const headers = Object.keys(dados[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const valores = dados.map(item => headers.map(header => item[header] || ''));
    sheet.getRange(2, 1, valores.length, headers.length).setValues(valores);
    
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1A237E').setFontColor('#FFFFFF').setFontWeight('bold');
    
    return {
      success: true,
      url: novaPlanilha.getUrl(),
      nome: nomePlanilha,
      message: 'Relatório exportado com sucesso!'
    };
    
  } catch(e) {
    return { success: false, message: 'Erro: ' + e.message };
  }
}

function getDetalhesAluno(alunoId) {
  try {
    const alunos = getAlunos();
    const aluno = alunos.find(a => a.ID === alunoId);
    
    if (!aluno) return { success: false, message: 'Aluno não encontrado' };
    
    const frequencias = getDadosDaPlanilha('FREQUENCIAS');
    const frequenciasAluno = frequencias.filter(f => f['Aluno ID'] === alunoId);
    
    return {
      success: true,
      aluno: aluno,
      totalFrequencias: frequenciasAluno.length,
      ultimaFrequencia: frequenciasAluno.length > 0 ? frequenciasAluno[frequenciasAluno.length - 1] : null
    };
  } catch(e) {
    return { success: false, message: 'Erro: ' + e.message };
  }
}

function getNotificacoesUsuario(usuarioId, apenasNaoLidas = false) {
  try {
    const notificacoes = getDadosDaPlanilha('NOTIFICACOES');
    let filtradas = notificacoes.filter(n => n.UsuarioID === usuarioId);
    
    if (apenasNaoLidas) {
      filtradas = filtradas.filter(n => !n.Lida || n.Lida === 'FALSE' || n.Lida === false);
    }
    
    return filtradas.slice(-10).reverse();
  } catch(e) {
    return [];
  }
}

function marcarNotificacaoLida(notificacaoId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NOTIFICACOES');
    const notificacoes = getDadosDaPlanilha('NOTIFICACOES');
    const index = notificacoes.findIndex(n => n.ID === notificacaoId);
    
    if (index !== -1) {
      sheet.getRange(index + 2, 6).setValue(true);
      sheet.getRange(index + 2, 8).setValue(new Date());
      return { success: true };
    }
    return { success: false };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function abrirPainelAdmin() {
  SpreadsheetApp.getUi().alert("O painel administrativo completo está disponível no Web App.\nAcesse pelo link de implantação.");
}
function forcarPermissao() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  console.log("Permissões OK");
}
// Função auxiliar para gerar link direto de imagem do Drive
function getLinkDiretoImagem(fileId) {
  return "https://drive.google.com/uc?export=view&id=" + fileId;
}
// Adicione no final do codigo.gs
function getLinkDiretoImagem(fileId) {
  return "https://drive.google.com/uc?export=view&id=" + fileId;
}
// Função para buscar lista de células para o formulário
function getListaCelulas() {
  const dados = getDadosDaPlanilha('CELULAS');
  // Retorna apenas ID e Nome para preencher o select
  return dados.map(c => ({ id: c.ID, nome: c.Nome }));
}
// Função para buscar lista de ministérios
function getListaMinisterios() {
  const dados = getDadosDaPlanilha('MINISTERIOS');
  // Retorna ID e Nome (ex: MIN001 e Louvor)
  return dados.map(m => ({ id: m.ID, nome: m.Nome }));
}
// Função avançada para buscar aulas com detalhes do professor e formatação
function getAulasDetalhadas(escolaFiltro) {
  const aulas = getDadosDaPlanilha('AULAS');
  const professores = getDadosDaPlanilha('PROFESSORES');
  
  // Filtra
  let aulasFiltradas = aulas;
  if (escolaFiltro) {
    aulasFiltradas = aulas.filter(a => a['Tipo Escola'] === escolaFiltro);
  }

  // Ordena por data
  aulasFiltradas.sort((a, b) => {
    if (!a.Data) return 1;
    if (!b.Data) return -1;
    return new Date(a.Data) - new Date(b.Data);
  });

  return aulasFiltradas.map(aula => {
    // 1. Foto do Professor
    let fotoProf = '';
    const prof = professores.find(p => p.ID === aula['Professor ID'] || p.Nome === aula['Professor Nome']);
    if (prof && prof['Foto URL'] && prof['Foto URL'].length > 10) {
       let url = prof['Foto URL'];
       try {
         let id = null;
         if (url.indexOf('id=') > -1) id = url.split('id=')[1].split('&')[0];
         else if (url.indexOf('/d/') > -1) id = url.split('/d/')[1].split('/')[0];
         if (id) fotoProf = "https://lh3.googleusercontent.com/d/" + id;
       } catch(e) {}
    }

    // 2. Formata Data (Visual)
    let dataF = aula.Data;
    if (aula.Data && aula.Data.includes('-')) {
        const partes = aula.Data.split('-'); // 2026-02-02
        dataF = `${partes[2]}/${partes[1]}/${partes[0]}`; // 02/02/2026
    }

    // 3. Hora (Agora já vem limpa "19:30" da função principal)
    let horaF = aula.Horário || '00:00';
    // Garantia extra: se ainda tiver data misturada, limpa
    if (horaF.length > 5) horaF = horaF.substring(0, 5);

    return {
      ...aula,
      'DataFormatada': dataF,
      'HoraFormatada': horaF,
      'FotoProfessor': fotoProf,
      'ProfessorNome': prof ? prof.Nome : (aula['Professor Nome'] || 'Professor'),
      'CargoProfessor': prof ? (prof.Ministerio || 'Professor') : 'Docente',
      'Descricao': aula['Descrição'] || '', 
      'DataOriginal': aula.Data
    };
  });
}
// FUNÇÃO AUXILIAR PARA UPLOAD DE FOTOS NO REGISTRO
function uploadImagemParaDrive(base64Data, nomeArquivo) {
  try {
    // 1. Limpa o cabeçalho do base64 se existir (data:image/jpeg;base64,...)
    let encoded = base64Data;
    if (encoded.includes(',')) {
      encoded = encoded.split(',')[1];
    }

    // 2. Define ou cria a pasta para as fotos
    const nomePasta = "SistemaEscolaBiblica_Fotos_Perfil";
    const pastas = DriveApp.getFoldersByName(nomePasta);
    let pasta;
    
    if (pastas.hasNext()) {
      pasta = pastas.next();
    } else {
      pasta = DriveApp.createFolder(nomePasta);
      // Deixa a pasta pública (opcional, mas ajuda na herança de permissão)
      pasta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    // 3. Cria o arquivo (Blob)
    const blob = Utilities.newBlob(
      Utilities.base64Decode(encoded), 
      'image/jpeg', 
      nomeArquivo
    );

    // 4. Salva o arquivo na pasta
    const arquivo = pasta.createFile(blob);

    // 5. Garante permissão pública para a imagem aparecer no site
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 6. Retorna o link direto (Thumbnail grande) para evitar redirecionamentos
    return "https://drive.google.com/uc?export=view&id=" + arquivo.getId();

  } catch (e) {
    console.error("Erro no upload da imagem de perfil:", e);
    return ""; // Retorna vazio se der erro, para não quebrar o cadastro
  }
}
function arrumarCabecalhosExatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('ALUNOS');
  
  if (!aba) {
    // Se não existir, cria
    ss.insertSheet('ALUNOS');
  }

  // A ordem EXATA que você pediu
  const cabecalhos = [
    'ID',                // Coluna A
    'Data Cadastro',     // Coluna B
    'Nome Completo',     // Coluna C
    'Data Nascimento',   // Coluna D
    'Email',             // Coluna E
    'Telefone',          // Coluna F
    'Foto URL',          // Coluna G
    'Célula',            // Coluna H
    'Ministério',        // Coluna I
    'Tipo Escola',       // Coluna J
    'Batizado',          // Coluna K
    'Data Batismo',      // Coluna L
    'Igreja Batismo',    // Coluna M
    'Interesse Servir',  // Coluna N
    'Observações',       // Coluna O
    'Qr code',           // Coluna P
    'User ID',           // Coluna Q
    'Status'             // Coluna R
  ];

  // Limpa a primeira linha e define os novos títulos
  aba.getRange(1, 1, 1, aba.getMaxColumns()).clearContent();
  const range = aba.getRange(1, 1, 1, cabecalhos.length);
  range.setValues([cabecalhos]);
  
  // Estilo
  range.setFontWeight("bold").setBackground("#EFEFEF").setHorizontalAlignment("center");
  aba.setFrozenRows(1);
  
  console.log("Cabeçalhos da aba ALUNOS organizados!");
  return "Pronto! A aba ALUNOS está com as colunas na ordem correta.";
}
/* --- FUNÇÃO QUE FALTAVA: BUSCAR ALUNOS DA ESCOLA DO PROFESSOR --- */
function getAlunosPorEscolaProfessor(professorId) {
  // 1. Busca todos os professores para achar o atual
  const professores = getDadosDaPlanilha('PROFESSORES');
  
  // Converte para String para garantir a comparação
  const professor = professores.find(p => String(p.ID).trim() === String(professorId).trim());

  // Se não encontrar o professor ou ele não tiver escola definida, retorna vazio
  if (!professor || !professor.Escola) {
    console.log("Professor ou Escola não encontrados para o ID: " + professorId);
    return []; 
  }

  const escolaDoProf = professor.Escola;
  console.log("Buscando alunos da escola: " + escolaDoProf);

  // 2. Busca todos os alunos
  const alunos = getDadosDaPlanilha('ALUNOS');

  // 3. Filtra apenas os alunos daquela escola
  return alunos.filter(a => {
    // Verifica a coluna 'Tipo Escola' (padrão) ou 'Escola' (caso tenha mudado)
    const escolaAluno = a['Tipo Escola'] || a['Escola'];
    return escolaAluno === escolaDoProf;
  });
}
// ==========================================
// NOVA FUNÇÃO: RELATÓRIO AGREGADO PARA O PROFESSOR
// ==========================================

/* --- FUNÇÃO ESPECÍFICA PARA FREQUÊNCIA DOS ALUNOS (ACESSO PROFESSOR) --- */
function getFrequenciaAlunosDetalhada(professorId) {
  try {
    // Buscar professor para saber sua escola
    const professor = getProfessorPorUserId(professorId);
    if (!professor || !professor.Escola) {
      return { success: false, message: 'Professor ou escola não encontrados' };
    }
    
    const escolaProfessor = professor.Escola;
    
    // Buscar alunos da mesma escola
    const alunos = getDadosDaPlanilha('ALUNOS');
    const alunosEscola = alunos.filter(a => 
      (a['Tipo Escola'] === escolaProfessor || a['Escola'] === escolaProfessor) && 
      a['Status'] !== 'inativo'
    );
    
    // Buscar todas as frequências DESSA ESCOLA
    const frequencias = getDadosDaPlanilha('FREQUENCIAS');
    const frequenciasEscola = frequencias.filter(f => f.Escola === escolaProfessor);
    
    // Buscar total de aulas dessa escola
    const aulas = getDadosDaPlanilha('AULAS');
    const totalAulasEscola = aulas.filter(a => 
      a['Tipo Escola'] === escolaProfessor && a.Status === 'realizada'
    ).length;
    
    // Usar 40 como padrão se não tiver aulas realizadas
    const totalAulasParaCalculo = totalAulasEscola > 0 ? totalAulasEscola : 40;
    
    // Calcular para cada aluno
    const resultado = alunosEscola.map(aluno => {
      // Filtrar frequências deste aluno DESSA ESCOLA
      const freqAluno = frequenciasEscola.filter(f => 
        f['Aluno ID'] === aluno['ID'] || f['Aluno Nome'] === aluno['Nome Completo']
      );
      
      // Contar frequências
      const totalFrequencias = freqAluno.length;
      
      // Calcular porcentagem baseado no total de aulas realizadas
      const porcentagem = totalFrequencias > 0 ? 
        Math.min(100, Math.round((totalFrequencias / totalAulasParaCalculo) * 100)) : 0;
      
      // Última frequência
      const ultimaFreq = freqAluno.length > 0 ? 
        freqAluno.sort((a, b) => new Date(b.Data) - new Date(a.Data))[0] : null;
      
      // Contar frequências por aula
      const frequenciasPorAula = {};
      freqAluno.forEach(f => {
        if (f['Aula ID']) {
          frequenciasPorAula[f['Aula ID']] = (frequenciasPorAula[f['Aula ID']] || 0) + 1;
        }
      });
      
      return {
        id: aluno['ID'],
        nome: aluno['Nome Completo'],
        escola: aluno['Tipo Escola'] || aluno['Escola'] || escolaProfessor,
        celula: aluno['Célula'] || '-',
        totalFrequencias: totalFrequencias,
        porcentagem: porcentagem,
        ultimaData: ultimaFreq ? formatarData(ultimaFreq.Data) : 'Nunca',
        ultimaAula: ultimaFreq && ultimaFreq['Aula ID'] ? ultimaFreq['Aula ID'] : '-',
        aulasDistintas: Object.keys(frequenciasPorAula).length,
        status: porcentagem >= 75 ? '🟢 Regular' : 
               porcentagem >= 50 ? '🟡 Atenção' : '🔴 Crítico',
        frequenciasDetalhadas: freqAluno.map(f => ({
          data: formatarData(f.Data),
          aulaId: f['Aula ID'] || '-',
          horario: f.Horário ? formatarHora(f.Horário) : '-'
        }))
      };
    });
    
    // Ordenar por porcentagem (maior primeiro)
    resultado.sort((a, b) => b.porcentagem - a.porcentagem);
    
    return {
      success: true,
      escola: escolaProfessor,
      totalAlunos: resultado.length,
      totalAulasRealizadas: totalAulasEscola,
      totalAulasParaCalculo: totalAulasParaCalculo,
      alunosComFrequencia: resultado.filter(a => a.totalFrequencias > 0).length,
      mediaPorcentagem: resultado.length > 0 ? 
        Math.round(resultado.reduce((sum, a) => sum + a.porcentagem, 0) / resultado.length) : 0,
      dados: resultado
    };
    
  } catch(e) {
    console.error('Erro ao buscar frequência detalhada:', e);
    return { success: false, message: 'Erro: ' + e.message };
  }
}

// Função auxiliar para formatar data
function formatarData(dataStr) {
  try {
    const data = new Date(dataStr);
    return Utilities.formatDate(data, "GMT-3", "dd/MM/yyyy");
  } catch(e) {
    return dataStr;
  }
}
function getAulasDoDiaPorEscola(escola, data = null) {
  try {
    const dataBusca = data ? new Date(data) : new Date();
    const dataFormatada = dataBusca.toISOString().split('T')[0];
    
    const aulas = getDadosDaPlanilha('AULAS');
    
    return aulas.filter(aula => {
      if (!aula.Data || aula['Tipo Escola'] !== escola || aula.Status !== 'agendada') {
        return false;
      }
      
      const dataAula = new Date(aula.Data);
      const dataAulaFormatada = dataAula.toISOString().split('T')[0];
      
      return dataAulaFormatada === dataFormatada;
    });
    
  } catch(e) {
    console.error('Erro ao buscar aulas do dia:', e);
    return [];
  }
}
function registrarFrequenciaComAulaId(dados) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FREQUENCIAS');
    const agora = new Date();
    
    // Buscar aluno
    const alunos = getDadosDaPlanilha('ALUNOS');
    const aluno = alunos.find(a => 
      a['ID'] === dados.alunoId || 
      a['User ID'] === dados.alunoId || 
      a['Nome Completo'] === dados.alunoNome
    );
    
    // Determinar escola
    let escola = '';
    if (aluno) {
      escola = aluno['Tipo Escola'] || aluno['Escola'] || '';
    }
    if (!escola) {
      escola = dados.escola || '';
    }
    
    // Usar o aulaId enviado pelo front-end
    let aulaId = dados.aulaId || '';
    
    // Verificar se já registrou para ESTA AULA
    const freqs = getDadosDaPlanilha('FREQUENCIAS');
    if (aulaId) {
      const jaRegistrouEstaAula = freqs.some(f => 
        (f['Aluno ID'] === dados.alunoId || f['Aluno Nome'] === dados.alunoNome) &&
        f['Aula ID'] === aulaId
      );
      
      if (jaRegistrouEstaAula) {
        return { 
          success: false, 
          jaRegistrado: true, 
          message: 'Já registrou frequência para esta aula!' 
        };
      }
    }
    
    // Verificar se já registrou HOJE (independente da aula)
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    const jaRegistrouHoje = freqs.some(f => {
      try {
        const dataFreq = f.Data ? new Date(f.Data) : null;
        if (!dataFreq) return false;
        
        dataFreq.setHours(0, 0, 0, 0);
        const mesmoAluno = f['Aluno ID'] === dados.alunoId || f['Aluno Nome'] === dados.alunoNome;
        
        return mesmoAluno && dataFreq.getTime() === hoje.getTime();
      } catch(e) {
        return false;
      }
    });
    
    if (jaRegistrouHoje && !aulaId) {
      return { 
        success: false, 
        jaRegistrado: true, 
        message: 'Já registrou frequência hoje!' 
      };
    }
    
    // Processar foto
    let fotoURL = '';
    if (dados.fotoData) {
      try {
        let encoded = dados.fotoData;
        if (encoded.includes(',')) {
          encoded = encoded.split(',')[1];
        }
        const blob = Utilities.newBlob(
          Utilities.base64Decode(encoded), 
          'image/jpeg', 
          'freq_' + (dados.alunoId || 'aluno') + '_' + Date.now() + '.jpg'
        );
        const file = DriveApp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoURL = file.getUrl();
      } catch(uploadError) {
        console.error('Erro no upload da foto:', uploadError);
      }
    }
    
    // Gerar ID e link do mapa
    const registroId = 'FREQ' + Utilities.getUuid().slice(0, 8).toUpperCase();
    const mapsLink = `https://www.google.com/maps/search/?api=1&query=${dados.latitude || 0},${dados.longitude || 0}`;
    
    // Registrar na planilha
    sheet.appendRow([
      registroId,                    // ID
      agora,                         // Data
      dados.alunoId || '',           // Aluno ID
      dados.alunoNome,               // Aluno Nome
      'Sim',                         // Presente
      agora.toLocaleTimeString('pt-BR'), // Horário
      mapsLink,                      // Localização
      fotoURL,                       // Foto URL
      dados.dispositivo || 'Web App', // Dispositivo
      aulaId,                        // Aula ID (SEMPRE enviado)
      escola                         // Escola
    ]);
    
    return { 
      success: true, 
      message: 'Frequência registrada com sucesso!',
      fotoURL: fotoURL, 
      localizacaoURL: mapsLink,
      aulaId: aulaId,
      escola: escola,
      id: registroId
    };
    
  } catch(e) {
    console.error('Erro ao registrar frequência com aulaId:', e);
    return { success: false, message: 'Erro: ' + e.message };
  }
}
// Adicione esta função ao final do seu codigo.js
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Modifique a função doGet para incluir o Service Worker
function doGet(e) {
  const { queryString } = e.parameter;
  
  // Se for requisição para o manifesto
  if (queryString === 'manifest') {
    return HtmlService.createHtmlOutputFromFile('manifest')
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Retorna o app normal
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('Escola Bíblica IEQ Sede Região 1065')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
