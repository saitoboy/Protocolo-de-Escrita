# Protocolo-de-Escrita

Protocolos de Escrita
Introdução
Na Secretaria de Educação, realizamos avaliações diagnósticas, incluindo o Avalia Muriaé para avaliar o desempenho em matemática e português. Também implementamos os protocolos de escrita. Para os Anos Iniciais, em 2024, realizaremos os protocolos no formato qualitativo. As professoras serão responsáveis por avaliar e registrar o nível de escrita e leitura dos alunos. Os protocolos são categorizados da seguinte maneira:

Educação Infantil - Escrita Espontânea (1º Período e 2º Período)
Anos Iniciais - Ditado Conceitual (1º ao 3º Ano)
Anos Iniciais - Fluência de Leitura (2º ao 5º Ano)
Anos Iniciais - Produção de Texto (4º e 5º Ano)
Registro de Notas
As professoras devem apenas selecionar seus alunos e registrar a resposta pedida. Caso algum aluno não esteja listado, a professora pode adicioná-lo manualmente.

Passo 1 - Extrair dados do Sislame
Você irá na aba "Relatórios" e "BI", depois clicará em "Dados Cadastrais Alunos".


A planilha vem com algumas linhas em branco. Selecione elas e exclua. A planilha deve ter umas 60 colunas, no entanto, não precisamos de todas elas. Para isso, você usará o script abaixo, feito em Python Pandas, que deixará apenas os dados que você precisa:

python
Copiar código
# TIRAR COLUNAS DESNECESSÁRIAS DO SISLAME #
import pandas as pd

# Carregar o arquivo Excel
df = pd.read_excel('ALUNOS_DADOS_CADASTRAIS.xlsx')

# Selecionar apenas as colunas necessárias
colunas_necessarias = ['ESCOLA', 'NOME', 'ETAPA', 'CD TURMA']
df_selecionado = df[colunas_necessarias]

# Mostrar o cabeçalho e as três primeiras linhas do DataFrame 
print("Operação realizada com sucesso!\n")
print("Cabeçalho:")
print(df_selecionado.head())

# Salvar o novo DataFrame em um novo arquivo Excel
df_selecionado.to_excel('resultado.xlsx', index=False)
Passo 2 - Separar os Dados por Turmas
Com o script abaixo, você irá separar os dados para cada turma, criando um Excel para cada uma.

python
Copiar código
import pandas as pd

# Carregar o arquivo Excel
df = pd.read_excel('resultado.xlsx')

# Ordenar os nomes dos alunos em ordem alfabética
df.sort_values(by='NOME', inplace=True)

# Agrupar os dados por turma
grupos = df.groupby('CD TURMA')

# Iterar sobre cada grupo e salvar em planilhas separadas
for grupo, dados_grupo in grupos:
    escola = dados_grupo['ESCOLA'].iloc[0]
    etapa = dados_grupo['ETAPA'].iloc[0]
    turma = dados_grupo['TURMA'].iloc[0]
    nome_arquivo = f'{escola}_{etapa}_{turma}_{grupo}.xlsx'
    
    # Salvar os dados da turma em uma planilha separada
    dados_grupo.to_excel(nome_arquivo, index=False)
    print(f'Planilha "{nome_arquivo}" foi criada com sucesso!')
Passo 3 - Transformar Turmas em Linhas
Transforme todas essas planilhas separadas em linhas, onde cada planilha ficará em uma única linha de uma planilha. Os alunos serão separados por vírgula, o que é importante para usar esses dados na construção dos forms.

python
Copiar código
import os
import pandas as pd

# Caminho para a pasta com os arquivos
caminho_pasta = r'C:\Users\guilherme.saito\Documents\Scripts\P'

# Lista para armazenar os dados extraídos
dados = []

# Percorre os arquivos na pasta
for arquivo in os.listdir(caminho_pasta):
    if arquivo.endswith('.xlsx'):
        # Extrai o nome da escola, ano e turma do arquivo
        nome_arquivo, extensao = os.path.splitext(arquivo)
        partes = nome_arquivo.split('_', 1)
        if len(partes) == 2:
            nome_escola, resto = partes
            ano_turma_cd, extensao = os.path.splitext(resto)
            ano_turma_cd_parts = ano_turma_cd.rsplit('_', 1)
            if len(ano_turma_cd_parts) == 2:
                ano_turma, cd_turma = ano_turma_cd_parts
                
                # Caminho completo para o arquivo XLSX
                caminho_arquivo = os.path.join(caminho_pasta, arquivo)
                
                # Lê a planilha do arquivo XLSX
                df_planilha = pd.read_excel(caminho_arquivo)
                
                # Extrai os nomes dos alunos da coluna 'NOME'
                alunos = df_planilha['NOME'].tolist()
                
                # Ordena os nomes dos alunos em ordem alfabética
                alunos.sort()
                
                # Adiciona os dados à lista
                dados.append({'Escola': nome_escola, 'Ano': ano_turma, 'Turma_CD': cd_turma, 'Alunos': ','.join(alunos)})

# Cria um DataFrame com os dados
df = pd.DataFrame(dados)

# Reorganiza as colunas para que 'Turma_CD' seja a última coluna
cols = df.columns.tolist()
cols.remove('Turma_CD')
cols.append('Turma_CD')
df = df[cols]

# Exporta o DataFrame para um arquivo XLSX
df.to_excel('turmas_por_linhas.xlsx', index=False)
Veja que eu escolhi para onde o script deveria ler, que era naquela pasta. Recomendo deixar em pasta separada porque ele vai ler tudo que está na pasta, ok? Se tudo der certo, ficará assim.

Criação de Formulários no Google Forms
Agora você deverá ir no SEU DRIVE. Os Scripts não funcionam em drives compartilhados, e o email projetos@edu está bugado para script também.

Antes de você continuar, faça uma coluna antes da escola e use a função CONCATENAR (sheets) para que você já tenha um "nome do formulário" que você possa identificar depois. A função é essa:

excel
Copiar código
=CONCATENAR(B2;"_";C2)
Bom, agora vamos ao script:

javascript
Copiar código
function Linhas_Forms() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var planilha = ss.getSheetByName('Página1');
  var dados = planilha.getDataRange().getValues();
  var pastaDestinoId = '16gwUwuyVDeRq1HSN0ynlxUPygBICb6XN';
  var pastaDestino = DriveApp.getFolderById(pastaDestinoId);
  var linhaInicial = 1;

  // Encontrar a primeira linha vermelha
  for (var i = 1; i < dados.length; i++) {
    var corLinha = planilha.getRange(i + 1, 1).getBackground();
    if (corLinha === 'red') {
      linhaInicial = i + 1;
      break;
    }
  }

  for (var i = linhaInicial; i < dados.length; i++) {
    try {
      var tituloFormulario = dados[i][0];
      var descricaoFormulario = dados[i][1];
      var novoFormulario = FormApp.create(tituloFormulario);
      novoFormulario.setDescription(descricaoFormulario);
      var arquivoFormulario = DriveApp.getFileById(novoFormulario.getId());
      arquivoFormulario.moveTo(pastaDestino);
      var perguntaAluno = novoFormulario.addMultipleChoiceItem()
        .setTitle('Aluno')
        .setChoiceValues(dados[i][3].toString().split(","))
        .setRequired(true);
      
      var perguntaAlunoNaoEncontrado = novoFormulario.addTextItem()
        .setTitle('Caso não encontre o nome de algum aluno, informe aqui:');
      
      var perguntaNivelEscrita = novoFormulario.addMultipleChoiceItem()
        .setTitle('Nível de Escrita')
        .setChoiceValues(['GARATUJA', 'PRÉ-SILÁBICA', 'SILÁBICA', 'ALFABÉTICA'])
        .setRequired(true);

      // Configurar as opções de envio de resposta
      novoFormulario.setRequireLogin(true); // Login não necessário
      novoFormulario.setLimitOneResponsePerUser(false); // Permitir múltiplas respostas
      novoFormulario.setCollectEmail(false); // Coletar endereço de e-mail
      novoFormulario.setPublishingSummary(true); // Publicar resumo das respostas
      novoFormulario.setShowLinkToRespondAgain(true); // Mostrar link para responder novamente
      novoFormulario.setAllowResponseEdits(false); // Permitir edição das respostas

      // Mudar a cor da linha para verde
      var linha = planilha.getRange(i + 1, 1, 1, planilha.getLastColumn());
      linha.setBackground('green');
      Logger.log('Formulário criado: ' + novoFormulario.getPublishedUrl());
    } catch (e) {
      // Erro ao gerar o formulário, alterar a cor da linha para vermelho
      var linha = planilha.getRange(i + 1, 1, 1, planilha.getLastColumn());
      linha.setBackground('red');
      Logger.log('Erro ao criar formulário na linha ' + (i + 1) + ': ' + e.message);
    }
  }
}
Lembre-se de mudar o ID da pasta e talvez você tenha que colocar o da planilha também. Se tudo der certo, ficará assim:
