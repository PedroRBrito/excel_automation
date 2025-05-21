# **Comparador de Planilhas Excel**

> Se o GIF abaixo não rodar automaticamente, clique no ícone "play" no canto superior direito da imagem ou verifique suas [configurações de acessibilidade no GitHub](https://github.com/settings/accessibility).

<img src="assets/demo_excel_automation.gif" alt="Demo" width="800">

Um aplicativo desenvolvido em Python com PyQt5 e Pandas para comparar planilhas Excel, realizar cálculos personalizados e exibir os resultados de forma clara e intuitiva.

Este projeto foi criado para automatizar o tratamento de dados de planilhas, economizando tempo em tarefas repetitivas como: leitura, limpeza, e exportação de informações. É útil para quem trabalha com grandes volumes de dados no Excel e precisa de uma interface simples para processá-los com um clique.



## **Funcionalidades**

### **Carregar Planilhas**:

<img src="assets/table_carregada.png" alt="Planilha carregada" width="500">

  - Permite carregar duas planilhas Excel para comparação.
  - Exibe os dados carregados em tabelas interativas.


### **Lista de colunas para realizar os cálculos**:

<img src="assets/coluna_calculo.png" alt="Escolha de colunas para os cálculos" width="00">

  - Após carregar os arquivos, o usuário pode selecionar as colunas de cada planilha que serão usadas como base para os cálculos.
  - O programa permite escolher as colunas específicas de cada planilha.


### **Comparação de Dados**:

<img src="assets/calculos.png" alt="Cálculos possíveis" width="800">

  - Realiza cálculos personalizados entre as colunas selecionadas:

    - Soma
    - Diferença
    - Multiplicação
    - Divisão
    - Média
    - Diferença Percentual

### **Janela de Comparação**:

<img src="assets/janela_comparacao.png" alt="Janela de Comparação" width="500">

  - Após a comparação, os resultados são exibidos em uma nova janela.
  - O usuário pode visualizar os dados comparados em uma tabela interativa.


### **Selecionar Colunas para Exportação**:

<img src="assets/colunas_exportar_selecao.png" alt="Seleção de colunas para exportação" width="500">

  - Na Janela de Comparação, o usuário pode escolher quais colunas do resultado final serão exportadas.
  - Na lista é permitido selecionar as colunas desejadas.


### **Exportar Resultados**:

<img src="assets/salvando_excel.png" alt="Salvando excel na pasta" width="500">

<img src="assets/excel_salvo.png" alt="Arquivo excel aberto" width="500">

  - Permite exportar os resultados da comparação para um arquivo Excel.



## **Tecnologias Utilizadas**

- **Python 3.10+**
- **PyQt5**: Para a interface gráfica.
- **Pandas**: Para manipulação de dados.
- **Openpyxl**: Para gerenciar arquivos .xlsx.



## **Como Executar o Projeto**

### **Pré-requisitos**
Certifique-se de ter o Python 3.9 ou superior instalado em sua máquina. Instale as dependências do projeto usando o `pip`.

> O pacote PyQt5 precisa ser instalado manualmente com `pip install PyQt5`.

### **Passos**
1. Clone o repositório:
   ```bash
   git clone https://github.com/PedroRBrito/excel-automation.git
   cd excel-automation
