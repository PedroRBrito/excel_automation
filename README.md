# **Comparador de Planilhas Excel**

<img src="assets/demo_excel_automation.gif" alt="Demo" width="800">

Um aplicativo desenvolvido em Python com PyQt5 para comparar planilhas Excel, realizar cálculos personalizados e exibir os resultados de forma clara e intuitiva.

---

## **Funcionalidades**

- **Carregar Planilhas**:

<img src="assets/table_carregada.png" alt="Planilha carregada" width="500">

  - Permite carregar duas planilhas Excel para comparação.
  - Exibe os dados carregados em tabelas interativas.

---

- **Lista de colunas para realizar os cálculos**:

<img src="assets/coluna_calculo.png" alt="Escolha de colunas para os cálculos" width="00">

  - Após carregar os arquivos, o usuário pode selecionar as colunas de cada planilha que serão usadas como base para os cálculos.
  - O programa permite escolher as colunas específicas de cada planilha.

---

- **Comparação de Dados**:

<img src="assets/calculos.png" alt="Cálculos possíveis" width="800">

  - Realiza cálculos personalizados entre as colunas selecionadas:

    - Soma
    - Diferença
    - Multiplicação
    - Divisão
    - Média
    - Diferença Percentual

---

- **Janela de Comparação**:

<img src="assets/janela_comparacao.png" alt="Janela de Comparação" width="500">

  - Após a comparação, os resultados são exibidos em uma nova janela.
  - O usuário pode visualizar os dados comparados em uma tabela interativa.

---

- **Selecionar Colunas para Exportação**:

<img src="assets/colunas_exportar_selecao.png" alt="Seleção de colunas para exportação" width="500">

  - Na Janela de Comparação, o usuário pode escolher quais colunas do resultado final serão exportadas.
  - Na lista é permitido selecionar as colunas desejadas.

---

- **Exportar Resultados**:

<img src="assets/salvando_excel.png" alt="Salvando excel na pasta" width="500">

<img src="assets/excel_salvo.png" alt="Arquivo excel aberto" width="500">

  - Permite exportar os resultados da comparação para um arquivo Excel ou CSV.

---

## **Tecnologias Utilizadas**

- **Python 3.9+**
- **PyQt5**: Para a interface gráfica.
- **Pandas**: Para manipulação de dados.

---

## **Como Executar o Projeto**

### **Pré-requisitos**
Certifique-se de ter o Python 3.9 ou superior instalado em sua máquina. Instale as dependências do projeto usando o `pip`.

### **Passos**
1. Clone o repositório:
   ```bash
   git clone https://github.com/PedroRBrito/excel-automation.git
   cd excel-automation
