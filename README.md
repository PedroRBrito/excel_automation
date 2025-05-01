# **Comparador de Planilhas Excel**

![Tela Principal](assets/janela.png)

Um aplicativo desenvolvido em Python com PyQt5 para comparar planilhas Excel, realizar cálculos personalizados e exibir os resultados de forma clara e intuitiva.

---

## **Funcionalidades**

- **Carregar Planilhas**:

![Planilha carregada](assets/table_carregada.png)

  - Permite carregar duas planilhas Excel para comparação.
  - Exibe os dados carregados em tabelas interativas.

---

- **Lista de colunas para realizar os cálculos**:

![Escolha de colunas para os cálculos](assets/coluna_calculo.png)

  - Após carregar os arquivos, o usuário pode selecionar as colunas de cada planilha que serão usadas como base para os cálculos.
  - O programa permite escolher as colunas específicas de cada planilha.

---

- **Comparação de Dados**:

![Cálculos possíveis](assets/calculos.png)

  - Realiza cálculos personalizados entre as colunas selecionadas:

    - Soma
    - Diferença
    - Multiplicação
    - Divisão
    - Média
    - Diferença Percentual

---

- **Janela de Comparação**:

![Janela de Comparação](assets/janela_comparacao.png)

  - Após a comparação, os resultados são exibidos em uma nova janela.
  - O usuário pode visualizar os dados comparados em uma tabela interativa.

---

- **Selecionar Colunas para Exportação**:

![Seleção de colunas para exportação](assets/colunas_exportar_selecao.png)

  - Na Janela de Comparação, o usuário pode escolher quais colunas do resultado final serão exportadas.
  - Checkboxes permitem selecionar as colunas desejadas.

---

- **Exportar Resultados**:

![Salvando excel na pasta](assets/salvando_excel.png)

![Arquivo excel aberto](assets/excel_salvo.png)

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
   git clone https://github.com/seu-usuario/excel-automation.git
   cd excel-automation
