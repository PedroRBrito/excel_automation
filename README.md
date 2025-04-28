# Comparador de Planilhas em Python (PyQt5 + Pandas)

Aplicação desktop feita com Pandas e PyQt5 para leitura, comparação e exportação de planilhas Excel (.xlsx). O programa permite carregar dois arquivos, comparar valores com base em uma coluna comum (como "Categoria") e gerar automaticamente o cálculo escolhido pelo usuário.

## Funcionalidades

- Leitura de duas planilhas Excel.
- Visualização dos dados com ordenação.
- Barra de progresso durante o carregamento.
- Comparação entre dois valores.
- Cálculo automático dos valores.
- Exportação do resultado para um novo arquivo Excel.

## Tecnologias Utilizadas

- **PyQt5**: Para a criação da interface gráfica do usuário.
- **Pandas**: Para o processamento e manipulação de dados em planilhas.
- **Python**: Linguagem de programação utilizada para o desenvolvimento da aplicação.

## Requisitos

- Python 3.x
- PyQt5 (`pip install pyqt5`)
- Pandas (`pip install pandas`)

## Como Rodar o Projeto

1. **Clone o repositório**:

   ```bash
   git clone https://github.com/seu-usuario/comparador-planilhas.git
   cd comparador-planilhas
   ```

2. **Instale as dependências**:

   ```bash
   pip install -r requirements.txt
   ```

3. **Execute a aplicação**:
   ```bash
   python comparador_planilhas.py
   ```

## Como Usar

1. Ao abrir a aplicação, você verá duas áreas de visualização de tabelas.
2. Clique em "Abrir primeiro arquivo" para carregar a primeira planilha Excel.
3. Clique em "Abrir segundo arquivo" para carregar a segunda planilha Excel.
4. Escolha as colunas das planilhas para serem usadas no cálculo.
5. Selecione as operações que deseja realizar (Soma, Multiplicação, Divisão, Média, Diferença e Diferença Percentual).
6. Clique em "Comparar planilhas" para visualizar a comparação entre os dados das duas planilhas.
7. Você pode exportar o resultado para um novo arquivo Excel.
