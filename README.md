# Validador de Dados

Este projeto é um validador de dados entre planilhas Excel de duas pastas diferentes (SIFAC e SISPAT). Ele compara os dados dos empregados e seus contratos, além de validar a competência entre os arquivos.

## Funcionalidades

- Seleção de pastas contendo arquivos Excel.
- Validação de competência entre os arquivos das pastas SIFAC e SISPAT.
- Comparação de dados dos empregados e contratos entre as planilhas.
- Geração de uma planilha de validação com o status de cada empregado.

## Requisitos

- Python 3.x
- Bibliotecas:
  - customtkinter
  - tkinter
  - openpyxl
  - os
  - datetime
  - logging

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/validador-de-dados.git
   ```
2. Navegue até o diretório do projeto:
   ```bash
   cd validador-de-dados
   ```
3. Instale as dependências:
   ```bash
   pip install customtkinter openpyxl
   ```

## Uso

1. Execute o script:
   ```bash
   python Validador\ 4.0.py
   ```
2. Na interface gráfica, selecione as pastas contendo os arquivos Excel do SIFAC e SISPAT.
3. Clique no botão "Validar Dados" para iniciar a validação.
4. Acompanhe o progresso na barra de progresso e veja o status na interface.

## Estrutura do Código

- `ValidadorDeDados`: Classe principal que gerencia a interface gráfica e a lógica de validação.
  - `__init__`: Inicializa a classe e configura o logging e a GUI.
  - `setup_logging`: Configura o logging.
  - `setup_gui`: Configura a interface gráfica.
  - `criar_widgets`: Cria os widgets da interface gráfica.
  - `selecionar_pasta_sifac`: Abre o diálogo para selecionar a pasta SIFAC.
  - `selecionar_pasta_sispat`: Abre o diálogo para selecionar a pasta SISPAT.
  - `obter_empregados_sispat`: Obtém os empregados e contratos da pasta SISPAT.
  - `validar_competencia`: Valida a competência entre os arquivos das pastas SIFAC e SISPAT.
  - `obter_competencia_sispat`: Obtém a competência do arquivo SISPAT.
  - `obter_competencia_sifac`: Obtém a competência do arquivo SIFAC.
  - `validar_dados`: Valida os dados entre as planilhas SIFAC e SISPAT.
  - `run`: Inicia a interface gráfica.
  - Adicionadas novas funcionalidades ao Validador de Dados:
  - Implementação da função `obter_competencia_sifac` para obter a competência do arquivo SIFAC.
  - Implementação da função `formatar_competencia` para formatar a competência em um formato padrão.
  - Implementação da função `validar_competencia` para validar se as competências dos arquivos SIFAC e SISPAT são iguais.
  - Implementação da função `criar_lista_sispat` para criar uma lista de empregados do SISPAT que não foram encontrados no SIFAC.
  - Implementação da função `validar_dados` para validar os dados dos arquivos SIFAC e SISPAT e gerar uma aba "Dados_Validados" com o status de cada empregado.
  - Implementação das funções auxiliares `remover_colunas_em_branco` e `ajustar_largura_colunas` para melhorar a formatação das planilhas geradas.

Essas funcionalidades melhoram a capacidade do sistema de validar e comparar dados entre os arquivos SIFAC e SISPAT, além de gerar relatórios detalhados sobre o status dos empregados.

## Contribuição

1. Faça um fork do projeto.
2. Crie uma nova branch:
   ```bash
   git checkout -b minha-nova-funcionalidade
   ```
3. Faça suas alterações e commit:
   ```bash
   git commit -m 'Adiciona nova funcionalidade'
   ```
4. Envie para o repositório remoto:
   ```bash
   git push origin minha-nova-funcionalidade
   ```
5. Abra um Pull Request.

## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
