# Automação de Consultas - PGE SP

Este projeto é uma automação para realizar consultas de débitos no site da Procuradoria Geral do Estado de São Paulo (PGE SP) usando Python, Selenium e integração com Excel. A interface gráfica foi desenvolvida em Tkinter para facilitar o uso.

## Funcionalidades
- Automação de consultas de débitos no site da PGE SP.
- Preenchimento automatizado de planilhas Excel com resultados das consultas.
- Geração de PDFs dinâmicos para cada consulta realizada.
- Suporte a resolução de reCAPTCHA via API do AntiCaptcha.
- Interface gráfica com:
  - Seleção de arquivo Excel e diretório de saída.
  - Nomeação personalizada para demandas.
  - Barra de progresso e área de logs para acompanhamento em tempo real.

## Pré-requisitos
- Python 3.8 ou superior.
- Dependências listadas no arquivo `requirements.txt`:
  - `selenium`
  - `openpyxl`
  - `anticaptchaofficial`
  - `tkinter` (padrão em instalações do Python para Windows).
- Navegador Google Chrome e Chromedriver instalados.
- Conta na [AntiCaptcha](https://anti-captcha.com/) para resolver captchas.

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu_usuario/automacao_pge_sp.git
   cd automacao_pge_sp
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure o arquivo `chave_API.py` com sua chave de API do AntiCaptcha:
   ```python
   chave_api = "sua_chave_de_api"
   ```

4. Certifique-se de que o Chromedriver está configurado no `PATH` ou na mesma pasta do projeto.

## Uso

1. Execute o programa:
   ```bash
   python nome_do_script.py
   ```

2. Na interface gráfica:
   - Escolha o arquivo Excel contendo os dados de consulta.
   - Insira um nome para a demanda.
   - Escolha o diretório onde os PDFs e resultados serão salvos.
   - Clique em **Iniciar** para começar o processamento.

## Estrutura do Projeto

- **`start_process`**: Função principal que executa as consultas e preenche os resultados no Excel.
- **`gerar_pdf_dinamico`**: Gera PDFs das páginas de resultados das consultas.
- **`main`**: Cria a interface gráfica do usuário.
- **`configurar_chrome_options`**: Configura o navegador para automação.

## Principais Tecnologias
- **[Selenium](https://www.selenium.dev/)**: Automação de interação com navegadores.
- **[Tkinter](https://docs.python.org/3/library/tkinter.html)**: Interface gráfica para aplicativos Python.
- **[OpenPyXL](https://openpyxl.readthedocs.io/)**: Manipulação de arquivos Excel.
- **[AntiCaptcha](https://anti-captcha.com/)**: Resolução automática de reCAPTCHA.

## Licença
Este projeto está licenciado sob a [MIT License](LICENSE).

## Contato
- **Autor**: Matheus Madureira da Fonseca
- **E-mail**: madureira-matheus@hotmail.com
