# Buscador de preços

Automação de buscar preços no Google Shopping e no Buscapé, com filtragem por faixa de preço, termos específicos e envio dos resultados por email.

## Funcionalidades

- Busca automática de produtos no Google Shopping e Buscapé
- Filtragem por faixa de preço mínimo e máximo
- Exclusão de produtos com termos indesejados (ex: "usado", "defeito")
- Consolidação dos resultados numa planilha Excel
- Envio automático dos resultados por email

## Bibliotecas utilizadas

- [Selenium](https://www.selenium.dev/) - automação do navegador
- [WebDriver Manager](https://github.com/SergeyPirogov/webdriver_manager) - gerenciamento do ChromeDriver
- [Pandas](https://pandas.pydata.org/) - manipulação de dados e DataFrames
- [OpenPyXl](https://openpyxl.readthedocs.io/) - geração de arquivos Excel
- [Python Dotenv](https://github.com/theskumar/python-dotenv) - gerenciamento de variáveis de ambiente
- smtplib - envio de emails via SMTP (nativa do Python)
- email.mime - montagem e formatação dos emails (nativa do Python)

## Instalação

**1. Clone o repositório:**
```bash
git clone https://github.com/pedro-henrique-gg/busca-preco.git
cd busca-preco
```

**2. Instale as dependências:**
```bash
pip install -r requirements.txt
```

**3. Configure as variáveis de ambiente:**

Crie um arquivo `.env` na raiz do projeto com o seguinte conteúdo:
```
gmail_remetente=seu@gmail.com
gmail_destinatario=destino@gmail.com
gmail_senha=sua_senha_de_app
```

## Como executar

Configure o arquivo `buscas.xlsx` com os produtos desejados seguindo esse formato:

| Nome | Termos banidos | Preço mínimo | Preço máximo |
|---|---|---|---|
| iphone 12 64gb | usado defeito | 2000 | 3500 |

Depois execute o arquivo `main.py`

O script vai abrir o chrome automaticamente, realizar as buscas, salvar os resultados em `ofertas.xlsx` e enviar o arquivo por email.

## Observações 

- O script foi criado e testado no **Linux (Zorin OS)**
- É necessário ter o **Google Chrome** instalado, ou usar um navegador diferente no selenium
- A senha do email deve ser uma **senha de app** gerada nas configurações de segurança da conta Google, não a senha normal da conta
- O Google Shopping pode exibir dois layouts diferentes - o script lida com ambos automaticamente