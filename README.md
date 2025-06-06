
# Relatório de propostas API Cartos

Esse projeto gera uma planilha atualizada de propostas buscando na api da Cartos.

Esse projeto roda em Docker, bastando realizar a parametrização dos dados em um arquivo .env,
que tem um .env.example para saber quais são esses dados, e caso precise atualizar os parâmetros de busca, 
os faça no arquivo app.py, tanto para o mês desejado quanto para o total de itens desejados.

Hoje o mês padrão é o de maio e caso precise alterar mude as variáveis STARTDATE e ENDDATE.

## Variáveis de Ambiente

Para rodar esse projeto, você vai precisar adicionar a seguinte variável de ambiente no seu .env

  `API_EMAIL`
  `API_PASSWORD`


## Rodando localmente

- Clone o projeto

```bash
  https://github.com/valdir-veloon/relatorio-api
```

- Entre no diretório do projeto

```bash
  cd relatorio-api
```

- Renomeie as variáveis .env.example para .env e preencha com os dados de acesso

- Configure os parâmetros em app.py de data e limites, e rode o comando abaixo e uma nova planilha será gerada na sua máquina

```bash
  docker compose up --build
```

## Stack utilizada

**Back-end:** Python com Docker


## Autor

- [@Valdir Silva](https://github.com/valdir-veloon)
