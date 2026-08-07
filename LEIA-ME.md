# Robô CEAGESP — Cotações de Pescado (Atacado)

Coleta os preços de atacado dos pescados de interesse em <https://ceagesp.gov.br/cotacoes/>
(categoria **PESCADOS**, Entreposto da Capital) e acumula o histórico.

📊 **Página de consulta: <https://tbrena.github.io/robo-ceagesp/>** — atualizada
automaticamente a cada coleta.

## Produtos monitorados

FILE DE TILAPIA · PANGASIUS · PINTADO · TAMBAQUI · TILÁPIA · TRUTA

Para incluir/remover produtos, edite a lista `PRODUTOS` no topo de `robo_ceagesp.py`.
A comparação ignora acentos e maiúsculas, mas o nome precisa bater com o do boletim
(ex.: `FILE DE TILAPIA` e `TILAPIA` são produtos diferentes no site).

## Como funciona

1. Abre a página e lê a lista de datas com boletim publicado (a própria página traz essa
   lista embutida — é o que alimenta o calendário do site).
2. Para cada data ainda não coletada, envia a consulta (`POST` com categoria + data) e
   lê a tabela de resultados.
3. Filtra os produtos monitorados e grava.

Rodar mais de uma vez no mesmo dia não duplica nada: o robô só busca as datas que ainda
não estão no histórico. Se a máquina ficar dias desligada, ele recupera sozinho tudo o que
o site ainda estiver exibindo (o site mantém cerca de 7 boletins).

**Importante:** a CEAGESP publica o boletim **3 vezes por semana** (normalmente segunda,
quarta e sexta), não todos os dias. Rodar diariamente é o certo — nos dias sem boletim novo
o robô simplesmente registra "nenhum boletim novo" e encerra. Alguns produtos também não
são cotados em todos os boletins; nesse caso aparece um aviso no log e a data fica sem
linha para aquele produto.

## Arquivos gerados (pasta `dados`)

| Arquivo | Conteúdo |
|---|---|
| `historico.csv` | histórico acumulado, `;` como separador, UTF-8 com BOM (abre direto no Excel) |
| `boletim_AAAA-MM-DD.csv` | cada boletim coletado, individualmente |
| `cotacoes_pescado.xlsx` | planilha com abas **Histórico** e **Último boletim**, já formatada |

A pasta `docs` guarda a página de consulta (`index.html`), regerada a cada coleta nova e
publicada em <https://tbrena.github.io/robo-ceagesp/>. Ela traz o último boletim com a
variação percentual sobre o anterior, gráfico da série (em R$/kg ou indexado a 100),
filtros por produto e período, tabela completa e download em CSV. Para reconstruir a
página sem consultar o site: `python gerar_pagina.py`.

Colunas: `data`, `categoria`, `produto`, `classificacao`, `unidade_peso`, `menor`, `comum`,
`maior`, `quilo`, `coletado_em`. **Menor / Comum / Maior** são os preços em R$ —
"Comum" é o valor mais praticado.

Log de execução em `logs/robo.log`.

## API para outros sistemas

Os dados também saem em JSON e CSV estáticos, publicados junto com a página e atualizados
nas mesmas duas coletas diárias. Não há servidor: são arquivos servidos por HTTPS pelo
GitHub Pages, que envia `Access-Control-Allow-Origin: *` — então dá para consumir até
direto do navegador. Leitura apenas, sem autenticação.

| Endpoint | Devolve |
|---|---|
| `api/index.json` | índice: período coberto, totais e lista de endpoints |
| `api/ultimo.json` | último boletim, com `variacao_percentual` sobre o anterior |
| `api/cotacoes.json` | histórico completo |
| `api/cotacoes.csv` | histórico completo em CSV (Power BI, Excel → Dados → Da Web) |
| `api/produtos.json` | produtos monitorados e seus identificadores |
| `api/produto/{id}.json` | série de um produto (`tilapia`, `file-de-tilapia`, `truta`…) |

Base: <https://tbrena.github.io/robo-ceagesp/api/>

Datas em `AAAA-MM-DD`, preços como número decimal com ponto, valores ausentes como `null`.

```python
import requests
u = requests.get("https://tbrena.github.io/robo-ceagesp/api/ultimo.json").json()
for r in u["dados"]:
    print(r["produto"], r["comum"], r["variacao_percentual"], "%")
```

Os arquivos são gerados por `gerar_api.py` (rode sozinho com `python gerar_api.py`).
O cache do GitHub Pages é de poucos minutos, então uma alteração pode demorar um pouco
para aparecer no consumidor.

## Histórico importado de planilha

O histórico começa em **07/01/2026**, importado de uma planilha Excel própria; a coleta
automática assumiu a partir de **20/07/2026**. Para importar outra planilha:

```bash
python importar_excel.py "caminho\da\planilha.xlsx"
```

Ela precisa ter uma linha de cabeçalho com `Data | Produto | Clas. | Uni/Peso | Menor |
Comum | Maior` (a ordem não importa). Opções: `--aba NOME` escolhe a aba, `--todos`
importa todos os produtos em vez de só os monitorados, e `--simular` mostra o resultado
sem gravar nada — vale sempre rodar antes.

Em conflito (mesma data, produto e classificação), **o registro do robô prevalece**, já
que ele lê direto do site da CEAGESP. As divergências são listadas no log para conferência.

### Ressalvas sobre os dados importados

- Na conferência da primeira importação, 20 dos 23 valores sobrepostos bateram exatamente
  com o que o robô coletou. Os 3 restantes estavam na linha de **27/07/2026** da planilha,
  com valores que o site publica para **29/07/2026** — inclusive com a mesma cobertura de
  produtos daquele boletim. Ou seja, aquele bloco da planilha estava com a data trocada.
  Prevaleceram os valores do robô, e o 29/07 correto entrou pela coleta automática.
- A planilha não tinha os boletins de **10/07/2026** e 29/07/2026. O de 10/07 continua
  ausente: quando percebemos, o site já não o exibia mais (ele mantém cerca de 7 boletins).
- Os dados anteriores a 20/07/2026 **não puderam ser conferidos contra o site**, pelo mesmo
  motivo. Dado o erro encontrado em 27/07, é possível que existam outros pontos com data
  trocada nesse trecho.

## Uso manual

```bash
python robo_ceagesp.py
```

| Comando | O que faz |
|---|---|
| `python robo_ceagesp.py` | coleta os boletins ainda não coletados (uso normal / agendado) |
| `python robo_ceagesp.py --data 03/08/2026` | coleta uma data específica |
| `python robo_ceagesp.py --todas` | recoleta todas as datas que o site exibe |
| `python robo_ceagesp.py --tudo` | coleta **todos** os pescados da categoria, não só a lista |

## Execução automática na nuvem (GitHub Actions)

**Ativo** desde 04/08/2026 em <https://github.com/tbrena/robo-ceagesp> (repositório privado).
O robô roda nos servidores do GitHub, **independente do seu PC estar ligado**.
A configuração está em `.github/workflows/coleta-ceagesp.yml`.

- Roda **09:00 e 16:00 (horário de Brasília), de segunda a sexta**.
- Após cada coleta, os arquivos da pasta `dados` são atualizados no próprio repositório —
  é de lá que você baixa o `cotacoes_pescado.xlsx`.
- A planilha também fica anexada a cada execução, na aba **Actions**, por 90 dias.
- Dá para disparar na hora pela aba **Actions → Coleta CEAGESP - Pescado → Run workflow**.
  O campo *Parametros extras* aceita as mesmas opções da linha de comando — deixe vazio
  para a coleta normal, ou use `--todas` / `--tudo` / `--data 03/08/2026`.
- É gratuito nessa escala (cada execução leva menos de um minuto).

Pontos a saber:

- O GitHub pode atrasar execuções agendadas em alguns minutos nos horários de pico. Sem
  impacto aqui — o boletim do dia continua sendo coletado.
- Se o repositório ficar **60 dias sem nenhuma atividade sua**, o GitHub desativa o
  agendamento e avisa por e-mail; basta reativar com um clique na aba Actions.
- Se um dia a CEAGESP bloquear as requisições vindas de datacenter, a execução falha e o
  GitHub envia e-mail. Nesse caso, o plano B é o agendamento local descrito abaixo.
- O site da CEAGESP às vezes não responde. O robô tenta 3 vezes antes de desistir e falhar
  com e-mail. Não há perda: a coleta seguinte recupera o boletim, já que o site mantém
  cerca de 7 deles.
- A publicação no GitHub Pages é um passo separado, que já falhou por conta própria
  (fila travada em 06/08/2026, deixando o site com dados velhos enquanto a coleta constava
  como bem-sucedida). Por isso o workflow agora confere se o Pages publicou o commit da
  coleta; se não publicar, ele pede um novo build e, persistindo, **falha de propósito**
  para que o e-mail de erro chegue em vez do problema passar despercebido.

## Alternativa: execução automática no Windows

Serve como plano B se a coleta na nuvem parar de funcionar. Abra o PowerShell na pasta do
projeto e rode (não precisa de administrador — as tarefas ficam no seu usuário):

```powershell
.\agendar_windows.ps1
```

Isso cria duas tarefas, às 09:00 e às 16:00, de segunda a sexta, apontando para
`executar_robo.bat`. Se o PC estiver desligado no horário, a tarefa roda assim que ele for
ligado. Para outro horário, `.\agendar_windows.ps1 -Horarios 07:30`; para desfazer,
`.\agendar_windows.ps1 -Remover`.

Rodar duas vezes ao dia não gera duplicidade: a segunda execução só busca boletim que
ainda não estiver no histórico.

## Requisitos

Python 3 com `requests` e `openpyxl` — instale com `pip install -r requirements.txt`.

## Observações técnicas

- O certificado TLS do site já apresentou problema de validação. O robô tenta primeiro com
  verificação normal e, se falhar, refaz a requisição sem verificar o certificado e registra
  um **aviso** no log. São dados públicos de preço, mas vale saber que isso acontece.
- O robô depende do HTML atual do site. Se a CEAGESP mudar o layout, ele para com erro
  claro no log (`layout do site mudou?`) em vez de gravar dados errados.
- Entre uma consulta e outra há uma pausa de 2 segundos para não sobrecarregar o servidor.
