# Robô CEAGESP — Cotações de Pescado (Atacado)

Coleta os preços de atacado dos pescados de interesse em <https://ceagesp.gov.br/cotacoes/>
(categoria **PESCADOS**, Entreposto da Capital) e acumula o histórico.

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

Colunas: `data`, `categoria`, `produto`, `classificacao`, `unidade_peso`, `menor`, `comum`,
`maior`, `quilo`, `coletado_em`. **Menor / Comum / Maior** são os preços em R$ —
"Comum" é o valor mais praticado.

Log de execução em `logs/robo.log`.

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

## Alternativa: execução automática no Windows

`executar_robo.bat` é o ponto de entrada para o Agendador de Tarefas do Windows.
O caminho do Python está fixo no `.bat`; se o Python for reinstalado em outro lugar,
ajuste essa linha.

Para ativar o agendamento **às 09:00 e às 16:00, de segunda a sexta**, rode estes dois
comandos no PowerShell (não precisa de administrador — a tarefa fica no seu usuário):

```powershell
schtasks /Create /TN "Robo CEAGESP Pescado - manha" /TR "'C:\Users\thiago.IEA\Desktop\robo-ceagesp\executar_robo.bat'" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 09:00 /F
```

```powershell
schtasks /Create /TN "Robo CEAGESP Pescado - tarde" /TR "'C:\Users\thiago.IEA\Desktop\robo-ceagesp\executar_robo.bat'" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 16:00 /F
```

Conferir, rodar na hora ou remover:

```powershell
schtasks /Query /TN "Robo CEAGESP Pescado - manha" /V /FO LIST
```

```powershell
schtasks /Run /TN "Robo CEAGESP Pescado - manha"
```

```powershell
schtasks /Delete /TN "Robo CEAGESP Pescado - manha" /F
```

Rodar duas vezes ao dia não gera duplicidade: a segunda execução só busca boletim que
ainda não estiver no histórico. O computador precisa estar ligado e com seu usuário
logado no horário.

## Requisitos

Python 3 com `requests` e `openpyxl` (ambos já instalados nesta máquina).

## Observações técnicas

- O certificado TLS do site já apresentou problema de validação. O robô tenta primeiro com
  verificação normal e, se falhar, refaz a requisição sem verificar o certificado e registra
  um **aviso** no log. São dados públicos de preço, mas vale saber que isso acontece.
- O robô depende do HTML atual do site. Se a CEAGESP mudar o layout, ele para com erro
  claro no log (`layout do site mudou?`) em vez de gravar dados errados.
- Entre uma consulta e outra há uma pausa de 2 segundos para não sobrecarregar o servidor.
