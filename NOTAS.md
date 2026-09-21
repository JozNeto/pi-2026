# Notas — retomada do Projeto Integrador IV (SolarSync)

Contexto rápido para retomar de onde paramos.

## O que mudou

O projeto anterior (Energia x População SP) foi **totalmente descartado** —
toda a estrutura (docs/specs/src/tests/reports/data/.venv) foi apagada.
Novo projeto: **SolarSync**, evolução de um sistema IoT de monitoramento de
geração solar fotovoltaica que já existe e funciona.

## Proposta nova (fonte)

`Projeto Integrador IV  Sistema de Monitoramento de Geração S.pdf` — nome de
arquivo tem espaços duplos/caracteres especiais, o Read direto por path
falhou; funcionou copiando via glob no shell (`for f in *.pdf; do cp "$f"
copia.pdf; done`) antes de ler.

Esse PDF é a documentação técnica do sistema já construído, **não** um
enunciado formal de disciplina — não há rubrica/prazo/entregáveis
explícitos ali (confirmado com o usuário: não existe outro PDF com o
enunciado oficial).

## Arquitetura já existente (fora do nosso escopo — tratar como dado)

ESP32 WROOM32E + MAX485 (RS485/Modbus) lendo o inversor Anenji → API
Node.js → **MongoDB Atlas** → dashboard React + Vite. Alertas hoje são
baseados em regra fixa (ex.: tensão AC fora de 212–230V), não em ML.

## Escopo do NOSSO trabalho (confirmado com o usuário)

Só a camada de **aprendizagem de máquina / análise preditiva**, consumindo
os dados já armazenados no MongoDB Atlas. Não vamos mexer no
firmware/API/frontend existentes.

## Direção de ML decidida (confirmado com o usuário)

**Núcleo do projeto:**
1. **Forecasting de geração solar** — prever potência FV nas próximas
   horas/dia a partir da série histórica (candidatos: Prophet, gradient
   boosting com features de lag, LSTM se o volume justificar).
2. **Detecção de anomalias / manutenção preditiva** — substituir os
   alertas por regra fixa por um modelo que aprende o padrão normal
   (Isolation Forest, autoencoder, ou z-score adaptativo por período do
   dia) e sinaliza desvios.

**Extensões possíveis (não confirmadas, avaliar depois de ver os dados
reais):**
3. Previsão de autonomia da bateria.
4. Monitoramento de eficiência/degradação (geração real vs. esperada).

## Acesso aos dados (desbloqueado em 2026-09-20)

Banco `anenji_monitor` (MongoDB Atlas, usuário SOMENTE LEITURA — regra do
usuário: nunca executar nada que altere o banco; só find/count/aggregate
sem `$out`/`$merge`). A connection string NÃO é gravada em arquivo: usar via
variável de ambiente `MONGO_URI` (pedir ao usuário de novo se preciso).
Não mexer na coleção `users` (contém hashes de senha).

Coleção de telemetria: **`telemetries`** (`telemetry` e
`system.buckets.telemetry` existem mas estão vazias).

Perfil real (2026-09-20): 105.096 registros, 2026-09-06 12:42 → 2026-09-20
22:58 (~14,5 dias, ainda coletando), ~1 leitura a cada 11,7 s (~7,4k/dia).
Campos: timestamp(datetime), pvVoltage, pvCurrent, pvPower, acInputVoltage,
acInputFrequency, acOutputVoltage, acOutputFrequency, batteryVoltage,
batteryCurrent(+In/Out), batteryPower, batteryWattsIn, batterySOC, loadWatts,
loadVA, loadCurrent, inverterLoadPercent, inverterMode (SBU 59%, SUB 41%,
GRID 4 regs), inverterStatus (sempre ONLINE — sem falhas rotuladas).

Achados que afetam o projeto:
- Só ~2 semanas de histórico: sem sazonalidade além da diária; forecasting
  deve ser de curto prazo (horas/dia seguinte). Registrar como limitação.
- Sem rótulos de anomalia (status sempre ONLINE) → detecção NÃO supervisionada;
  avaliação via eventos naturais: 3.389 regs com acInputVoltage<100 (queda de
  rede), 38 regs loadWatts>3000, 6 regs batteryVoltage<40, 6 lacunas >600 s.
- Nenhum dado meteorológico no banco.
- FUSO RESOLVIDO: timestamps do Mongo estão em UTC (o último registro tem
  ~3 s de diferença do horário UTC da extração). Convertidos para UTC−3;
  geração vai de ~06h a ~17h locais, pico médio às 10h (coerente com sol).
- Achado: com bateria cheia (SOC≥99%) o inversor limita o pvPower à carga
  (ex.: 19/09) → pvPower medido NÃO é geração potencial (censura).

## Estado do trabalho (2026-09-20) — V1 do Relatório Parcial

Estrutura enxuta (sem specs/TDD formais, por prazo): `src/` (scripts),
`data/raw` + `data/processed` (git-ignored), `reports/` (JSONs, figuras,
`final/Relatorio_Parcial_V1_SolarSync.docx`). `.venv` no projeto.
Pipeline (rodar nesta ordem, com `MONGO_URI` no ambiente só p/ a 1ª etapa):
`python -m src.export_telemetry` → `src.eda` → `src.forecast` →
`src.anomaly` → `src.fig_arquitetura` → `src.build_relatorio`.
Resultados V1 (base 105.166 regs, extraída 20/09): previsão em 15 min com
walk-forward em 8 dias de teste (13–20/09): 15 min → persistência vence;
1 h → Gradient Boosting vence (MAE ~257 W vs 332 W persistência, ganho ~22%,
8/8 dias); 3 h e dia seguinte → climatologia vence. Anomalias (Isolation
Forest, 1%) sinalizam picos de carga; sobreposição quase nula com a regra
atual (tensão fora de 212–230 V).
DOCX: python-docx + Word COM (script de atualização de campos/PDF no
scratchpad da sessão). Placeholders amarelos no docx: [Curso], [Polo(s)],
[Cidade]. Sumário/listas são campos do Word (atualizar com F9 se editar).
Refs verificadas via web: Inman et al. 2013, Bergmeir & Benítez 2012, Lei
14.300 (ementa), FPP3. Modbus spec sem URL (link 404). Demais refs clássicas
(Friedman 2001, Liu 2008, Pedregosa 2011, Hoerl 1970, Hyndman 2006,
Chandola 2009, Hastie 2009, Ke 2017, McKinney 2010) citadas de memória —
conferir antes da entrega.

## Próximos passos (Quinzenas 5–7)

1. Dados meteorológicos abertos como exógenas (3 h / dia seguinte).
2. Tratar censura por bateria cheia (flag de SOC≥99% / modelo por regime).
3. Mais histórico (coleta segue) + hiperparâmetros com validação aninhada.
4. Definir "anomalia" com Adriano, rotular eventos, comparar variantes.
5. Integrar resultados ao dashboard React (endpoint de leitura).
6. Relatório Final + Avaliação Colaborativa + Vídeo (vence 06/11).
