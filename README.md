# MEGATRON v7.3 FINAL — Refatorado (MENU ORIGINAL PRESERVADO)

Este repositório agora contém uma versão **refatorada em código único** do projeto, com foco em UX TUI, robustez de automação Selenium e operação offline com fila de pendências.

## Arquivo principal

- `megatron_v73_final.py`

## Como executar (Windows)

1. Instale os pacotes já usados no projeto:
   - `openpyxl`
   - `prompt_toolkit`
   - `rich`
   - `keyboard`
   - `pywin32`
   - `selenium`
2. Garanta Chrome + ChromeDriver compatíveis no PATH.
3. Rode:

```bash
python megatron_v73_final.py
```

Na primeira execução, o sistema pedirá CPF/Senha e gravará em `meu_acesso.txt`.

---

## Interface e atalhos

### MENU PRINCIPAL (formato preservado)
- **F1..F7**: lançamento rápido por horário da grade
- **F8**: anotação avulsa (Excel)
- **F9**: configurar grade / importar SIEPE / sincronizar turmas-alunos / relatório semanal / fluxo linear unificado / rotina diária linear unificada
- **F10**: importar alunos da turma (ambiente de chamada) / lançamento manual (wizard)
- **F11**: viagem no tempo (data)
- **F12**: central da turma (checklist)
- **P**: pendências (revisar/enviar)
- **ESC**: sair

### Melhorias visuais adicionadas
- Destaque de **AGORA** e **PRÓXIMO horário**
- Resumo superior: `Hoje: X turmas | Y chamadas pendentes | Z alertas | 📤 Pendências: N`
- Status por atalho/horário com `✅` e `⏳`
- Bloqueio claro de chamadas em fim de semana/recesso

---

## Novas funcionalidades obrigatórias implementadas

1. **Wizard de chamada (5 etapas)**
   - turma → modo (setas+enter) → seleção checkbox → preview → execução
   - opções de voltar/cancelar/salvar pendência
   - sugestão de repetir última chamada no manual

2. **Seleção de alunos com checkbox + filtro incremental**
   - ↑↓ navega, espaço marca, enter confirma, ESC cancela
   - busca incremental por nome/apelido
   - fuzzy básico para tolerância de erro de digitação
   - sincroniza automaticamente a base de alunos a partir da página de chamada do SIEPE

3. **Pré-visualização com alertas**
   - alerta para 0 faltosos
   - alerta forte para >35% faltosos

4. **Modo offline + fila de pendências**
   - arquivo: `pendencias_megatron.json`
   - grava pendência em falha SIEPE e/ou dia não letivo
   - tela de pendências: reenviar uma/todas, ver detalhes, marcar resolvida

5. **Checklist como Central da Turma**
   - agrupado por turmas do dia
   - marcação concluída `☐ -> ✅`
   - tarefa com data e horário opcional
   - sugestão automática de caso crítico

6. **Excel mais amigável**
   - freeze panes em `B2` no Diário
   - legenda fixa `P/F/N`
   - atualização de aba `Resumo Semanal` por turma
   - comentários concatenados com separador + data/hora

7. **Importar SIEPE com preview**
   - preview de grade detectada antes de aplicar
   - aplicar manhã/tarde/tudo
   - mapeamento manual de inferência quando necessário
   - seleção de professor com TAB para montar grade real por docente
   - ajuste dinâmico dos 7 horários do menu a partir dos horários reais do professor
   - aviso de conflitos de alocação (mesmo slot com múltiplas turmas)

---

## Modo offline e reenvio de pendências

Quando a automação SIEPE falha:
1. o sistema oferece modal com opções:
   - tentar novamente
   - salvar pendência
   - cancelar
2. ao salvar pendência, ela entra em `pendencias_megatron.json`.
3. use atalho **P** no painel para:
   - reenviar uma
   - reenviar todas pendentes/falhas
   - ver detalhes
   - marcar resolvida manualmente

---

## Onde colar/substituir

- Substitua seu arquivo principal atual por `megatron_v73_final.py`.
- Mantenha os arquivos de dados existentes no mesmo diretório:
  - `horarios_megatron.json`
  - `lembretes_megatron.json`
  - `status_turmas_megatron.json`
  - `dump_quadros_cache.json`
  - `memoria_professores.json`
  - `meu_acesso.txt`
  - `banco_alunos_megatron.json` (opcional, para lista completa de alunos por turma)
- Novos arquivos criados automaticamente:
  - `pendencias_megatron.json`
  - `mapeamento_turmas_megatron.json` (cache de mapeamentos manuais na importação do SIEPE)


## Sincronização linear do SIEPE (manual)
- O robô detecta automaticamente quantas turmas o professor possui quando você escolher a opção de sincronização no menu F9.
- Há opção de manter o SIEPE aberto ao final da coleta para inspeção/depuração.
- A navegação agora usa um motor mais robusto com cache de iframe por contexto (`diario`, `frequencia`, `quadro`), waits por estado do DOM e fallback ordenado de seletores.
- A seleção de turma no Diário foi reforçada com estratégia por lista + fuzzy matching, reduzindo falhas por variação de texto no SIEPE.
- Se a coleta falhar, é gerado `debug_coleta_*.html` para diagnóstico do DOM.
- **Somente atualiza a lista de alunos quando você pedir** (modo "Turmas + alunos").
- O perfil detectado é salvo em `perfil_siepe_megatron.json`.


## Regras de turmas para coleta
- A coleta automática de turmas considera as turmas-alvo: **6A, 6B, 7A, 7B, 8A, 8B, 9A, 9B**.
- Na importação de horários por professor, o robô exibe também a carga semanal detectada por turma.
- O match turma↔quadro agora prioriza mapeamento por **ID do quadro** (evita preencher tudo com a mesma turma).
- A resolução de turma foi reforçada para inferência por **slot/célula** com fallback por maioria clara e logs estruturados de importação.
- A importação também lê o cabeçalho do detalhe (linha **Turma: ...**) para resolver a turma com mais precisão antes de qualquer fallback.
- Se houver conflito entre mapeamento manual por ID e a turma do detalhe (`Turma: ...`), o sistema prioriza a turma do detalhe e registra aviso.
- Na importação de grade, há etapa linear para **zerar mapeamentos antigos por ID** (recomendado) antes de aplicar a resolução por turma meta do detalhe.
- O mapeamento manual prioriza `id:{qid}` como chave única e estável para evitar colisões de label.


- Se o mapeamento manual por ID estiver dominado por uma única turma, o sistema alerta e oferece limpeza automática antes da importação.

- Em **F10**, a ação de importação de alunos aparece antes da chamada manual e permite escolher entre **Atualizar lista de alunos** ou **Manter lista atual**.

## Fluxo Linear Unificado (F9)
- Novo modo guiado em sequência: sincronizar turmas -> (opcional) sincronizar alunos -> importar grade do professor -> relatório semanal.
- Ideal para reduzir erro operacional no dia a dia.
- A opção **Rotina diária linear unificada** adiciona um pipeline único: sincronização opcional -> importação -> relatório.


## Onde ver as variáveis salvas
As variáveis de execução ficam em memória durante o uso, e os dados persistidos ficam em arquivos JSON/TXT no mesmo diretório do script:
- `horarios_megatron.json` → grade semanal
- `lembretes_megatron.json` → checklist/central da turma
- `status_turmas_megatron.json` → status de chamadas realizadas
- `pendencias_megatron.json` → fila offline de reenvio
- `mapeamento_turmas_megatron.json` → mapeamento manual por `id:{qid}`
- `perfil_siepe_megatron.json` → turmas detectadas/sincronização SIEPE
- `banco_alunos_megatron.json` → base local de alunos
- `dump_quadros_cache.json` → cache de quadros coletados
- `memoria_professores.json` → memória de professores detectados
- `meu_acesso.txt` → credenciais salvas localmente

Comandos rápidos para inspecionar:
- `python -m json.tool horarios_megatron.json`
- `python -m json.tool mapeamento_turmas_megatron.json`
- `python -m json.tool perfil_siepe_megatron.json`
