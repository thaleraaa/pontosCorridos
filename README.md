# pontosCorridos
Gerenciamento de tabela de pontos corridos desenvolvido em Excel com VBA.

---

## Sobre o Projeto
Aplicação desenvolvida inteiramente em VBA dentro do Excel para gerenciar campeonatos no formato de pontos corridos. O projeto conta com interface visual própria e lógica dinâmica de classificação.

---

## Funcionalidades

### 🏠 Home
- Botão para criar novo campeonato
- Modal de configuração solicitando nome do campeonato e quantidade de times (2 a 20, apenas números pares)

### 📊 Tabela de Classificação
Gerada automaticamente ao criar o campeonato, com todos os times e pontuação zerada. Critérios de classificação:

| Critério | Descrição |
|----------|-----------|
| PTS | Pontos |
| J | Jogos disputados |
| V | Vitórias |
| E | Empates |
| D | Derrotas |
| GP | Gols pró |
| GC | Gols contra |
| SG | Saldo de gols |

### 🔄 Rodadas Dinâmicas
- Exibe apenas os jogos da rodada atual, evitando poluição visual
- Atualização automática da tabela ao registrar resultados
- Edição de resultados já registrados com reação correta da tabela

---

## Tecnologias
- Microsoft Excel
- VBA (Visual Basic for Applications)
