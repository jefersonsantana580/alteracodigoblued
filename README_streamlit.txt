# 📊 Sugestão Automática de Troca de Código – Streamlit

Este aplicativo Streamlit executa a alocação automática de códigos de produção
com base em deltas positivos e negativos, utilizando **saldo mestre por código**,
respeitando mês e modelo, garantindo que **nenhum incremento ultrapasse o TOTAL permitido**.

---

## ✅ Objetivo

- Substituir códigos com **delta negativo** por códigos com **delta positivo**
- Controlar corretamente o **saldo total disponível por PRODUCT**
- Criar **incrementos reais** apenas quando houver saldo remanescente
- Marcar **cortes explícitos** quando não existir saldo possível
- Gerar um Excel final **consistente e auditável**

---

## 📂 Estrutura do Excel de Entrada

O arquivo Excel deve conter **duas abas obrigatórias**:

### Aba `FILAS`

Campos importantes:
- **Coluna A**: `Blued` (mês / data)
- `NR_FILA`
- `PRODUCT`
- `PRODUCT SERIES`
- **Coluna I**: `PRODUCT PROPOSTO` (pode estar vazia)

### Aba `Delta Alteração`

- Uma linha por `PRODUCT`
- Colunas mensais no formato `jan/26`, `fev/26`, `mar/26`, etc.
- Valores:
  - Positivos → incremento
  - Negativos → corte

---

## 🧠 Regras de Negócio Implementadas

### 1) Saldo mestre por código

- Soma todos os **deltas positivos** por PRODUCT
- Cada unidade de saldo pode ser usada **uma única vez**

### 2) Substituição de deltas negativos

Para cada delta negativo:
1. Prioridade: **mesmo mês e mesmo modelo**
2. Depois: **outro modelo no mesmo mês**
3. Sem saldo disponível: `PRODUCT PROPOSTO = corte`

Cada substituição consome **1 unidade do saldo mestre**.

### 3) Incrementos puros

Após processar **todos os deltas negativos**:
- Se ainda houver saldo positivo:
  - Criar novas filas
  - `NR_FILA = incremento 1, incremento 2, ...`
  - `PRODUCT PROPOSTO = código com saldo`

Garantia:
```
substituições + incrementos = soma(delta positivo)
```

---

## 📤 Estrutura do Excel de Saída

### Aba `FILAS`
- Linhas originais + linhas de incremento
- Coluna I (`PRODUCT PROPOSTO`) sempre preenchida com:
  - código sugerido
  - ou `corte`
- Incrementos identificados pela coluna `NR_FILA`

### Aba `Delta Alteração`
- Mantida sem alterações (auditoria)

---

## ▶️ Execução

### Via Streamlit
```bash
streamlit run app.py
```

### Via script Python
```bash
python saldo_mestre_sugestao_filas.py
```

---

## ✅ Garantias

- Não há duplicação de saldo
- Incrementos nunca excedem o TOTAL
- Todas as decisões são determinísticas e rastreáveis
- Excel original não é alterado

---

✅ Solução pronta para uso mensal recorrente em produção.