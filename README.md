# 🧾 Comparador de Notas – Posto vs PMM

Este script em Python foi desenvolvido para **automatizar a conferência de notas fiscais entre o posto de combustível e o sistema interno da Prefeitura**, já que **a cada fechamento de quinzena costumam ocorrer grandes divergências nos valores registrados**.  

O objetivo é facilitar a auditoria, economizar tempo e reduzir erros manuais, garantindo que todas as informações estejam consistentes entre as planilhas do **POSTO** e da **PMM**.

---

## ⚙️ Funcionalidades

O programa:
- 📂 Lê duas planilhas Excel (`POSTO.xlsx` e `PMM.xlsx`);
- 🔍 Compara os registros de notas com base na coluna `Numero_Nota`;
- 🚨 Identifica e exibe divergências detalhadas sobre:
  - Notas que aparecem **apenas em uma das planilhas**;
  - Diferenças nas **datas de emissão**;
  - Divergências no **tipo de combustível**;
  - Variações nos **valores de litragem**;
- ✅ Exibe um relatório claro diretamente no terminal com todas as inconsistências encontradas.

---

## 🗂️ Estrutura esperada das planilhas

As planilhas devem conter as seguintes colunas:

| Coluna            | Descrição                              |
|--------------------|----------------------------------------|
| `Data`             | Data da emissão da nota                |
| `Numero_Nota`      | Número da nota fiscal (chave de comparação) |
| `Tipo_Combustivel` | Tipo de combustível (ex: Gasolina, Diesel) |
| `Litragem`         | Quantidade de litros abastecidos       |
| `Setor`            | Setor responsável pelo abastecimento   |

> ⚠️ Certifique-se de que os nomes das colunas correspondem exatamente aos definidos no script.

---

## 💻 Como usar

1. Coloque os arquivos `POSTO.xlsx` e `PMM.xlsx` na mesma pasta do script `comparador_de_notas.py`.  
2. Abra um terminal nessa pasta.  
3. Execute o comando:

```bash
python comparador_de_notas.py
```
O resultado será exibido no terminal, indicando notas faltantes ou divergentes, como no exemplo:

```yaml
✔ Planilha 'POSTO.xlsx' carregada com 120 registros.
✔ Planilha 'PMM.xlsx' carregada com 118 registros.

--- INICIANDO ANÁLISE DE DIVERGÊNCIAS ---

[AVISO] Notas encontradas APENAS na Planilha POSTO.xlsx:
  - Data: 05/10/2025 | Nota: 0123 | Litragem: 45.0

[ERRO] Notas com DATAS divergentes:
  - Nota: 0145 (Setor: Transporte) | Planilha POSTO.xlsx: 10/10/2025 vs Planilha PMM.xlsx: 11/10/2025

--- ANÁLISE CONCLUÍDA ---
```

🧠 Dependências

O programa utiliza as bibliotecas:

pandas

numpy

openpyxl (para leitura de arquivos Excel)
🧩 Configurações internas

No início do script, é possível alterar os nomes dos arquivos e colunas conforme sua necessidade:
```python
ARQUIVO_A = 'planilha_a.xlsx'
ARQUIVO_B = 'planilha_b.xlsx'

COLUNA_DATA = 'Data'
COLUNA_NOTA = 'Numero_Nota'
COLUNA_TIPO = 'Tipo_Combustivel'
COLUNA_VALOR = 'Litragem'
COLUNA_SETOR = 'Setor'
```

🧠 Dependências

O programa utiliza as bibliotecas:
- pandas
- numpy
- openpyxl

Instale-as com:
```python
pip install pandas numpy openpyxl]
```
