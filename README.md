# 🚚 Comparador de Notas – Sistema de Auditoria de Abastecimentos

Um **dashboard completo em Streamlit** para auditar planilhas de abastecimento da Prefeitura (PMM) e do Posto, identificando divergências, inconsistências, consumo por secretaria e gerando relatórios profissionais em PDF e Excel.

## 🧾 Funcionalidades Principais

### 🔍 1. Comparação entre planilhas POSTO x PMM
- Detecta diferenças de litragem  
- Identifica notas presentes só em uma planilha  
- Detecta tipos de combustível divergentes  
- Verifica datas inconsistentes  

### ⚠️ 2. Identificação automática de problemas
- Notas duplicadas  
- Litragem negativa  
- Notas repetidas no mesmo dia  
- Mesma nota aparecendo em dias diferentes  
- Itens que aparecem somente em uma planilha  

### 🏛️ 3. Consumo por Secretaria
- Gasolina (POSTO / PMM)  
- Diesel (POSTO / PMM)  
- Comparação lado a lado  
- Filtro por setor  

### 📊 4. Resumo Executivo
Mostra rapidamente:
- Totais de litros  
- Totais de gasolina e diesel  
- Quantidade de divergências  

### 📥 5. Download de Relatórios
- Excel completo  
- PDF profissional  

## 📁 Estrutura do Projeto
```
comparador-notas/
├── app.py                
├── README.md             
├── POSTO.xlsx            
├── PMM.xlsx              
└── requirements.txt      
```

## 🛠️ Tecnologias
- Python
- Streamlit
- Pandas
- Altair
- ReportLab
- OpenPyXL

## 🚀 Como Executar
```
pip install -r requirements.txt
streamlit run app.py
```
