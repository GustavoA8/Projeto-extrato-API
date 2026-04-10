# 💼 Gerador de Relatórios Bancários

Aplicação desktop desenvolvida em Python para leitura, processamento e geração de relatórios bancários automatizados a partir de extratos em Excel.

---

## 🚀 Funcionalidades

* 📄 Leitura automática de extratos bancários (Excel)
* 🏦 Suporte a múltiplos bancos:

  * Itaú
  * Santander
  * Bradesco
* 📊 Geração de relatório estruturado com:

  * Resumo financeiro
  * Créditos e Débitos
  * Tarifas
  * Aplicações por grupo
* 🧠 Classificação inteligente de lançamentos por grupo
* 🖥️ Interface gráfica moderna com CustomTkinter
* 📂 Gerenciamento de grupos de aplicações via Excel

---

## 🖼️ Interface

Interface simples, moderna e intuitiva:

* Seleção de banco
* Upload do extrato
* Geração com um clique
* Tela de gerenciamento de grupos

---

## 🛠️ Tecnologias Utilizadas

* Python 3.x
* Pandas
* OpenPyXL
* CustomTkinter
* Tkinter
* Pillow

---

## 📁 Estrutura do Projeto

```
Projeto-extrato/
│
├── app.py                  # Interface principal
├── core/                  # Lógica de parsing e agrupamento
├── services/              # Processamento principal
├── banks/                 # Regras específicas por banco
├── utils/                 # Funções auxiliares
├── img/                   # Imagens (logos)
├── Excel/                 # Arquivos de entrada
├── grupo-aplic.xlsx       # Mapeamento de grupos
└── requirements.txt
```

---

## ⚙️ Instalação

Clone o repositório:

```bash
git clone https://github.com/GustavoA8/Projeto-extrato-API.git
cd Projeto-extrato-API
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

---

## ▶️ Como Executar

```bash
python app.py
```

---

## 📌 Como Usar

1. Selecione o banco
2. Escolha o arquivo de extrato (.xls ou .xlsx)
3. (Se necessário) informe o nome do condomínio
4. Clique em **GERAR RELATÓRIO**
5. O arquivo será gerado automaticamente na mesma pasta

---

## 🧠 Diferenciais do Projeto

* Separação modular por banco
* Código organizado em camadas (core, services, banks)
* Sistema de agrupamento flexível via Excel
* Interface amigável (não depende de terminal)
* Pronto para virar produto interno/empresa

---

## 👨‍💻 Autor

**Gustavo Araujo de Souza**

* GitHub: https://github.com/GustavoA8

---

## 📄 Licença

Este projeto está sob a licença MIT.
