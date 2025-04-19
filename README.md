![Python](https://img.shields.io/badge/python-3.10+-blue)
---

# 📊 Automação de Relatórios SGD

Automação completa para geração e análise de relatórios do sistema **SGD**, com download automático, tratamento em Excel, criação de **tabelas dinâmicas** e **gráficos interativos**, tudo através de uma interface moderna em Python.

🔗 **Repositório GitHub:** [github.com/meninbom/relSGD](https://github.com/meninbom/relSGD)

---

## ✨ Funcionalidades

- 🔒 **Login Automático** no sistema SGD via navegador controlado por Selenium.
- 📅 **Seleção de datas** e geração automática de relatórios Excel.
- 📊 **Tabelas Dinâmicas** com técnicos nas linhas e status nas colunas.
- 📈 **Gráficos:** Barra (em `A3`) e Pizza (em `N3`) diretamente no Excel.
- 🖥 **Interface Gráfica (GUI):** Criada com `customtkinter`.
- 📤 **Exportação:** para **CSV**, **PDF** simulado, e **imagens** dos gráficos.
- 📁 **Comparação de Períodos:** análise técnica ou situacional entre dois relatórios.
- 📝 **Logging completo** no terminal.
- ⚙️ **Executável funcional incluso** no diretório do projeto.

---

## 📁 Estrutura

```bash
relSGD/
├── relSGD.py          # Arquivo principal do projeto
├── README.md          # Este arquivo
├── requirements.txt   # Dependências
├── relSGD.exe         # Executável funcional da aplicação
```

---

## 🧩 Requisitos

- **SO:** Windows 10 ou 11 (64-bit)
- **Python:** 3.10+
- **Navegador:** Google Chrome ou Brave (versão compatível com seu ChromeDriver)
- **Bibliotecas:** listadas em `requirements.txt`

---

## 🚀 Como usar

### 1. Clone o repositório

```bash
git clone https://github.com/meninbom/relSGD.git
cd relSGD
```

### 2. Instale o ambiente (opcional, se for rodar o `.py`)

```bash
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Rode o sistema

Você pode:
- Executar o **relSGD.exe** diretamente, **ou**
- Rodar o script com Python:

```bash
python relSGD.py
```

---
