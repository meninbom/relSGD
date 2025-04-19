![Python](https://img.shields.io/badge/python-3.10+-blue)
---

# 📊 Automação de Relatórios - SGD

Este projeto automatiza o processo de geração de relatórios do sistema **SGD**, desde o login e download de arquivos até o processamento, visualização e exportação dos dados. Tudo isso por meio de uma interface moderna feita com `customtkinter`, combinando `Selenium`, `pandas`, `openpyxl` e `matplotlib` para entregar relatórios com tabelas dinâmicas e gráficos (barras e pizza).

---

## 🚀 Funcionalidades

- **🔐 Automação Web**  
  Login automático no SGD, seleção de datas e download do relatório em Excel.
  
- **📈 Processamento de Dados**  
  Geração de tabelas dinâmicas e gráficos (barras em `A3`, pizza em `N3`).

- **🖥 Interface Gráfica**  
  Interface amigável para inserir credenciais, definir períodos e acompanhar logs em tempo real.

- **📤 Exportações**  
  Salva relatórios como PDF (simulado), CSV e imagens dos gráficos.

- **📊 Comparação de Períodos**  
  Permite comparar dois períodos distintos, por técnico ou status.

- **🧾 Logging Completo**  
  Todas as etapas são registradas em `automacao_relatorios.log` e exibidas na interface.

---

## 🧰 Requisitos

- **Sistema Operacional:** Windows 10/11 (64-bit)  
- **Python:** 3.10 ou superior  
- **Navegador:** Google Chrome (versão 135.x ou compatível) ou Brave  
- **Dependências:** Listadas em `requirements.txt`

---

## ⚙️ Instalação

### 1. Clone o repositório
```bash
git clone https://seu-repositorio.git
cd automacao-relatorios-sgd
```

### 2. Crie o ambiente virtual
```bash
python -m venv venv
.\venv\Scripts\activate  # Para Windows
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Configure o navegador
- Verifique sua versão do Chrome em `chrome://settings/help`  
- O `webdriver-manager` irá baixar automaticamente o ChromeDriver compatível

---

## 🛠 Configuração

### 🔗 URLs do SGD
Edite a função `executar_script()` no `relSGD.py` e substitua os links fictícios:
```python
driver.get("https://sgd.dominiosistemas.com.br")  # Login
driver.get("https://sgd.dominiosistemas.com.br/programacoes.html")  # Relatórios
```

### 🧩 XPaths
Verifique os XPaths na função `executar_script()`:
```python
user_input = driver.find_element(By.XPATH, '/html/body/div/form/input[2]')
```
Adapte de acordo com a estrutura real do HTML (de preferência, use IDs ou seletores CSS).

### 🗂 Permissões
Garanta que o diretório de download (ex: `~/Downloads`) tenha permissão de escrita. Rode o script como administrador, se necessário.

---

## ▶️ Como Usar

### 1. Execute o script
```bash
python relSGD.py
```

### 2. Preencha os campos na interface
- **Usuário e Senha:** suas credenciais do SGD  
- **Navegador:** "Chrome" ou "Brave"  
- **Download:** pasta onde os relatórios serão salvos  
- **Período:** datas no formato `ddmmaa` (ex: `010125` = 01/01/2025)  
- **Salvar Configurações:** grava em `config.json`

### 3. Clique em **"▶️ Executar"**
- O script fará login, baixará e processará o relatório.
- Gera um arquivo renomeado como `010125_310125.xlsx`.
- O progresso aparece na interface e é salvo no log.

---

## 📌 Funcionalidades Futuras (em desenvolvimento)

- Visualização interativa dos gráficos
- Exportação direta para PDF
- Comparação entre múltiplos períodos
- Melhorias visuais e responsividade da interface

---

Se quiser, posso gerar uma versão com emojis reduzidos ou um toque mais técnico e formal — só avisar!
