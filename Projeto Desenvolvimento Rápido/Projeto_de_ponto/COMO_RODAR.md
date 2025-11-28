# 📘 GUIA COMPLETO - GERADOR DE FOLHA DE PONTO

## 🎯 O QUE É ESTE PROGRAMA?

Este é um **Gerador de Folha de Ponto** que cria PDFs automaticamente para funcionários, respeitando:
- ✅ Escalas de trabalho (5X1, 5X2, 6X1, 12X36)
- ✅ Feriados nacionais e locais
- ✅ Folgas manuais
- ✅ Data de admissão
- ✅ Organização por posto/ano/funcionário

---

## � O QUE O PROGRAMA USA (BIBLIOTECAS)

O programa importa várias bibliotecas Python. Veja o que cada uma faz:

### ✅ Bibliotecas Nativas (já vêm com Python - não precisa instalar):
```python
import os              # Gerencia pastas e arquivos
import io              # Trabalha com dados em memória
import json            # Salva/carrega configurações em formato JSON
import base64          # Converte logo da empresa
import traceback       # Mostra erros detalhados
from datetime import datetime, timedelta, date  # Trabalha com datas
from typing import Dict, Any, List, Optional, Tuple  # Ajuda na organização do código
import unicodedata     # Remove acentos de textos
```

### 📦 Bibliotecas EXTERNAS (precisa instalar via pip):

```python
# GUI (Interface Gráfica)
from tkinter import ...  # JÁ VEM COM PYTHON - Interface gráfica
from tkcalendar import DateEntry  # ⚠️ INSTALAR - Calendário na tela

# PDF
from reportlab.lib.pagesizes import A4  # ⚠️ INSTALAR - Gera PDFs
from reportlab.pdfgen import canvas     # ⚠️ INSTALAR
from reportlab.lib.units import mm      # ⚠️ INSTALAR
from reportlab.lib.utils import ImageReader  # ⚠️ INSTALAR
from reportlab.lib.colors import black  # ⚠️ INSTALAR

# Planilhas Excel
import pandas as pd  # ⚠️ INSTALAR - Lê arquivos .xlsx/.xls/.csv
```

### 📋 RESUMO: O que você PRECISA instalar:
- **pandas** → Lê planilhas Excel/CSV
- **tkcalendar** → Mostra calendário na interface
- **reportlab** → Gera os PDFs
- **openpyxl** → Ajuda o pandas a ler arquivos .xlsx modernos

---

## 🚀 INSTALAÇÃO RÁPIDA (RECOMENDADO)

### Método Automático - Usando Arquivo .BAT

A forma mais fácil de instalar e usar o programa:

#### 🔧 Primeira Instalação:

1. **Dê duplo clique no arquivo:** `INSTALADOR_COMPLETO.bat`
   - Ele irá:
     - ✅ Baixar e instalar o Python automaticamente
     - ✅ Criar o ambiente virtual (.venv)
     - ✅ Instalar todas as bibliotecas necessárias
   - **Tempo estimado:** 5-10 minutos (primeira vez)
   - **Requer:** Conexão com a internet

2. **Aguarde a mensagem:** "Instalação concluída com sucesso!"

3. **Pronto!** Agora é só usar o programa

#### ▶️ Para Usar o Programa (Sempre):

**Dê duplo clique no arquivo:** `ABRIR_PROGRAMA.bat`
- O programa abrirá instantaneamente!
- Use este arquivo sempre que quiser abrir o programa

> 💡 **DICA:** Crie um atalho do `ABRIR_PROGRAMA.bat` na área de trabalho para acesso rápido!

---

## 🖥️ INSTALAÇÃO MANUAL (Avançado)

Se preferir instalar manualmente ou já tem Python instalado:

### PASSO 1: Instalar Python

1. **Baixe Python 3.10 ou superior:**
   - https://www.python.org/downloads/
   - ✅ **IMPORTANTE:** Marque "Add Python to PATH" durante instalação

### PASSO 2: Criar Ambiente Virtual e Instalar Bibliotecas

Abra o PowerShell na pasta do programa e execute:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install pandas tkcalendar reportlab openpyxl
```

### PASSO 3: Rodar o Programa

```powershell
python gerador_ponto.py
```



---

## 📋 COMO USAR O PROGRAMA

### 4.1 - CONFIGURAR PERÍODO

1. **Defina as datas:**
   - **Início:** Data de início do período (ex: 01/01/2025)
   - **Fim:** Data de fim do período (ex: 31/01/2025)
   - **1ª Folga:** Data da primeira folga do ciclo (importante para escalas 5X2, 12X36, etc)

2. **Escolha o tipo de escala:**
   - **5X1:** 5 dias trabalha, 1 folga
   - **5X2:** 5 dias trabalha, 2 folgas
   - **6X1 (FIXO):** 6 dias trabalha, domingo folga
   - **6X1 (INTERCALADA):** 6 dias trabalha, folga alternada
   - **12X36:** 12 horas trabalha, 36 horas folga

### 4.2 - CADASTRAR FERIADOS

1. Clique no botão **"🗓️ FERIADOS"**
2. **Para adicionar feriado:**
   - Escolha a data no calendário
   - Digite o nome (ex: "NATAL", "ANO NOVO")
   - Escolha o tipo:
     - **NACIONAL:** Todos os funcionários folgam
     - **LOCAL:** Apenas funcionários de postos específicos folgam
   - Se for LOCAL, selecione a cidade
3. Clique em **"ADICIONAR"**

### 4.3 - CADASTRAR CIDADES E POSTOS

1. Clique no botão **"🏙️ CIDADES"**
2. **Para adicionar cidade:**
   - Clique em **"ADICIONAR CIDADE"**
   - Digite o nome da cidade (ex: "RECIFE", "OLINDA")
3. **Para vincular postos à cidade:**
   - **Dê duplo clique na cidade** OU selecione e clique em "EDITAR POSTOS"
   - Marque os postos que pertencem àquela cidade
   - Clique em **"SALVAR"**
   - **Você verá TODOS os postos já cadastrados**, mesmo os de planilhas antigas!

> **💡 DICA:** Feriados locais só afetam postos da cidade selecionada!

### 4.4 - CARREGAR PLANILHA DE FUNCIONÁRIOS

1. Clique no botão **"📊 CARREGAR"**
2. **Escolha uma opção no menu:**

   **OPÇÃO 1: 📂 CARREGAR ARQUIVO**
   - Selecione sua planilha Excel (.xlsx, .xls ou .csv)
   - A planilha pode conter as seguintes colunas:
     - **NOME** (obrigatório)
     - **CPF**
     - **MATRÍCULA** ou **MATRICULA**
     - **FUNÇÃO** ou **FUNCAO** ou **CARGO**
     - **POSTO**
     - **ADMISSÃO** ou **ADMISSAO** (formato: DD/MM/AAAA)
     - **FILIAL** (nome da empresa/filial)
     - **CNPJ** (da filial)
     - **ENDEREÇO** (da filial)
     - **CIDADE** (da filial)
   - Os funcionários aparecerão na lista

   **OPÇÃO 2: 📥 BAIXAR MODELO EXCEL**
   - Baixa um arquivo Excel modelo com:
     - ✅ **TODAS as 10 colunas** já criadas
     - ✅ 3 exemplos completos de funcionários
     - ✅ Formato correto das datas (DD/MM/AAAA)
     - ✅ Exemplos de CNPJ, endereço, cidade formatados
   - **IMPORTANTE:** Apague os 3 exemplos e insira seus dados reais
   - Use este modelo para garantir que está tudo correto!

> **💡 DICA:** Se os dados da FILIAL, CNPJ, ENDEREÇO e CIDADE forem iguais para todos os funcionários, você pode preencher apenas a primeira linha e copiar para as demais!

> **🔒 IMPORTANTE:** Ao carregar uma nova planilha, os postos e cidades anteriores são **PRESERVADOS**! Isso significa:
> - ✅ Postos de planilhas antigas continuam disponíveis
> - ✅ Cidades cadastradas não são perdidas
> - ✅ Vínculos entre cidades e postos são mantidos
> - ✅ Você pode alternar entre diferentes planilhas sem perder configurações

### 4.5 - ADICIONAR OCORRÊNCIAS (FOLGAS/FERIADOS EXTRAS)

Se algum funcionário teve uma folga ou feriado específico:

1. Selecione o funcionário na lista
2. Clique com botão direito → **"GERENCIAR OCORRÊNCIAS"**
3. Selecione os dias e marque como:
   - **FOLGA:** Dia de folga do funcionário
   - **FERIADO:** Feriado específico
4. Clique em **"SALVAR"**

### 4.6 - CONFIGURAR LOGO DA EMPRESA

O logo aparecerá no cabeçalho de todos os PDFs gerados.

1. Clique no botão **"🖼️ LOGO"**
2. **Selecione a imagem do logo:**
   - Formatos aceitos: .png, .jpg, .jpeg, .bmp, .gif
   - Tamanho máximo: 5MB
   - Recomendado: Fundo transparente (.png)
3. O logo será salvo automaticamente

> **💡 DICA:** Se você já tinha um arquivo `logo.txt`, o programa continuará usando-o caso não selecione um novo logo pelo botão.

### 4.7 - SALVAR CONFIGURAÇÕES

**IMPORTANTE:** Sempre salve antes de gerar os PDFs!

1. Clique no botão **"💾 SALVAR"**
2. Isso salva:
   - ✅ Feriados cadastrados
   - ✅ Cidades e postos vinculados
   - ✅ Ocorrências dos funcionários
   - ✅ Período e escala configurados

### 4.8 - GERAR OS PDFs

1. Clique no botão **"📄 GERAR PDFs"**
2. **Preencha as informações da empresa:**
   - Nome da Filial
   - CNPJ
   - Endereço
   - Cidade
3. Clique em **"GERAR"**
4. Aguarde o processamento

**Os PDFs serão salvos em:**
```
Pontos Gerados/
  └── [NOME_DO_POSTO]/
      └── [ANO]/
          └── [NOME_FUNCIONARIO]/
              └── 01.2025_NOME_FUNCIONARIO.pdf
```

---

## 🗂️ ESTRUTURA DE ARQUIVOS

```
📁 NOVA PONTO - Backup/
  ├── 📄 gerador_ponto.py          ← Programa principal
  ├── 📄 escalas_store.json        ← Dados salvos (criado automaticamente)
  ├── 📄 logo.txt                  ← Logo da empresa (base64 - OPCIONAL)
  ├── 📄 COMO_RODAR.md            ← Este guia
  └── 📁 Pontos Gerados/           ← PDFs gerados (criado automaticamente)
      └── 📁 [POSTOS]/
          └── 📁 [ANOS]/
              └── 📁 [FUNCIONARIOS]/
                  └── 📄 XX.AAAA_NOME.pdf
```

---

## 🔧 PROBLEMAS COMUNS E SOLUÇÕES

### ❌ Erro: "python não é reconhecido como comando"

**Causa:** Python não foi adicionado ao PATH do Windows

**Solução:**
1. Desinstale o Python atual (Painel de Controle → Programas)
2. Baixe novamente: https://www.python.org/downloads/
3. **Marque "Add Python to PATH"** na primeira tela da instalação
4. Clique em "Install Now"
5. Reinicie o computador
6. Teste no PowerShell: `python --version`

### ❌ Erro: "No module named 'pandas'" (ou reportlab, tkcalendar, openpyxl)

**Causa:** Bibliotecas não foram instaladas

**Solução:**

#### Se NÃO usa ambiente virtual:
```powershell
pip install pandas tkcalendar reportlab openpyxl
```

#### Se USA ambiente virtual:
```powershell
cd "C:\Users\andre.luis\Desktop\NOVA PONTO - Backup"
.\venv\Scripts\Activate.ps1
pip install pandas tkcalendar reportlab openpyxl
```

### ❌ Erro: "Activate.ps1 cannot be loaded because running scripts is disabled"

**Causa:** PowerShell bloqueia scripts por segurança

**Solução:**
1. Abra PowerShell **como Administrador** (botão direito → "Executar como Administrador")
2. Execute:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. Digite `S` e pressione Enter
4. Feche e abra o PowerShell normalmente
5. Tente ativar o venv novamente:
   ```powershell
   .\venv\Scripts\Activate.ps1
   ```

### ❌ Erro: "pip não é reconhecido como comando"

**Causa:** Python instalado sem pip ou PATH incorreto

**Solução:**
```powershell
python -m pip install --upgrade pip
python -m pip install pandas tkcalendar reportlab openpyxl
```

### ❌ PDFs não estão sendo gerados

**Causas possíveis:**
1. Você não clicou em **"💾 SALVAR"** antes de gerar
2. A planilha não tem a coluna NOME
3. Planilha com formato incorreto

**Solução:**
1. Sempre salve antes de gerar
2. **Use o modelo Excel:** Clique em CARREGAR → BAIXAR MODELO EXCEL
3. Preencha o modelo com seus dados
4. Todas as colunas opcionais podem ficar vazias, mas NOME é obrigatório

### ❌ Feriado local não está funcionando

**Solução:**
1. Certifique-se de que a cidade foi cadastrada em **"🏙️ CIDADES"**
2. Vincule os postos corretos à cidade (duplo clique na cidade)
3. No feriado, escolha tipo **"LOCAL"** e selecione a cidade
4. Salve tudo antes de gerar

### ❌ Escalas não estão respeitando as folgas

**Solução:**
1. Configure a data da **"1ª Folga"** corretamente
2. Para escala **12X36**, a 1ª Folga deve ser o primeiro dia que o funcionário NÃO trabalha
3. Salve e gere novamente

---

## � POSSO MOVER OU RENOMEAR A PASTA?

### ✅ SIM! Pode fazer as duas coisas sem quebrar o programa!

O aplicativo **NÃO depende** do nome ou localização da pasta. Todos os caminhos são relativos.

#### 🔄 Renomear a Pasta:

Você pode renomear de `NOVA PONTO - Backup` para qualquer nome (ex: `FOLHA_PONTO`, `Sistema_Ponto`, etc).

**O que muda:**
- Apenas atualize o caminho quando for rodar:
  ```powershell
  # Se renomear para "FOLHA_PONTO":
  cd "C:\Users\andre.luis\Desktop\FOLHA_PONTO"
  python gerador_ponto.py
  ```

**Dica:** Use nomes sem espaços e caracteres especiais para facilitar (ex: `FOLHA_PONTO` é melhor que `Nova Ponto - v2.0`)

---

#### 📁 Mover a Pasta:

Pode mover para **qualquer lugar**:
- ✅ Desktop → Documentos
- ✅ C:\ → D:\ (outro HD)
- ✅ Computador → Pendrive
- ✅ Local → Rede compartilhada
- ✅ PC Casa → PC Trabalho

**Exemplos válidos:**
```powershell
C:\Users\andre.luis\Documents\PONTO\
D:\Trabalho\Sistemas\PONTO\
E:\Pendrive\PONTO\
\\Servidor\Compartilhado\PONTO\
```

**O que você precisa fazer:**
1. Mova a pasta normalmente (arraste ou copie)
2. Atualize o comando `cd` para o novo caminho:
   ```powershell
   cd "D:\Trabalho\PONTO"
   python gerador_ponto.py
   ```

**⚠️ ATENÇÃO: Se usar Ambiente Virtual (venv):**

O ambiente virtual guarda caminhos absolutos. Ao mover a pasta:

1. **Delete a pasta `venv/`** antes de mover
2. **Mova a pasta** para o novo local
3. **Recrie o venv:**
   ```powershell
   cd "novo\caminho\PONTO"
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   pip install pandas tkcalendar reportlab openpyxl
   ```

**✅ Todos os seus dados são preservados!**
- `escalas_store.json` (configurações e logo)
- `logo.txt` (logo da empresa - se existir)
- `Pontos Gerados/` (PDFs gerados)

---

## �📱 COPIAR PARA OUTRA MÁQUINA

### Cenário 1: Instalação Simples (SEM ambiente virtual)

1. **Copie toda a pasta** para o novo computador
   - Pode ser por: Pendrive, Email, OneDrive, Google Drive, Rede, etc.
   - Inclui: `gerador_ponto.py`, `escalas_store.json`, `COMO_RODAR.md`
   - OPCIONAL: `logo.txt` (se não usar o botão LOGO)

2. **No novo computador, instale o Python** (ver Passo 1)
   - https://www.python.org/downloads/
   - ✅ Marque "Add Python to PATH"

3. **Instale as bibliotecas:**
   ```powershell
   cd "caminho\para\NOVA PONTO - Backup"
   pip install pandas tkcalendar reportlab openpyxl
   ```

4. **Execute:**
   ```powershell
   python gerador_ponto.py
   ```

**🎉 Pronto! O arquivo `escalas_store.json` já contém todas as suas configurações salvas!**

---

### Cenário 2: Com Ambiente Virtual (Recomendado)

1. **Copie APENAS os arquivos essenciais** (NÃO copie a pasta `venv/`):
   - `gerador_ponto.py`
   - `escalas_store.json`
   - `COMO_RODAR.md`
   - OPCIONAL: `logo.txt` (se existir)
   - Pasta `Pontos Gerados/` (se quiser manter PDFs antigos)

2. **No novo computador, instale o Python**

3. **Crie novo ambiente virtual:**
   ```powershell
   cd "caminho\para\NOVA PONTO - Backup"
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   pip install pandas tkcalendar reportlab openpyxl
   ```

4. **Execute:**
   ```powershell
   python gerador_ponto.py
   ```

> ⚠️ **IMPORTANTE:** NUNCA copie a pasta `venv/` para outra máquina! Sempre crie um novo ambiente virtual.

---

## 🆘 GUIA RÁPIDO DE COMANDOS

### Comandos Essenciais:

```powershell
# Navegar até a pasta
cd "C:\Users\andre.luis\Desktop\NOVA PONTO - Backup"

# Ver versão do Python
python --version

# Ver versão do pip
pip --version

# Listar bibliotecas instaladas
pip list

# Instalar bibliotecas
pip install pandas tkcalendar reportlab openpyxl

# Atualizar pip
python -m pip install --upgrade pip

# Rodar o programa
python gerador_ponto.py
```

### Comandos de Ambiente Virtual:

```powershell
# Criar ambiente virtual (só faz UMA VEZ)
python -m venv venv

# Ativar ambiente virtual (PowerShell)
.\venv\Scripts\Activate.ps1

# Ativar ambiente virtual (CMD)
venv\Scripts\activate.bat

# Desativar ambiente virtual
deactivate

# Deletar ambiente virtual (se der problema)
Remove-Item -Recurse -Force venv
```

---

## 📞 DICAS FINAIS

✅ **Sempre salve** antes de fechar o programa
✅ **Faça backup** do arquivo `escalas_store.json` regularmente
✅ **Teste com poucos funcionários** antes de gerar tudo
✅ **Verifique os PDFs** gerados antes de distribuir
✅ **Use o duplo clique** nas cidades para editar postos rapidamente
✅ **Baixe o modelo Excel** antes de criar sua planilha (botão CARREGAR → BAIXAR MODELO)
✅ **Configure o logo** uma única vez (botão 🖼️ LOGO) - ele fica salvo para sempre
✅ **Troque de planilha sem medo** - postos e cidades anteriores são preservados automaticamente

---

## 📝 VERSÃO

**Versão do Programa:** V.5.39.00

**Novidades desta versão:**
- 🖼️ Botão para selecionar logo da empresa (sem precisar converter para base64)
- 📥 Modelo Excel completo para download (10 colunas com exemplos)
- 📂 Menu ao clicar em CARREGAR (Carregar Arquivo ou Baixar Modelo)
- 📁 Pasta de saída renomeada para "Pontos Gerados"
- 🔒 **HISTÓRICO DE POSTOS:** Postos e cidades são preservados ao trocar de planilha

**Data deste Guia:** 17/Novembro/2024

---

**📧 Desenvolvido para facilitar a gestão de folhas de ponto!** 🚀
