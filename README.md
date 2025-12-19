# 🧩 Guia Universal — Desbloquear Planilha Excel Protegida (XLSX / XLSM)

Este tutorial ensina **como remover proteção de planilhas do Excel** usando **Prompt de Comando / PowerShell**, funcionando em **qualquer computador Windows**, **qualquer local do arquivo** e **qualquer planilha `.xlsx` ou `.xlsm`**.

> ⚠️ **Importante:**
> - Este método **NÃO quebra senha para abrir o arquivo**.
> - Funciona apenas para **“Planilha protegida”** (edição bloqueada).
> - Método legal apenas para arquivos **de sua autoria ou com autorização**.

---

## 📌 Quando este método funciona?

✔ Arquivo abre normalmente no Excel
✔ Aparece mensagem **“Planilha protegida”**
✔ Não pede senha ao abrir o arquivo

❌ Não funciona para:
- Senha para abrir o arquivo
- Projeto VBA protegido por senha

---

## 🛠️ O que vamos fazer (conceito)

Arquivos `.xlsx` e `.xlsm` são **ZIPs disfarçados**. Vamos:
1. Renomear para `.zip`
2. Extrair os arquivos internos
3. Remover a tag `<sheetProtection />`
4. Compactar novamente
5. Renomear de volta para Excel

---

## 🧪 Exemplo genérico (serve para QUALQUER arquivo)

Neste guia, usaremos um nome **exemplo**:
```
ARQUIVO.xlsx
```
👉 Sempre **substitua pelo nome real do seu arquivo**.

---

## 🚀 PASSO A PASSO UNIVERSAL

### 1️⃣ Abrir o PowerShell

Pressione:
```
Win + R → powershell
```

---

### 2️⃣ Ir até a pasta onde o arquivo está

Exemplos:

**Área de Trabalho**
```powershell
cd $env:USERPROFILE\Desktop
```

**Outra pasta (exemplo):**
```powershell
cd "C:\Caminho\Da\Pasta"
```

💡 Dica: você pode segurar **Shift + botão direito** na pasta → *Abrir no Terminal*

---

### 3️⃣ Renomear o arquivo para ZIP

```powershell
Rename-Item "ARQUIVO.xlsx" "ARQUIVO.zip"
```

> Funciona igual para `.xlsm`

---

### 4️⃣ Extrair o conteúdo

```powershell
Expand-Archive "ARQUIVO.zip" temp
```

✔ Se a pasta `temp` foi criada, está tudo certo

---

### 5️⃣ Remover a proteção das planilhas

```powershell
Get-ChildItem temp\xl\worksheets\*.xml |
ForEach-Object {
    (Get-Content $_) -replace '<sheetProtection[^>]*/>', '' |
    Set-Content $_
}
```

✔ Remove a proteção de **todas as abas**

---

### 6️⃣ Recompactar o arquivo

```powershell
Compress-Archive temp\* "ARQUIVO_DESBLOQUEADO.zip"
```

---

### 7️⃣ Renomear de volta para Excel

```powershell
Rename-Item "ARQUIVO_DESBLOQUEADO.zip" "ARQUIVO_DESBLOQUEADO.xlsx"
```

(ou `.xlsm`, conforme o arquivo original)

---

### 8️⃣ (Opcional) Limpar arquivos temporários

```powershell
Remove-Item temp -Recurse -Force
```

---

## ✅ Resultado Final

Você terá um novo arquivo:
```
ARQUIVO_DESBLOQUEADO.xlsx
```

✔ Planilhas editáveis
✔ Fórmulas preservadas
✔ Macros preservadas (XLSM)

---

## ⚠️ Avisos importantes

- Se o Excel disser:
  > “O arquivo foi reparado”

  👉 Clique em **SIM** (normal)

- Se o nome do arquivo tiver espaços, **use aspas**

- Não abra o arquivo enquanto estiver como `.zip`

---

## 🧠 Dicas Avançadas

🔹 Funciona em **lote** (vários arquivos)
🔹 Pode ser automatizado em **script `.ps1`**
🔹 Ideal para TI, contabilidade, fiscal e auditoria

---

## 📎 Conclusão

Este é um **método coringa**, rápido e confiável para desbloquear planilhas protegidas do Excel **sem usar programas externos**.

Use com responsabilidade ✅

