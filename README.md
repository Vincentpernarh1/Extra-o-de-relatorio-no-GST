# GST Report Extraction

Automacao para extrair relatorios do GST e consolidar os arquivos em uma unica execucao.

## Resumo
- Fluxo unico: extracao + consolidacao.
- Credenciais em arquivo externo.
- Sem necessidade de ajustes manuais frequentes para execucao do navegador.

## Estrutura do Projeto
- `main.py`: executa a extracao e oferece consolidacao ao final.
- `main_organizar.py`: executa somente a consolidacao (uso opcional).
- `Credencial/usuario.json`: usuario e senha.
- `Downloads_Auxiliar/`: arquivos baixados.
- `Arquivos_Consolidados/`: arquivos finais.

## Como Atualizar Credenciais
1. Abra a pasta `Credencial`.
2. Abra o arquivo `usuario.json` com Bloco de Notas.
3. Atualize apenas os valores de `Usuario` e `Senha`.
4. Salve o arquivo.

Exemplo:

```json
{
  "Usuario": "SEU_USUARIO",
  "Senha": "SUA_SENHA"
}
```

## Build do Executavel (.exe)

### Pre-requisitos
```bash
pip install pyinstaller
playwright install chromium
```

Antes de gerar o build, confirme que o caminho abaixo existe (ajuste o numero de versao do chromium se for diferente):
```
C:\Users\<usuario>\AppData\Local\ms-playwright\chromium-1228\chrome-win64\chrome.exe
```
Se a versao for diferente, atualize tanto o comando abaixo quanto `get_playwright_browser_path()` em `main.py`.

### Comando de build
Execute na pasta do projeto:
```bat
pyinstaller --noconfirm --onefile --windowed --noconsole --name "GST Report Extraction" --icon "C:/Users/perna/Desktop/STALLANTIS/Extra-o-de-relatorio-no-GST/Credencial/icon.ico" --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1228\chrome-win64;ms-playwright\chromium-1228\chrome-win64" main.py
```

O executavel e gerado em `dist\GST Report Extraction.exe`.

### Distribuicao para outra maquina
Copie junto com o `.exe`:
```
GST Report Extraction.exe
Credencial/
  ├── usuario.json
  └── icon.ico
```
As pastas `Downloads_Auxiliar/` e `Arquivos_Consolidados/` sao criadas automaticamente na primeira execucao.

### Problemas comuns no build
- **"Chromium executable not found"**: a versao do chromium instalada localmente (`%LOCALAPPDATA%\ms-playwright\`) mudou. Atualize o numero de versao e o nome da pasta (`chrome-win` vs `chrome-win64`) no comando de build e em `get_playwright_browser_path()`.
- **.exe maior que 200MB**: esperado, o Chromium bundled ocupa ~150-200MB.

## Como Executar
### Opcao 1: Executavel
1. Execute `GST Report Extraction.exe`.
2. Confirme a mensagem para iniciar.
3. Aguarde o fim da extracao.
4. Confirme a consolidacao quando solicitado.

### Opcao 2: Python
```bash
py main.py
```

## Resultado Esperado
- 4 arquivos baixados em `Downloads_Auxiliar`.
- Arquivo(s) consolidado(s) em `Arquivos_Consolidados`.

## Troubleshooting (Simples)
### 1) Erro de login
- Verifique `Credencial/usuario.json`.
- Confirme se usuario e senha estao corretos.

### 2) Nao consolidou no final
- Verifique se os arquivos foram baixados em `Downloads_Auxiliar`.
- Rode novamente e confirme a consolidacao.
- Se necessario, rode apenas:
```bash
py main_organizar.py
```

### 3) Processo interrompido
- Feche o navegador aberto pelo processo.
- Execute novamente.
- Evite usar o navegador durante a automacao.

### 4) Sem arquivos no consolidado
- Confirme se existem arquivos `.xlsx` em `Downloads_Auxiliar`.
- Rode a consolidacao novamente.

## Boas Praticas de Operacao
- Mantenha conexao de rede estavel.
- Nao feche o navegador durante a execucao.
- Atualize credenciais somente no arquivo `Credencial/usuario.json`.

## Versao
- Versao atual: 3.0
- Ultima atualizacao: 2026-04-13
