# auto_relatorio

CLI em Go para exportar dados do MySQL (tabelas de experiência do paciente) para CSV e gerar um PowerPoint com gráficos de pizza por pergunta.

## Principais recursos

- Exporta CSV com `;` (Excel pt-BR) e BOM UTF-8 (acentos OK no Excel)
- Filtro de período por mês/ano (mês fechado) ou por início/fim (RFC3339)
- `--replace`: mapeia códigos 1..7 para texto (compatível com a macro VBA), exceto perguntas 16 e 20
- Remoção de duplicados consecutivos por paciente com tolerância de segundos (`--dedupe-sec`)
- `--pptx`: cria 1 slide por pergunta (exceto 16 e 20) com pizza + legenda

## Requisitos

- Go 1.22+
- Python 3.10+ (para montar o `.pptx` via `python-pptx`)
- Acesso ao MySQL (observação: alguns provedores exigem liberação de IP)

## Instalação

### 1) Clonar e compilar

```bash
go build .
```

### 2) Dependências do PowerPoint (uma vez)

Crie um ambiente virtual e instale as dependências Python:

```powershell
python -m venv .venv
\.venv\Scripts\python.exe -m pip install -U pip
\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Configuração (.env)

Crie um arquivo `.env` (não deve ser commitado) baseado em `.env.example`.

Opção recomendada (o programa monta o DSN sozinho):

- `MYSQL_HOST`
- `MYSQL_PORT`
- `MYSQL_DB`
- `MYSQL_USER`
- `MYSQL_PASS`

Alternativa:

- `MYSQL_DSN` (ou `--dsn`)

## Uso

### Exportar CSV (padrão: mês anterior fechado)

```powershell
./auto_relatorio.exe
```

### Exportar por mês/ano

```powershell
./auto_relatorio.exe --month=12 --year=2025 --replace
```

### Exportar por início/fim (RFC3339)

```powershell
./auto_relatorio.exe --start=2025-12-01T00:00:00-03:00 --end=2026-01-01T00:00:00-03:00
```

### Gerar PPTX automaticamente

Gera o CSV e, ao final, monta o PPTX e uma pasta com os PNGs:

```powershell
./auto_relatorio.exe --month=12 --year=2025 --replace --pptx=auto
```

Saídas esperadas:

- `relatorio_YYYY_MM.csv`
- `relatorio_YYYY_MM.pptx`
- `relatorio_YYYY_MM_png/manifest.json` + PNGs

### Gerar PPTX a partir de um CSV existente (sem banco)

```powershell
./auto_relatorio.exe --pptx-from=relatorio_2025_12.csv --pptx=relatorio_2025_12.pptx
```

## CSV (Excel)

- Separador: `;`
- BOM UTF-8 habilitado por padrão (`--bom=true`) para evitar `Ã£` no lugar de `ã`

## Publicando no GitHub

Arquivos sensíveis e gerados (como `.env`, `.venv/`, relatórios e PNGs) já estão cobertos por `.gitignore`.

Passo a passo típico:

```powershell
git init
git add .
git commit -m "Initial commit"

# crie o repositório no GitHub e depois:
git remote add origin https://github.com/SEU_USUARIO/SEU_REPO.git
git branch -M main
git push -u origin main
```

## Troubleshooting

- `Access denied` / não conecta pelo Go mas acessa via phpMyAdmin: verifique se o provedor exige liberação do IP da sua máquina para acesso remoto.


