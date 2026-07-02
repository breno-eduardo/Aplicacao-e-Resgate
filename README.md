# Gerador Template Aplicação e Resgate GPR - Web

Versão web em Flask do gerador de XLSX.

## Rodar localmente

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abra: http://localhost:8000

## Deploy no Coolify

1. Suba este projeto para o GitHub.
2. No Coolify, crie uma Application.
3. Conecte o repositório.
4. Use Dockerfile.
5. Porta interna: 8000.
6. Configure o domínio/subdomínio.
