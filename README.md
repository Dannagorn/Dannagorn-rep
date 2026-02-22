# Dannagorn-rep

Script de automação escolar (Megatron).

## Execução no Windows (PowerShell)

Se `python` não funcionar no seu Windows, use o launcher `py`:

```powershell
py --version
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m py_compile robo_siepe.py
py robo_siepe.py
```

## Arquivos principais

- `robo_siepe.py`: versão executável principal.
- `megatron_v6.py`: versão base/refatorada equivalente.
- `requirements.txt`: dependências Python.

## Observação importante

Se aparecer erro como `SyntaxError` com texto `diff --git` ou `index ...`, o arquivo `.py` foi contaminado com texto de patch.
Use o `robo_siepe.py` deste repositório (limpo) para executar.
