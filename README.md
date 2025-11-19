# Bot Fiscalização

Automação GUI em Python para processar notas (obras), extrair DXF e listas técnicas e gerar relatórios Excel.

## Resumo
Este projeto fornece uma interface gráfica (Tkinter + ttkthemes) para inserir notas, controlar quantidade de postes, executar rotinas de extração (DXF / lista técnica) e gerar um relatório final em Excel.

## Arquivos principais
- [main.py](main.py) — ponto de entrada da aplicação e a interface gráfica (classe `Janela_principal`).
- [functions.py](functions.py) — implementa as rotinas de automação usadas pelo GUI, incluindo:
  - [`EntrarNota`](functions.py)
  - [`ExtrairDxf`](functions.py)
  - [`ExtrairLista`](functions.py)
  - [`fecharNota`](functions.py)

## Dependências
- Python 3.8+
- pandas
- ttkthemes
- tkinter (incluído na maioria das distribuições Python)

Instale dependências:
```bash
pip install pandas ttkthemes
```

## Como usar
1. Abra o repositório e execute a aplicação:
```bash
python main.py
```
2. Insira as notas no campo de texto e clique em "Iniciar".
3. Para cada nota será solicitado o número de postes e observações (ex.: "Sem lista técnica" ou "Bloqueada").
4. Ao final, gere o relatório em Excel pelo botão "Relatório".

## Estrutura do repositório
```
Bot Fiscalização.spec
coordenadas.json
functions.py
main.py
build/
...
```

## Observações
- As funções de automação estão em [functions.py](functions.py). Ajuste ou expanda essas funções conforme necessário.
- Para gerar o arquivo Excel a aplicação usa `pandas.DataFrame.to_excel`.

## Licença
Escolha uma licença apropriada (ex.: MIT) e adicione um arquivo `LICENSE`.
