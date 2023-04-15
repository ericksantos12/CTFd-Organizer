<!-- <p align="center">
  <a href="" rel="noopener">
 <img width=200px height=200px src="https://i.imgur.com/6wj0hh6.jpg" alt="Project logo"></a>
</p> -->

<h1 align="center">🚩 CTFd Organizer</h1>

<div align="center">

  [![Status](https://img.shields.io/badge/status-active-success.svg)]() 
  [![GitHub Issues](https://img.shields.io/github/issues/ericksantos12/CTF-Organizer.svg)](https://github.com/ericksantos12/CTF-Organizer/issues)
  [![GitHub Pull Requests](https://img.shields.io/github/issues-pr/ericksantos12/CTF-Organizer.svg)](https://github.com/ericksantos12/CTF-Organizer/pulls)
  [![Gitmoji](https://img.shields.io/badge/gitmoji-%20😜%20😍-FFDD67.svg?style=flat-square)](https://gitmoji.dev)
</div>

<div align='center'>
  <img width=800 src=https://files.catbox.moe/megb4g.gif>
</div>

## 📝 Table of Contents

- [🧐 About](#-about)
- [🏁 Getting Started](#-getting-started)
  - [Prerequisites](#prerequisites)
  - [Installing](#installing)
- [🎈 Usage](#-usage)
- [⛏️ Built Using](#️-built-using)
- [✍️ Authors](#️-authors)

## 🧐 About

Este projeto lê dados de arquivos CSV, manipula e combina as informações usando a biblioteca pandas, pivota e reorganiza os dados, e salva o resultado final em um arquivo Excel. O objetivo é transformar informações sobre desafios e usuários que completaram esses desafios em uma tabela que mostra quais desafios cada usuário completou.

## 🏁 Getting Started

Este programa requer a instalação de algumas bibliotecas Python. Siga as instruções abaixo para instalar e executar o programa.

### Prerequisites

- Python 3.7 ou superior instalado
- pip (gerenciador de pacotes do Python) instalado


### Installing

1. Clone o repositório em seu computador
```bash
git clone https://github.com/ericksantos12/CTFd-Organizer.git
```

2. Acesse o diretório e instale as dependencias necessárias
```bash
cd CTFd-Organizer
pip install -r requirements.txt
```

## 🎈 Usage

1. Antes de executar o programa, coloque os arquivos csv em uma pasta `csv` no diretório do projeto.

Os arquivos devem ser renomeados para os nomes `challenges.csv`, `users.csv` e `solves.csv`.

1. Execute o programa
```bash
python app.py
```

Ao executar o programa, . O arquivo resultante será salvo na pasta `result` no diretório do projeto.

```bash
Make sure the files are in "./csv" folder

Do you want to continue? Output file will be overwritten [y/n] (y): 
```

## ⛏️ Built Using

- [Python](https://www.python.org/)
- [Pandas](https://pandas.pydata.org/)
- [Rich](https://rich.readthedocs.io/)

## ✍️ Authors

- [@michaelsantosti](https://github.com/michaelsantosti) - Idea
- [@ericksantos12](https://github.com/ericksantos12) - Code & Documentation
