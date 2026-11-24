# 📚 LudosPro Training Export Automation

![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-2EAD33?logo=playwright&logoColor=white)
![Word](https://img.shields.io/badge/Word-python--docx-2B579A)
![RPA](https://img.shields.io/badge/RPA-automação_web-6E40C9)

Robô em Python com Playwright que lê os treinamentos cadastrados na plataforma LudosPro e monta um único documento Word com todo o conteúdo: curso, módulos, atividades e o texto de cada slide.

## 📌 Sobre o projeto

A plataforma não exporta o conteúdo dos cursos. Para revisar ou reaproveitar o material, era preciso abrir curso por curso, módulo por módulo e cada slideshow em modo de edição, copiando o texto à mão.

O robô faz esse caminho sozinho e só lê. Toda tela de edição que ele abre é fechada com **Cancelar**, então nada é salvo na plataforma.

## 🚀 Funcionalidades

- Lista todos os cursos (96 por página, com paginação);
- Para cada curso, lê título e descrição;
- Entra em cada módulo e atividade e abre as mecânicas do tipo **Slideshow**, lendo título, introdução e título e texto de cada slide. As demais mecânicas (Quiz, This or That etc.) são ignoradas;
- Gera `saida/Treinamentos_LudosPro.docx`, com uma página por treinamento e títulos por módulo, atividade e slideshow.

## 🧠 Decisões de projeto

- **Somente leitura:** o robô abre a edição só para ler e sempre sai por **Cancelar** ou pelo botão de fechar.
- **Cache para retomar:** cada curso extraído vira `saida/cache/<id>.json`. Se a execução parar no meio, a próxima continua de onde parou. `--refazer` ignora o cache.
- **Espera pela tela certa:** a plataforma é uma SPA que mostra por um instante o curso anterior. O robô espera as requisições terminarem, o texto parar de mudar e o título bater com o da lista antes de ler.
- **Word salvo a cada curso:** se o Word estiver aberto, salva uma cópia `_novo` ao lado em vez de falhar.
- **Login guardado:** a sessão fica num perfil próprio do Chrome (`.chrome_ludos/`), fora do Git. Com a sessão salva, dá para rodar sem janela (`--oculto`).

## 🛠️ Tecnologias utilizadas

- Python
- Playwright
- python-docx

## ⚙️ Como usar

1. Instale as dependências:
   ```
   pip install -r requirements.txt
   playwright install chrome
   ```
2. Copie `.env.example` para `.env` e preencha o endereço do painel. Usuário e senha são opcionais: sem eles, faça o login na janela do Chrome e tecle Enter no terminal.
3. Rode:
   ```
   python Src/ludos_treinamentos.py              # todos os cursos
   python Src/ludos_treinamentos.py 138 137      # só esses cursos
   python Src/ludos_treinamentos.py --refazer    # ignora o cache
   python Src/ludos_treinamentos.py --oculto     # sem janela (exige login já salvo)
   ```

## 📂 Estrutura do projeto

ludospro-training-export-automation

- Src
  - ludos_treinamentos.py
- README.md
- requirements.txt
- .gitignore
- .env.example
