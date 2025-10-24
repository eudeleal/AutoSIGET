# 🧠 AUTOMA SIGET3000

Automatizador de digitação para o sistema **SIGET**, utilizado para preencher faixas horárias de linhas de ônibus a partir de um arquivo **CSV**.  
Este script lê as programações, organiza os blocos por linha/dia/sentido e executa o preenchimento automático via **PyAutoGUI**.

---

## 🚀 Funcionalidades

- Leitura automática de **planilhas CSV** com faixas horárias.
- Geração de **blocos agrupados por Linha, Dia e Sentido**.
- Automação da digitação no **SIGET** com controle via teclado.
- Sistema de **logs detalhados** com data e hora.
- Interface de linha de comando simples e interativa.
- Permite:
  - **[F10]** → Avançar  
  - **[F9]** → Repetir faixa ou voltar ao seletor  
  - **[F12]** → Encerrar execução a qualquer momento  

---

## ⚙️ Instalação

Clone este repositório e instale as dependências:

```bash
git clone https://github.com/eudeleal/automa-siget3000.git
cd automa-siget3000
pip install pyautogui keyboard tabulate
⚠️ Requisitos:

Python 3.8+

Permissões de automação ativas (no Windows, execute como administrador).

SIGET aberto e com a aba “Parâmetros de linha” pronta.

🧾 Estrutura do CSV
O arquivo .csv deve conter ponto e vírgula (;) como separador e possuir as colunas obrigatórias:

Copiar código
FaixaInicio; FaixaFinal; Intervalo; Percurso; TempTerm; Frota;
Linha; Dia; Sentido; Oso; LinhaOso
🧩 Exemplo:
yaml
Copiar código
FaixaInicio;FaixaFinal;Intervalo;Percurso;TempTerm;Frota;Linha;Dia;Sentido;Oso;LinhaOso
0440;0440;0;162;5;1;1001;Sab;0;4767-00;1001
0515;0530;15;170;5;3;1001;Sab;0;4767-01;100101
▶️ Como usar
Prepare o ambiente:

Deixe o SIGET aberto.

Vá até a aba “Parâmetros de linha”.

Insira a OSO da linha correta.

Apague as faixas existentes do Dia/Sentido a ser preenchido.

Posicione o cursor na primeira célula vazia da grade.

Execute o programa:

bash
Copiar código
python automa_siget3000.py
Informe o nome do arquivo CSV (sem a extensão).

Siga as instruções exibidas no terminal e use as teclas de controle (F10 / F9 / F12) conforme indicado.

🗂️ Estrutura de pastas
bash
Copiar código
automa-siget3000/
├── LOGs/
│   └── log_YYYY.MM.DD_HH.MM.SS.txt   # Arquivos de log automáticos
├── automa_siget3000.py               # Script principal
└── exemplo.csv                       # Exemplo de entrada
🧑‍💻 Autor
Eude Leal
📍 Salvador – BA, Brasil
🔗 github.com/eudeleal

📜 Licença
Este projeto é de uso interno e educativo.
Sinta-se livre para adaptar e aprimorar, mantendo os créditos originais.