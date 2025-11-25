# 🧠 AUTOMA SIGET3000

Automatizador de digitação para o sistema **SIGET**, utilizado para preencher faixas horárias de linhas de ônibus a partir de um arquivo **CSV**.  
Este script lê as programações e executa o preenchimento automático via **PyAutoGUI**.

## 🚀 Funcionalidades

- Leitura de um **arquivo CSV** com faixas horárias e OSOs para impressão.
- Geração de **blocos agrupados por Linha, Dia e Sentido**.
- Automação da digitação no **SIGET** com controle via teclado.
- Interface simples e interativa. 

## 🧾 Estrutura do CSV
O arquivo .csv deve conter ponto e vírgula (;) como separador e possuir as colunas obrigatórias:

| FaixaInicio | FaixaFinal | Intervalo | Percurso | TempTerm |
|------------ | ---------- | --------- | -------- | -------- |

| Frota | Linha | Dia | Sentido | Oso | LinhaOso |
| ----- | ----- | --- | ------- | --- | -------- |




Utilize a planilha auxiliar para a criação do CSV corretamente 

## ▶️ Como usar
- Prepare o ambiente:
  - Abra o SIGET.
  - Vá até a aba “Parâmetros de linha”.
  - Insira a OSO da linha correta.
  - Apague as faixas existentes do Dia e/ou Sentido a ser preenchido.
  - Posicione o cursor na primeira célula vazia da grade.

- Execute o programa e siga as instruções exibidas no terminal e use as teclas de ação:
  - **[F10]** → Avançar  
  - **[F9]** → Repetir faixa/impressão ou voltar ao seletor  
  - **[F12]** → Encerrar execução a qualquer momento 

## 📜 Licença
Este projeto é de uso interno e educativo.

Sinta-se livre para adaptar e aprimorar, mantendo os créditos originais.

#### By Eude Leal, Salvador – BA, Brasil