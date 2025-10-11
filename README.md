# Geocodificador de Endereços

Este é um aplicativo de desktop simples, desenvolvido em Python com uma interface gráfica (GUI) usando Tkinter. A principal função do aplicativo é converter uma lista de endereços de uma planilha Excel (.xlsx) em coordenadas geográficas (latitude e longitude) utilizando a API de Geocodificação do Google.

## Funcionalidades

- **Interface Gráfica Simples:** Uma janela fácil de usar para inserir a chave da API e selecionar o arquivo.
- **Processamento em Lote:** Lê uma planilha do Excel com múltiplos endereços.
- **Geocodificação:** Utiliza a API do Google Maps para obter a latitude e a longitude de cada endereço.
- **Feedback em Tempo Real:** Uma barra de progresso e um rótulo de status informam o andamento do processo.
- **Tratamento de Erros:** Captura e exibe status de erro da API (ex: `ZERO_RESULTS`, `REQUEST_DENIED`) e falhas de conexão.
- **Exportação de Resultados:** Salva uma nova planilha chamada `enderecos_geocodificados.xlsx` com os endereços originais e as novas colunas de latitude, longitude e status da geocodificação.

## Pré-requisitos

Antes de executar o aplicativo, você precisa ter o seguinte instalado:

- **Python 3.x**
- As seguintes bibliotecas Python:
  - `pandas`
  - `requests`
  - `openpyxl` (necessário para o pandas manipular arquivos `.xlsx`)
- Uma **Chave de API da Plataforma Google Maps** com a "API de Geocodificação" ativada. Você pode obter uma [aqui](https://developers.google.com/maps/documentation/geocoding/get-api-key).

## Como Instalar e Executar

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/Ron4ld0/ProjetoGeocodificador.git
    cd ProjetoGeocodificador
    ```

2.  **Instale as dependências:**
    ```bash
    pip install pandas requests openpyxl
    ```

3.  **Execute o aplicativo:**
    ```bash
    python geolocalizador.py
    ```

## Como Usar

1.  **Execute o script** `geolocalizador.py`.
2.  Na janela do aplicativo, **insira sua Chave da API do Google** no campo correspondente.
3.  **Clique em "Procurar..."** para selecionar a planilha Excel (`.xlsx`) que contém os endereços que você deseja geocodificar.
4.  **Clique no botão "Iniciar Geocodificação"**.
5.  Aguarde o processamento. A barra de progresso mostrará o andamento.
6.  Ao final, uma mensagem informará que o processo foi concluído. Um novo arquivo chamado `enderecos_geocodificados.xlsx` será criado no mesmo diretório do aplicativo, contendo os resultados.

## Formato da Planilha de Entrada

O aplicativo espera que a planilha de entrada contenha colunas que, juntas, formam o endereço completo. O script tentará ler as seguintes colunas (se existirem) para montar o endereço:

- `TIPO_LOGRADOURO` (Ex: Rua, Avenida)
- `LOGRADOURO` (Ex: Paulista)
- `NUMERO`
- `BAIRRO`
- `CIDADE`
- `ESTADO`
- `CEP`

Se sua planilha tiver nomes de coluna diferentes, você precisará ajustar o código no arquivo `geolocalizador.py` na função `run_geocoding`.

---
*Este projeto foi desenvolvido para automatizar o processo de geocodificação de endereços para análise de dados e mapeamento.*
