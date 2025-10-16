# Importa as bibliotecas necessárias
import pandas as pd
import requests
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# Função para geocodificar um endereço usando a API do Google Maps
def geocode_address(address, api_key):
    """
    Envia um endereço para a API de Geocodificação do Google,
    retorna a latitude, longitude e o status da requisição.
    """
    base_url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": address,
        "key": api_key,
        "language": "pt-BR"
    }
    try:
        # Faz a requisição GET para a API
        response = requests.get(base_url, params=params)
        response.raise_for_status()  # Lança um erro para respostas com status ruim (4xx ou 5xx)
        data = response.json()
        
        # Verifica se a API retornou um resultado válido
        if data["status"] == "OK":
            location = data["results"][0]["geometry"]["location"]
            return location["lat"], location["lng"], "OK"
        else:
            # Retorna o status de erro da API (ex: ZERO_RESULTS)
            return None, None, data['status']
            
    except requests.exceptions.RequestException as e:
        # Trata erros de conexão com a API
        return None, None, f"ERRO_CONEXAO: {e}"
    except Exception as e:
        # Trata outros erros inesperados
        return None, None, f"ERRO_INESPERADO: {e}"

# Classe principal da aplicação de interface gráfica (GUI)
class GeocoderApp:
    # Método construtor, onde a interface é inicializada
    def __init__(self, root):
        self.root = root
        self.root.title("Geocodificador de Endereços")
        self.root.geometry("500x350")

        # Cria o frame principal que conterá todos os widgets
        self.frame = tk.Frame(root, padx=10, pady=10)
        self.frame.pack(fill="both", expand=True)

        # --- Widgets da Interface ---
        
        # Campo para a chave da API
        tk.Label(self.frame, text="Chave da API do Google:").grid(row=0, column=0, sticky="w", pady=2)
        self.api_key_entry = tk.Entry(self.frame, width=60)
        self.api_key_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)

        # Campo para o caminho do arquivo e botão de procurar
        tk.Label(self.frame, text="Planilha de Endereços (.xlsx):").grid(row=2, column=0, sticky="w", pady=2)
        self.file_path_entry = tk.Entry(self.frame, width=50)
        self.file_path_entry.grid(row=3, column=0, sticky="ew", pady=2)
        self.browse_button = tk.Button(self.frame, text="Procurar...", command=self.browse_file)
        self.browse_button.grid(row=3, column=1, sticky="ew", padx=(5,0))
        
        # Botão para iniciar o processo
        self.start_button = tk.Button(self.frame, text="Iniciar Geocodificação", command=self.start_geocoding_thread, bg="green", fg="white", height=2)
        self.start_button.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(15,5))

        # Barra de progresso
        self.progress_bar = ttk.Progressbar(self.frame, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5)

        # Rótulo para exibir o status atual do processo
        self.status_label = tk.Label(self.frame, text="Pronto para começar.")
        self.status_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)
        
        # Configura a coluna para expandir com a janela
        self.frame.grid_columnconfigure(0, weight=1)

    # Abre uma janela para o usuário selecionar o arquivo .xlsx
    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
        if filepath:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, filepath)

    # Valida as entradas e inicia o processo de geocodificação em uma nova thread
    def start_geocoding_thread(self):
        api_key = self.api_key_entry.get()
        input_file = self.file_path_entry.get()

        # Validação simples para garantir que os campos não estão vazios
        if not api_key or not input_file:
            messagebox.showerror("Erro", "Por favor, insira a Chave da API e selecione um arquivo.")
            return

        # Desabilita o botão para evitar múltiplos cliques
        self.start_button.config(state="disabled")
        # Inicia a função principal em uma thread para não travar a interface
        threading.Thread(target=self.run_geocoding, args=(api_key, input_file), daemon=True).start()

    # Função principal que executa a geocodificação
    def run_geocoding(self, api_key, input_file):
        try:
            self.status_label.config(text="Lendo a planilha...")
            df = pd.read_excel(input_file, dtype=str).fillna('')
            total_rows = len(df)
            self.progress_bar["maximum"] = total_rows

            latitudes, longitudes, status_list = [], [], []

            # Itera sobre cada linha da planilha
            for index, row in df.iterrows():
                # Atualiza a interface com o progresso
                self.status_label.config(text=f"Processando linha {index + 1} de {total_rows}...")
                self.progress_bar["value"] = index + 1
                self.root.update_idletasks()

                # Monta o endereço completo a partir das colunas da planilha
                partes_endereco = [
                    row.get("Endereço", ""), row.get("Numero", ""), row.get("Bairro", ""),
                    row.get("Cidade", ""), row.get("UF", ""), row.get("CEP", "")
                ]
                address_completo = ", ".join(filter(None, partes_endereco))

                # Geocodifica o endereço se ele não estiver vazio
                if not address_completo:
                    lat, lng, status = None, None, "VAZIO"
                else:
                    lat, lng, status = geocode_address(address_completo, api_key)
                    time.sleep(0.1)  # Pequeno delay para não exceder os limites da API

                latitudes.append(lat)
                longitudes.append(lng)
                status_list.append(status)

            # Adiciona os resultados como novas colunas no DataFrame
            df['Latitude'] = latitudes
            df['Longitude'] = longitudes
            df['Status_Geocodificacao'] = status_list

            # Salva o DataFrame em um novo arquivo Excel
            output_filename = os.path.splitext(input_file)[0] + "_geocodificado.xlsx"
            df.to_excel(output_filename, index=False)
            
            messagebox.showinfo("Sucesso!", f"Processo concluído!\nArquivo salvo como:\n{output_filename}")

        except Exception as e:
            messagebox.showerror("Erro Durante o Processo", f"Ocorreu um erro: {e}")
        finally:
            # Garante que a interface seja reativada ao final do processo
            self.start_button.config(state="normal")
            self.status_label.config(text="Processo finalizado.")
            self.progress_bar["value"] = 0

# Ponto de entrada da aplicação
if __name__ == "__main__":
    root = tk.Tk()  # Cria a janela principal
    app = GeocoderApp(root)  # Instancia a classe da aplicação
    root.mainloop()  # Inicia o loop de eventos da interface