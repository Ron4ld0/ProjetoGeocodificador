import pandas as pd
import requests
import time
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def geocode_address(address, api_key):
    base_url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": address,
        "key": api_key,
        "language": "pt-BR"
    }
    try:
        response = requests.get(base_url, params=params)
        response.raise_for_status()
        data = response.json()
        if data["status"] == "OK":
            location = data["results"][0]["geometry"]["location"]
            return location["lat"], location["lng"], "OK"
        else:
            return None, None, data['status']
    except requests.exceptions.RequestException as e:
        return None, None, f"ERRO_CONEXAO: {e}"
    except Exception as e:
        return None, None, f"ERRO_INESPERADO: {e}"

class GeocoderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Geocodificador de Endereços")
        self.root.geometry("500x350")

        self.frame = tk.Frame(root, padx=10, pady=10)
        self.frame.pack(fill="both", expand=True)

        tk.Label(self.frame, text="Chave da API do Google:").grid(row=0, column=0, sticky="w", pady=2)
        self.api_key_entry = tk.Entry(self.frame, width=60)
        self.api_key_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)

        tk.Label(self.frame, text="Planilha de Endereços (.xlsx):").grid(row=2, column=0, sticky="w", pady=2)
        self.file_path_entry = tk.Entry(self.frame, width=50)
        self.file_path_entry.grid(row=3, column=0, sticky="ew", pady=2)
        self.browse_button = tk.Button(self.frame, text="Procurar...", command=self.browse_file)
        self.browse_button.grid(row=3, column=1, sticky="ew", padx=(5,0))
        
        self.start_button = tk.Button(self.frame, text="Iniciar Geocodificação", command=self.start_geocoding_thread, bg="green", fg="white", height=2)
        self.start_button.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(15,5))

        self.progress_bar = ttk.Progressbar(self.frame, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5)

        self.status_label = tk.Label(self.frame, text="Pronto para começar.")
        self.status_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)
        
        self.frame.grid_columnconfigure(0, weight=1)

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
        if filepath:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, filepath)

    def start_geocoding_thread(self):
        api_key = self.api_key_entry.get()
        input_file = self.file_path_entry.get()

        if not api_key or not input_file:
            messagebox.showerror("Erro", "Por favor, insira a Chave da API e selecione um arquivo.")
            return

        self.start_button.config(state="disabled")
        threading.Thread(target=self.run_geocoding, args=(api_key, input_file), daemon=True).start()

    def run_geocoding(self, api_key, input_file):
        try:
            self.status_label.config(text="Lendo a planilha...")
            df = pd.read_excel(input_file, dtype=str).fillna('')
            total_rows = len(df)
            self.progress_bar["maximum"] = total_rows

            latitudes, longitudes, status_list = [], [], []

            for index, row in df.iterrows():
                self.status_label.config(text=f"Processando linha {index + 1} de {total_rows}...")
                self.progress_bar["value"] = index + 1
                self.root.update_idletasks()

                partes_endereco = [
                    row.get("Endereço", ""), row.get("Numero", ""), row.get("Bairro", ""),
                    row.get("Cidade", ""), row.get("UF", ""), row.get("CEP", "")
                ]
                address_completo = ", ".join(filter(None, partes_endereco))

                if not address_completo:
                    lat, lng, status = None, None, "VAZIO"
                else:
                    lat, lng, status = geocode_address(address_completo, api_key)
                    time.sleep(0.1)

                latitudes.append(lat)
                longitudes.append(lng)
                status_list.append(status)

            df['Latitude'] = latitudes
            df['Longitude'] = longitudes
            df['Status_Geocodificacao'] = status_list

            output_filename = os.path.splitext(input_file)[0] + "_geocodificado.xlsx"
            df.to_excel(output_filename, index=False)
            
            messagebox.showinfo("Sucesso!", f"Processo concluído!\nArquivo salvo como:\n{output_filename}")

        except Exception as e:
            messagebox.showerror("Erro Durante o Processo", f"Ocorreu um erro: {e}")
        finally:
            self.start_button.config(state="normal")
            self.status_label.config(text="Processo finalizado.")
            self.progress_bar["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = GeocoderApp(root)
    root.mainloop()
