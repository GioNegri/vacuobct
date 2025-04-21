import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import matplotlib.pyplot as plt

def converter_pressao(valor):
    try:
        if isinstance(valor, str):
            match = re.search(r'-?\d+\.?\d*', valor.replace(",", "."))
            if match:
                return float(match.group())
        elif isinstance(valor, (int, float)):
            return float(valor)
    except:
        return None
    return None

def analisar_por_rack(df):
    df.columns = df.columns[:5]
    df = df.dropna(subset=[df.columns[0]])
    df[df.columns[0]] = pd.to_datetime(df[df.columns[0]])

    resultados = []
    tempo_total_segundos = (df[df.columns[0]].iloc[-1] - df[df.columns[0]].iloc[0]).total_seconds()

    for col in df.columns[1:]:
        df = df.sort_values(by=df.columns[0])
        df = df.drop_duplicates(subset=[df.columns[0]], keep="first")  # evita contar o mesmo minuto duas vezes

        df["diff_min"] = df[df.columns[0]].diff().dt.total_seconds() / 60
        df["diff_min"] = df["diff_min"].fillna(1)  # primeira linha vale 1 minuto

    for col in df.columns[1:-1]:  # ignora a última coluna "diff_min"
        df[col] = df[col].apply(converter_pressao)
        downtime_flags = df[col] < -5
        uptime_flags = df[col] >= -5

        tempo_downtime = df.loc[downtime_flags, "diff_min"].sum()
        tempo_uptime = df.loc[uptime_flags, "diff_min"].sum()
        tempo_total = tempo_downtime + tempo_uptime

        resultados.append({
            "rack": col,
            "uptime_pct": 100 * tempo_uptime / tempo_total if tempo_total else 0,
            "downtime_pct": 100 * tempo_downtime / tempo_total if tempo_total else 0,
            "tempo_uptime_min": tempo_uptime,
            "tempo_downtime_min": tempo_downtime
        })
    return resultados

def selecionar_arquivo():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if file_path:
        try:
            df = pd.read_excel(file_path)
            resultados = analisar_por_rack(df)

            resultado_texto = ""
            for r in resultados:
                resultado_texto += (
                    f"\n📊 {r['rack']}:\n"
                    f"   ❌ Downtime BCT: {r['uptime_pct']:.2f}% ({r['tempo_uptime_min']:.1f} minutos)\n"
                    f"   ✅ Uptime BCT: {r['downtime_pct']:.2f}% ({r['tempo_downtime_min']:.1f} minutos)\n"
                )

            messagebox.showinfo("Resultado por Rack", resultado_texto)
            
            plt.close('all')  #aqui e embaixo tem do except
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")
            
        # Criar gráfico de barras (colunas verticais)
        racks = [r['rack'] for r in resultados]
        uptimes = [r['uptime_pct'] for r in resultados]
        downtimes = [r['downtime_pct'] for r in resultados]
        uptimes_min = [r['tempo_uptime_min'] for r in resultados]
        downtimes_min = [r['tempo_downtime_min'] for r in resultados]
        
        x = range(len(racks))

        fig, ax = plt.subplots(figsize=(10,6))
        bars_uptime = ax.bar([i - 0.2 for i in x], uptimes, width=0.4,label='Downtime ❌', color='#FF5252' )
        bars_downtime = ax.bar([i + 0.2 for i in x], downtimes, width=0.4, label='Uptime ✅', color='#4CAF50')

        # Adiciona data labels com % e minutos
        def add_labels(bars, porcentagens, minutos):
            for bar, pct, min_ in zip(bars, porcentagens, minutos):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2, height + 1.2,
                        f"{pct:.1f}%\n({min_:.0f} min)",
                            ha='center', va='bottom', fontsize=8)

        add_labels(bars_uptime, uptimes, uptimes_min)
        add_labels(bars_downtime, downtimes, downtimes_min)

        ax.set_xticks(x)
        ax.set_xticklabels(racks, rotation=45)
        ax.set_ylabel('% de Tempo')
        ax.set_title('Tempo de Uptime e Downtime por Rack')
        ax.legend()

        ax.set_ylim(top=max(uptimes + downtimes) * 1.25)
        plt.tight_layout(pad=2.0)  # Adiciona um respiro
        plt.subplots_adjust(top=0.92, bottom=0.15)  # Ajuste fino se quiser 
        plt.show()  # Essa é a segunda (e última) janela que você quer abrir

        

#até aqui
# Interface Gráfica
root = tk.Tk()
root.title("Analisador de Vácuo por Rack (BCT)")
root.geometry("450x250")
root.resizable(False, False)

label = tk.Label(root, text="Clique no botão abaixo para selecionar um arquivo Excel", wraplength=400)
label.pack(pady=20)

botao = tk.Button(root, text="Selecionar Arquivo Excel", command=selecionar_arquivo, bg="#4CAF50", fg="white", padx=10, pady=5)
botao.pack()

root.mainloop()
