import re
import pandas as pd

# --- Configurações ---
edl_path = "Timeline 1.edl"   # caminho do seu arquivo EDL
excel_path = "markers_intervalos.xlsx"  # saída em Excel
fps = 30  # frames por segundo do vídeo

# --- Ler conteúdo do arquivo EDL ---
with open(edl_path, "r", encoding="utf-8", errors="ignore") as f:
    edl_content = f.readlines()

# Regex para capturar timecodes e marcador
pattern = re.compile(r"(\d{2}:\d{2}:\d{2}:\d{2}).*?\n\s*\|C:[^\|]*\|M:([^|]+)\|D:(\d+)")

markers = []

for i in range(len(edl_content) - 1):
    line = edl_content[i]
    next_line = edl_content[i + 1] if i + 1 < len(edl_content) else ""
    match = pattern.search(line + "\n" + next_line)
    if match:
        timecode, marker_name, duration_frames = match.groups()

        # Converter timecode para segundos
        hh, mm, ss, ff = map(int, timecode.split(":"))
        total_seconds = hh * 3600 + mm * 60 + ss + ff / fps

        markers.append({
            "Marker": marker_name.strip(),
            "Timecode": timecode,
            "Time (s)": total_seconds
        })

# --- Calcular intervalos entre marcadores ---
for i in range(1, len(markers)):
    markers[i]["Intervalo (s)"] = markers[i]["Time (s)"] - markers[i - 1]["Time (s)"]
markers[0]["Intervalo (s)"] = None  # o primeiro não tem intervalo

# --- Criar DataFrame e salvar no Excel ---
df = pd.DataFrame(markers)
df.to_excel(excel_path, index=False)

print(f"Arquivo Excel gerado: {excel_path}")
