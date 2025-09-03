import keyboard
import time
import pandas as pd # Importa a biblioteca pandas
from datetime import datetime # Para formatar a data e hora de forma legível

# Lista para armazenar os registros de tempo (timestamps)
marcas_de_tempo = []

# Variável para controlar o loop principal do programa
rodando = True

def processar_tecla(evento):
    """
    Esta função é chamada sempre que uma tecla é pressionada.
    Ela verifica se a tecla é 'ctrl' ou 'space' e age de acordo.
    """
    global rodando, marcas_de_tempo

    # Usamos evento.name para obter o nome da tecla pressionada
    # Verificamos por 'left ctrl' e 'right ctrl' para garantir que funcione
    if evento.name == 'ctrl' or evento.name == 'left ctrl' or evento.name == 'right ctrl':
        # Pega o tempo atual com alta precisão
        tempo_atual = time.time()
        marcas_de_tempo.append(tempo_atual)
        print(f"-> Marcação {len(marcas_de_tempo)} adicionada!")

    elif evento.name == 'space':
        print("\nTecla de espaço pressionada. Finalizando o programa...")
        rodando = False

def processar_e_exportar_resultados():
    """
    Calcula os intervalos, exibe um relatório final na tela
    e exporta todos os dados para um arquivo Excel.
    """
    print("\n--- Relatório Final ---")
    if not marcas_de_tempo:
        print("Nenhuma marcação foi registrada.")
        return

    print(f"\nTotal de marcações registradas: {len(marcas_de_tempo)}")
    
    # --- Preparação dos dados para o relatório e exportação ---
    dados_para_exportar = []
    intervalos = []

    for i, timestamp in enumerate(marcas_de_tempo):
        # Converte o timestamp para um formato de data e hora legível
        data_hora_legivel = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        
        intervalo_segundos = 0.0 # O primeiro item não tem intervalo anterior
        if i > 0:
            intervalo_segundos = timestamp - marcas_de_tempo[i-1]
            intervalos.append(intervalo_segundos)

        # Adiciona os dados formatados em um dicionário
        dados_para_exportar.append({
            "Marcação Nº": i + 1,
            "Data e Hora": data_hora_legivel,
            "Timestamp (Segundos)": timestamp,
            "Intervalo (Segundos)": intervalo_segundos if i > 0 else "N/A"
        })

    # --- Exibição do resumo na tela ---
    if len(marcas_de_tempo) > 1:
        print("\nIntervalos entre as marcações (em segundos):")
        for i, intervalo in enumerate(intervalos):
            print(f"  Intervalo {i+1} (entre marcação {i+1} e {i+2}): {intervalo:.4f} segundos")
        
        tempo_medio = sum(intervalos) / len(intervalos)
        print(f"\nTempo médio dos intervalos: {tempo_medio:.4f} segundos")

    # --- Exportação para o Excel usando Pandas ---
    try:
        # Cria um DataFrame do Pandas a partir da nossa lista de dicionários
        df = pd.DataFrame(dados_para_exportar)

        # Define o nome do arquivo de saída
        nome_arquivo = "relatorio_de_tempo.xlsx"
        
        # Exporta o DataFrame para um arquivo Excel, sem o índice padrão do pandas
        df.to_excel(nome_arquivo, index=False)
        
        print(f"\n✅ Dados exportados com sucesso para o arquivo '{nome_arquivo}'")

    except Exception as e:
        print(f"\n❌ Ocorreu um erro ao exportar para o Excel: {e}")

    print("\n--- Fim do Relatório ---")


# --- Início do Programa Principal ---
print("Programa de Marcação de Tempo iniciado.")
print("Pressione 'Ctrl' para adicionar uma marcação no tempo.")
print("Pressione 'Espaço' para finalizar e ver os resultados.")
print("-" * 40)

# Registra a função 'processar_tecla' para ser chamada a cada tecla pressionada
keyboard.on_press(processar_tecla)

# Loop principal que mantém o programa rodando
while rodando:
    time.sleep(0.1)

# Remove o "gancho" do teclado para limpar os recursos
keyboard.unhook_all()

# Após o loop terminar, exibe e exporta os resultados
processar_e_exportar_resultados()