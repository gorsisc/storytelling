# Bibliotecas Necessárias
# -------------------------------------------------------------------------

import pandas as pd
from nrclex import NRCLex
import matplotlib.pyplot as plt
import numpy as np
import os
import logging
import re # Para limpeza de texto com expressões regulares

# -------------------------------------------------------------------------
# Configuração do Logging (Controla as mensagens exibidas durante a execução)
# -------------------------------------------------------------------------
# Níveis: DEBUG (mais detalhado), INFO (padrão), WARNING, ERROR, CRITICAL
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s -
%(message)s')

# -------------------------------------------------------------------------
# CONFIGURAÇÃO DE ARQUIVOS E COLUNAS (ADAPTAR CONFORME
NECESSÁRIO)
# -------------------------------------------------------------------------
# Caminho para o arquivo de entrada Excel
# Exemplo: INPUT_FILE_PATH = "C:/Users/SeuUsuario/Documentos/meus_dados.xlsx"
# Exemplo: INPUT_FILE_PATH = "data/input/planilha_geral.xlsx"
INPUT_FILE_PATH = r'path/to/your/input_data.xlsx' # <<< MUDE AQUI

# Caminhos para os arquivos de saída (serão criados pelo script)
# Exemplo: OUTPUT_DIR = "data/output/"
OUTPUT_DIR = "path/to/your/output_folder/" # <<< MUDE AQUI (use / no final)
OUTPUT_INDIVIDUAL_PATH = os.path.join(OUTPUT_DIR,
'dados_analisados_completo.xlsx')
OUTPUT_SUMMARY_PATH = os.path.join(OUTPUT_DIR,
'resumo_comparativo_categorias.xlsx')
OUTPUT_CHART_PATH = os.path.join(OUTPUT_DIR,
'grafico_radar_comparativo_categorias.png')

# Nomes EXATOS das colunas na sua planilha Excel
# Exemplo: COL_TEXT = "TextoDoComentario"
# Exemplo: COL_CATEGORY = "Grupo"
COL_TEXT = 'comentários' # <<< MUDE AQUI (Coluna com os textos em INGLÊS)
COL_CATEGORY = 'Dia ' # <<< MUDE AQUI (Coluna com as categorias. Ex: 'Dia das
Mães')

# --- Definição das Emoções e Traduções ---
# Emoções primárias analisadas pelo NRCLex (usadas internamente)
EMOTIONS_TO_ANALYZE = ['fear', 'anger', 'anticipation', 'trust', 'surprise', 'sadness',
'disgust', 'joy']
# Todas as chaves retornadas por NRCLex (inclui sentimentos)
ALL_EMOTION_KEYS = EMOTIONS_TO_ANALYZE + ['positive', 'negative']

# Dicionário para traduzir nomes das emoções para Português (para gráficos e resumos)
TRADUCAO_EMOCOES = {
'fear': 'Medo', 'anger': 'Raiva', 'anticipation': 'Antecipação',
'trust': 'Confiança', 'surprise': 'Surpresa', 'sadness': 'Tristeza',
'disgust': 'Nojo', 'joy': 'Alegria', 'positive': 'Positivo', 'negative': 'Negativo'
}
# Tradução apenas das 8 emoções primárias (usado no gráfico e resumo principal)
TRADUCAO_EMOCOES_GRAFICO = {k: v for k, v in TRADUCAO_EMOCOES.items()
if k in EMOTIONS_TO_ANALYZE}
# Dicionário para traduzir colunas no resumo Excel (Raw e Perc)
TRADUCAO_COLUNAS_EXCEL = {
**{k: v for k, v in TRADUCAO_EMOCOES.items() if k in
EMOTIONS_TO_ANALYZE},
**{f"{k}_perc_cat": f"{v} (%)" for k, v in TRADUCAO_EMOCOES.items() if k in
EMOTIONS_TO_ANALYZE}
# Adicione positivo/negativo aqui se foram incluídos na análise principal e desejados no
Excel
}

# -------------------------------------------------------------------------
# Funções Auxiliares
# -------------------------------------------------------------------------

def clean_text(text):
"""
Limpa o texto removendo caracteres não alfanuméricos (exceto espaços),
convertendo para minúsculas e normalizando espaços.
Retorna string vazia se a entrada não for uma string.
"""
if not isinstance(text, str):
return ""
text = text.lower() # Converte para minúsculas
# Remove pontuações, etc., mantendo letras, números e espaços
text = re.sub(r'[^\w\s]', '', text)

# Remove espaços múltiplos e no início/fim
text = re.sub(r'\s+', ' ', text).strip()
return text

def analyze_emotions_detailed_perc(text):
"""
Analisa um único texto (após limpeza) usando NRCLex.

Calcula os scores brutos (contagem de palavras) e o percentual de cada emoção
relativo ao total de palavras de emoção encontradas *neste* texto.

Retorna:
pd.Series: Contendo scores brutos e percentuais para todas as emoções
definidas em ALL_EMOTION_KEYS (com sufixo _% para percentuais).
Retorna zeros se nenhuma emoção for encontrada ou em caso de erro.
As chaves/índice da série são os nomes em INGLÊS das emoções.
"""
cleaned_text = clean_text(text)
logging.debug(f"Texto Limpo: '{cleaned_text[:100]}...'") # Log para depuração

# Inicializa dicionários com zeros
raw_scores_dict = {key: 0 for key in ALL_EMOTION_KEYS}
perc_scores_dict = {f"{key}_%": 0.0 for key in ALL_EMOTION_KEYS}

if cleaned_text: # Só processa se houver texto após a limpeza
try:
# Cria o objeto NRCLex para o texto limpo
emotion_data = NRCLex(cleaned_text)
# Obtém a contagem bruta de palavras para cada emoção/sentimento
raw_scores_found = emotion_data.raw_emotion_scores
logging.debug(f"Scores Brutos NRCLex: {raw_scores_found}") # Log para depuração

# Atualiza o dicionário de scores brutos e calcula o total encontrado
total_raw_score = 0

for key, value in raw_scores_found.items():
# Normaliza chave 'anticip.' se aparecer em versões antigas
normalized_key = 'anticipation' if key == 'anticip.' else key
if normalized_key in raw_scores_dict: # Considera apenas as chaves esperadas
raw_scores_dict[normalized_key] = value
total_raw_score += value # Soma para calcular o percentual

# Calcula os percentuais se alguma palavra de emoção foi encontrada
if total_raw_score > 0:
logging.debug(f"Total Raw Score (para % intra-texto): {total_raw_score}")
for key in ALL_EMOTION_KEYS:
# Calcula percentual relativo ao total de emoções neste texto
perc_scores_dict[f"{key}_%"] = (raw_scores_dict[key] / total_raw_score) * 100
else:
logging.debug("Total Raw Score é 0, percentuais permanecem 0.")

except Exception as e:
# Loga um aviso se ocorrer erro no NRCLex para um texto específico
logging.warning(f"Erro durante análise NRCLex do texto: '{cleaned_text[:50]}...' -
Erro: {e}")
# Mantém scores zerados em caso de erro

# Combina os dicionários de scores brutos e percentuais
combined_scores = {**raw_scores_dict, **perc_scores_dict}
# Define a ordem das colunas/índice na série retornada
ordered_keys = ALL_EMOTION_KEYS + [f"{key}_%" for key in
ALL_EMOTION_KEYS]
# Retorna como Série pandas, garantindo todas as colunas e preenchendo com 0 se faltar
return pd.Series(combined_scores).reindex(ordered_keys).fillna(0)

def create_and_save_radar_chart_comparative(df_summary_perc, labels_pt, categories,
save_path):
"""

Cria e salva um gráfico de radar comparativo com múltiplas categorias (linhas coloridas).

Args:
df_summary_perc (pd.DataFrame): DataFrame com emoções (índice, já traduzido se
desejado
nos labels_pt) e categorias (colunas), contendo os valores
percentuais a serem plotados.
labels_pt (list): Lista dos nomes das emoções em português para os eixos do radar.
categories (list): Lista dos nomes das categorias (para a legenda e cores).
save_path (str): Caminho completo do arquivo para salvar a imagem do gráfico (ex:
.png).
"""
num_vars = len(labels_pt) # Número de eixos (emoções)
if num_vars == 0 or df_summary_perc.empty:
logging.warning("Dados insuficientes ou vazios para gerar o gráfico radar
comparativo.")
return

# Calcula ângulos para os eixos do radar
angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
angles += angles[:1] # Fecha o círculo

# Cria a figura e o eixo polar
fig, ax = plt.subplots(figsize=(11, 9), subplot_kw=dict(polar=True)) # Tamanho da figura

# Define cores distintas para as categorias
# 'tab10' tem 10 cores. Use 'viridis', 'plasma' ou outro se tiver mais categorias.
try:
colors = plt.cm.get_cmap('tab10', len(categories))
except ValueError: # Fallback se o número de categorias for inválido para o cmap
colors = plt.cm.get_cmap('viridis', len(categories))

max_val_overall = 0 # Para ajustar a escala do eixo Y

# Itera sobre cada categoria para plotar sua linha no gráfico
for i, category in enumerate(categories):
# Verifica se a categoria existe como coluna no DataFrame
if category in df_summary_perc.columns:
# Pega os valores percentuais para a categoria, preenche NaN com 0
values = df_summary_perc[category].fillna(0).values.flatten().tolist()
# Repete o primeiro valor no final para fechar a linha do radar
values += values[:1]
# Atualiza o valor máximo encontrado para ajustar a escala do eixo Y
max_val_overall = max(max_val_overall, max(values))

# Plota a linha da categoria no gráfico radar
ax.plot(angles, values, linewidth=1.5, linestyle='solid', color=colors(i),
label=category)
# Opcional: Preencher a área sob a linha (pode poluir com muitas categorias)
# ax.fill(angles, values, color=colors(i), alpha=0.1)
else:
logging.warning(f"Categoria '{category}' não encontrada nos dados para o gráfico.")

# Configurações visuais do gráfico
ax.set_xticks(angles[:-1]) # Posições dos ticks dos eixos das emoções
ax.set_xticklabels(labels_pt) # Nomes das emoções em Português nos eixos
ax.tick_params(axis='x', labelsize=10) # Tamanho da fonte dos labels das emoções

# Ajusta o eixo Y (valores percentuais)
max_val_plot = max_val_overall if max_val_overall > 0 else 1 # Evita divisão por zero ou
limite 0
yticks = np.linspace(0, max_val_plot, 5) # 5 níveis de ticks no eixo Y
ax.set_yticks(yticks)
ax.set_yticklabels([f"{i:.1f}%" for i in yticks]) # Formata labels do eixo Y como percentual
ax.set_ylim(0, max_val_plot * 1.05) # Deixa um pequeno espaço no topo

# Adiciona título e legenda

ax.set_title('Comparativo de Emoções por Categoria (% Relativo dentro da Categoria)',
size=16, y=1.12)
# Posiciona a legenda fora da área do gráfico para não cobrir os dados
ax.legend(loc='upper right', bbox_to_anchor=(1.35, 1.1), fontsize=10)
ax.grid(True) # Mostra a grade do radar

# Salva o gráfico
try:
# bbox_inches='tight' tenta ajustar a figura para que a legenda não seja cortada
plt.savefig(save_path, dpi=300, bbox_inches='tight')
logging.info(f"Gráfico radar comparativo salvo com sucesso em: {save_path}")
except Exception as e:
logging.error(f"Falha ao salvar o gráfico comparativo em {save_path}: {e}")
finally:
# Fecha a figura para liberar memória
plt.close(fig)

# ---------------------------------------------------------------------------
# Bloco Principal de Execução
# ---------------------------------------------------------------------------
if __name__ == "__main__":
logging.info(f"--- Iniciando Análise Comparativa de Emoções ---")
logging.info(f"Arquivo de Entrada: {INPUT_FILE_PATH}")
logging.info(f"Diretório de Saída: {OUTPUT_DIR}")

# --- 0. Verifica Diretório de Saída ---
try:
if not os.path.exists(OUTPUT_DIR):
os.makedirs(OUTPUT_DIR)
logging.info(f"Diretório de saída criado: {OUTPUT_DIR}")
except Exception as e:
logging.error(f"Não foi possível criar o diretório de saída '{OUTPUT_DIR}'. Verifique
permissões. Erro: {e}")

exit() # Termina se não puder criar o diretório

# --- (Opcional) Teste Rápido do NRCLex ---
# Ajuda a verificar se a biblioteca e seus dados estão funcionando
try:
test_text = "This is a happy and joyful test with a little fear."
test_emotion = NRCLex(clean_text(test_text))
test_scores = test_emotion.raw_emotion_scores
logging.info(f"--- Teste NRCLex Rápido ---")
if sum(test_scores.values()) > 0:
logging.info("Teste NRCLex: OK (Biblioteca parece funcional).")
else:
# Se o teste falhar, pode indicar problema com download de corpora
logging.warning("ALERTA TESTE NRCLex: Scores zerados no texto de teste.")
logging.warning("Verifique se executou 'python -m textblob.download_corpora'.")
logging.info(f"--- Fim Teste NRCLex ---")
except Exception as e:
logging.error(f"Erro durante o teste do NRCLex: {e}")
logging.error("Isso pode indicar problema na instalação do NRCLex ou
TextBlob/NLTK.")
logging.error("Certifique-se de ter instalado as bibliotecas e baixado os corpora ('python
-m textblob.download_corpora').")

# --- 1. Ler Planilha de Entrada ---
logging.info(f"Etapa 1: Lendo arquivo Excel
'{os.path.basename(INPUT_FILE_PATH)}'...")
try:
# Verifica se o arquivo existe
if not os.path.exists(INPUT_FILE_PATH):
logging.error(f"Erro Crítico: Arquivo de entrada não encontrado em
'{INPUT_FILE_PATH}'")
exit() # Termina se o arquivo não existe

# Lê a planilha usando pandas

df = pd.read_excel(INPUT_FILE_PATH, header=0) # Assume cabeçalho na linha 1

# Verifica se as colunas necessárias existem
required_cols = [COL_TEXT, COL_CATEGORY]
missing_cols = [col for col in required_cols if col not in df.columns]
if missing_cols:
# Erro fatal se colunas essenciais não forem encontradas
logging.error(f"Erro Crítico: Coluna(s) necessária(s) não encontrada(s) na planilha:
{missing_cols}")
logging.error(f"Colunas disponíveis na planilha: {df.columns.tolist()}")
logging.error(f"Verifique as variáveis 'COL_TEXT' ('{COL_TEXT}') e
'COL_CATEGORY' ('{COL_CATEGORY}') no script.")
exit() # Termina

logging.info(f"Planilha lida com sucesso. {len(df)} linhas encontradas.")
logging.info(f"Usando coluna '{COL_TEXT}' para análise de texto (deve estar em
INGLÊS).")
logging.info(f"Usando coluna '{COL_CATEGORY}' para agrupar categorias.")

except Exception as e:
logging.error(f"Erro inesperado ao ler a planilha Excel: {e}")
exit() # Termina em caso de erro de leitura

# --- 2. Analisar Emoções Linha a Linha ---
logging.info(f"Etapa 2: Analisando emoções para cada texto na coluna '{COL_TEXT}'...")
try:
# Aplica a função de análise a cada célula da coluna de texto
# O resultado é um DataFrame com as novas colunas de score raw e percentual
emotion_results_df = df[COL_TEXT].apply(analyze_emotions_detailed_perc)

# Junta o DataFrame original com os resultados da análise
df_analisado = pd.concat([df, emotion_results_df], axis=1)
logging.info("Análise de emoções linha a linha concluída.")
# Mostra as primeiras linhas no log (se em modo DEBUG)

logging.debug("Primeiras 5 linhas do DataFrame analisado (com scores):\n" +
df_analisado.head().to_string())

except Exception as e:
logging.error(f"Erro durante a aplicação da análise de emoções ou concatenação: {e}")
exit() # Termina se a análise principal falhar

# --- 3. Exportar Dados Individuais Completos ---
logging.info(f"Etapa 3: Exportando resultados individuais detalhados...")
try:
df_analisado.to_excel(OUTPUT_INDIVIDUAL_PATH, index=False,
engine='openpyxl')
logging.info(f"Resultados individuais salvos em: {OUTPUT_INDIVIDUAL_PATH}")
except Exception as e:
# Apenas avisa se a exportação individual falhar, mas continua
logging.error(f"Falha ao exportar dados individuais para
'{OUTPUT_INDIVIDUAL_PATH}': {e}")

# --- 4. Análise Comparativa Agregada por Categoria ---
logging.info(f"Etapa 4: Calculando análise agregada comparativa por
'{COL_CATEGORY}'...")
df_summary_comparativo = pd.DataFrame() # DataFrame para o Excel (Categorias x
Emoções PT Raw/%)
df_summary_perc_pivot = pd.DataFrame() # DataFrame para o Gráfico (Emoções PT x
Categorias %)

try:
# Calcula somas brutas por categoria para as 8 emoções primárias
grouped_raw_sums =
df_analisado.groupby(COL_CATEGORY)[EMOTIONS_TO_ANALYZE].sum()
logging.info(f"Somas brutas por categoria ({COL_CATEGORY}) calculadas.")
logging.debug("Somas brutas por categoria:\n" + grouped_raw_sums.to_string())

# Calcula percentuais relativos DENTRO de cada categoria
total_per_category = grouped_raw_sums.sum(axis=1) # Soma das 8 emoções por
categoria
# Evita divisão por zero se uma categoria não tiver nenhuma emoção primária
grouped_perc = grouped_raw_sums.apply(pd.to_numeric, errors='coerce') \
.div(total_per_category.replace(0, np.nan), axis=0).fillna(0) * 100
logging.info(f"Percentuais relativos dentro de cada categoria ({COL_CATEGORY})
calculados.")
logging.debug("Percentuais relativos por categoria (%):\n" + grouped_perc.to_string())

# --- Preparar Resumo para Excel (df_summary_comparativo) ---
# Formato: Índice=Categorias, Colunas=Emoções_PT (Raw + Perc)
grouped_perc_export = grouped_perc.copy()
grouped_perc_export.columns = [f"{col}_perc_cat" for col in
grouped_perc_export.columns] # Adiciona sufixo
df_temp_summary = pd.concat([grouped_raw_sums, grouped_perc_export], axis=1) #
Junta Raw e Perc
df_temp_summary_t = df_temp_summary.transpose() # Transpõe: Índice=Emoções
ENG Raw/Perc
# Traduz o índice (Emoções) para Português
df_temp_summary_t.index =
df_temp_summary_t.index.map(TRADUCAO_COLUNAS_EXCEL).fillna(df_temp_summar
y_t.index)
df_summary_comparativo = df_temp_summary_t.transpose() # Transpõe de volta
logging.info("DataFrame de resumo comparativo (para Excel) preparado.")
logging.debug("Resumo comparativo (Excel - PT):\n" +
df_summary_comparativo.to_string())

# --- Preparar Dados para Gráfico (df_summary_perc_pivot) ---
# Formato: Índice=Emoções_PT, Colunas=Categorias, Valores=%
df_perc_pivot_temp = grouped_perc.transpose() # Transpõe: Índice=Emoções ENG
# Traduz o índice (Emoções) para Português

df_perc_pivot_temp.index =
df_perc_pivot_temp.index.map(TRADUCAO_EMOCOES_GRAFICO).fillna(df_perc_pivot_
temp.index)
df_summary_perc_pivot = df_perc_pivot_temp # Atribui ao DataFrame final do gráfico
logging.info("DataFrame de percentuais (para Gráfico) preparado.")
logging.debug("Dados para Gráfico (df_summary_perc_pivot):\n" +
df_summary_perc_pivot.to_string())

except Exception as e:
logging.error(f"Erro durante o cálculo da análise agregada por categoria: {e}")
# DataFrames de resumo permanecerão vazios

# --- 5. Exportar Resumo Comparativo ---
logging.info(f"Etapa 5: Exportando resumo comparativo...")
if not df_summary_comparativo.empty:
try:
# Salva com categorias como índice e colunas traduzidas
df_summary_comparativo.to_excel(OUTPUT_SUMMARY_PATH, index=True,
index_label=COL_CATEGORY, engine='openpyxl')
logging.info(f"Resumo comparativo salvo em: {OUTPUT_SUMMARY_PATH}")
except Exception as e:
logging.error(f"Falha ao exportar resumo comparativo para
'{OUTPUT_SUMMARY_PATH}': {e}")
else:
# Aviso se o resumo não puder ser gerado devido a erro anterior
logging.warning("DataFrame de resumo comparativo está vazio (provavelmente devido
a erro na Etapa 4), exportação pulada.")

# --- 6. Criar e Salvar Gráfico Radar Comparativo ---
logging.info(f"Etapa 6: Gerando gráfico radar comparativo...")
if not df_summary_perc_pivot.empty:
# Obtém labels e categorias do DataFrame preparado para o gráfico
radar_labels_pt = df_summary_perc_pivot.index.tolist() # Emoções em PT
categories_list = df_summary_perc_pivot.columns.tolist() # Categorias

# Chama a função para criar e salvar o gráfico
create_and_save_radar_chart_comparative(
df_summary_perc_pivot, # Dados para plotar (Índice=Emoção PT,
Colunas=Categorias)
radar_labels_pt, # Labels dos eixos (Emoções PT)
categories_list, # Lista de categorias para legenda e cores
OUTPUT_CHART_PATH # Caminho para salvar
)
else:
# Aviso se o gráfico não puder ser gerado
logging.warning("Dados de percentual para o gráfico estão vazios (provavelmente devido a erro na Etapa 4), geração do gráfico pulada.")
logging.info(f"--- Análise Concluída ---")

