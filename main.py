############ Importação das bibliotecas utilizadas ############
import pandas as pd
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
from PIL import Image
import io
#import win32clipboard
import holidays

############ Acessando planilha pelo google planilhas ############

final_url = f"https://docs.google.com/spreadsheets/d/{st.secrets["csv_url"]}/gviz/tq?tqx=out:csv&gid={st.secrets["gid"]}"
db = pd.read_csv(final_url)

final_url_SGT = f"https://docs.google.com/spreadsheets/d/{st.secrets["csv_url"]}/gviz/tq?tqx=out:csv&gid={st.secrets["gid_SGT"]}"
db_SGT = pd.read_csv(final_url_SGT)

############ preprocessamento ############
nomes = {
    "0001 - EBER": "Eber",
    "0002 - ÉVERTON": "Éverton",
    "0003 - GABRIEL": "Gabriel",
    "0004 - HENRIQUE": "Henrique",
    "0005 - JOSANA": "Josana",
    "0006 - LUIZ": "Luiz",
    "0007 - MATEUS": "Mateus",
    "0008 - RAFAEL LOIOLA": "Rafael",
    "0009 - RODRIGO": "Rodrigo",
    "0010 - PEDRO DOURADO": "Pedro",
    "9997 - SAMUEL": "Samuel",
    "9998 - PEDRO": "Pedro Estagiário",
    "9999 - ISABELA": "Isabela"
}

colab_SGT = {
        "EVERTON FLEURY VICTORINO VALLE",
        "JOSANA BARCELAR BATISTA ANDRADE",
        "HENRIQUE NASCIMENTO BARROS",
        "MATEUS DANTAS SILVA",
        "EBER VASCONCELOS JUNIOR",
        "RAFAEL GOMES LOIOLA",
        "PEDRO VICTOR DOURADO SILVA",
        "GABRIEL DE OLIVEIRA RODRIGUES GOMES",
        "RODRIGO COSTA SILVA"
    }

listnomescompletos = {
    "Gabriel" : "GABRIEL DE OLIVEIRA RODRIGUES GOMES",
    "Eber" : "EBER VASCONCELOS JUNIOR",
    "Henrique" : "HENRIQUE NASCIMENTO BARROS",
    "Pedro" : "PEDRO VICTOR DOURADO SILVA",
    "Mateus" : "MATEUS DANTAS SILVA",
    "Josana" : "JOSANA BARCELAR BATISTA ANDRADE",
    "Rafael" : "RAFAEL GOMES LOIOLA",
    "Éverton" : "EVERTON FLEURY VICTORINO VALLE",
    "Isabela" : "ISABELA CRISÓSTOMO MELO",
    "Samuel" : "SAMUEL CARVALHO DE ALMEIDA",
    "Pedro Estagiário": "PEDRO VICTOR DOURADO SILVA"
}

# Renomeando planilhas
db = db.rename(columns={
    "Data ": "Data",
    "Nome ": "Nome",
    "NB+P? ": "NB+P?",
    "Carimbo de data/hora ": "Carimbo de data/hora",
    "B+P - Empresa Atendida ": "B+P - Empresa Atendida",
    "Empresa Atendida ": "Empresa Atendida",
    "Horas Dedicadas no atendimento ": "Horas Dedicadas no atendimento",
    "Tipo de Atividade ": "Tipo de Atividade",
    "Descrição da atividade ": "Descrição da atividade",
    "Selecione a etapa do projeto ": "Selecione a etapa do projeto"
})

def ajustar_mes(data):
    if data.day >= 26:
        # Se for dezembro, vai para janeiro do próximo ano
        if data.month == 12:
            return f"{data.year + 1}-01"
        else:
            return f"{data.year}-{data.month + 1:02d}"
    else:
        return data.strftime("%Y-%m")
    
def copy_matplotlib_fig_to_clipboard(fig):
    
    # 1. Salvar a figura em um buffer de bytes
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches='tight')
    buffer.seek(0) # Rebobina o buffer para o início
    
    # 2. Abrir os dados da imagem com o Pillow
    img = Image.open(buffer)
    
    # 3. O resto do código (converter para DIB e enviar ao clipboard)
    #    é exatamente igual ao da versão Plotly.
    output = io.BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    
    # win32clipboard.OpenClipboard()
    # win32clipboard.EmptyClipboard()
    # win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    # win32clipboard.CloseClipboard()
    
# Transformação de dados
db["Data"] = pd.to_datetime(db["Data"], errors="coerce", dayfirst=True)
db["Horas Dedicadas no atendimento"] = db["Horas Dedicadas no atendimento"].str.replace(",", ".").astype(float)
db["Mes"] = db["Data"].apply(ajustar_mes)
db["Ano"] = db["Data"].dt.strftime("%Y")  # Nova coluna para o ano
db = db.sort_values("Data")


############ Configs página Streamlit ############
st.set_page_config(
    page_title="Dash",
    page_icon="⚡",
    layout="wide"
)

# Menu principal na sidebar
pagina = "Colaboradores"

# Pegar meses e anos únicos já presentes no banco de dados
meses_existentes = sorted(db["Mes"].unique())
anos_existentes = sorted(db["Ano"].unique(), reverse=True)  # Ordena do mais recente

# Obter a data atual
hoje = pd.Timestamp.now()

# Se estivermos no dia 26 ou depois, adicionar o próximo mês à lista
if hoje.day >= 26:
    proximo_mes = (hoje + pd.DateOffset(months=1)).strftime("%Y-%m")
    if proximo_mes not in meses_existentes:
        meses_existentes.append(proximo_mes)

#Variável de verificação de grande intervalo selecionado
grandeIntervalo = False


# Selectbox para escolher o período (ano ou mês)
periodo_selecionado = st.sidebar.selectbox(
    "Período",
    ["Todos"] + anos_existentes + meses_existentes,
    index=0
)

# Lógica para determinar o período selecionado
if periodo_selecionado == "Todos":
    # Filtro para todos os dados
    db_filtered = db.copy()
    data_inicio = db["Data"].min()
    data_fim = db["Data"].max()
    periodo_titulo = "Todos os dados"
    grandeIntervalo = True
elif len(periodo_selecionado) == 4:  # É um ano (ex: "2025")
    # Filtro para o ano selecionado
    db_filtered = db[db["Ano"] == periodo_selecionado]
    data_inicio = pd.to_datetime(f"{periodo_selecionado}-01-01")
    data_fim = pd.to_datetime(f"{periodo_selecionado}-12-31")
    periodo_titulo = f"Ano {periodo_selecionado}"
    grandeIntervalo = True
else:  # É um mês (ex: "2025-01")
    # Filtro para o mês selecionado (mantém a lógica original)
    mes_ano = pd.to_datetime(periodo_selecionado + "-01")
    inicio_periodo = (mes_ano - pd.DateOffset(months=1)).replace(day=26)
    fim_periodo = mes_ano.replace(day=25)
    db_filtered = db[(db["Data"] >= inicio_periodo) & (db["Data"] <= fim_periodo)]
    data_inicio = inicio_periodo
    data_fim = fim_periodo
    periodo_titulo = f"Mês {periodo_selecionado}"
    grandeIntervalo = False


# Lógica para seleção de colaboradores baseada no período selecionado
if periodo_selecionado == "2025-01":
    colaboradores = [
        "0001 - EBER", "0002 - ÉVERTON", "0003 - GABRIEL", "0004 - HENRIQUE", 
        "0005 - JOSANA", "0006 - LUIZ", "0007 - MATEUS", "0008 - RAFAEL LOIOLA", 
        "9998 - PEDRO", "9999 - ISABELA"
    ]
elif periodo_selecionado == "2025-02":
    colaboradores = [
        "0001 - EBER", "0002 - ÉVERTON", "0003 - GABRIEL", "0004 - HENRIQUE", 
        "0005 - JOSANA", "0006 - LUIZ", "0007 - MATEUS", "0008 - RAFAEL LOIOLA", 
        "9999 - ISABELA"
    ]
elif periodo_selecionado == "2025-03":
    colaboradores = [
        "0001 - EBER", "0002 - ÉVERTON", "0003 - GABRIEL", "0004 - HENRIQUE", 
        "0005 - JOSANA", "0007 - MATEUS", "0008 - RAFAEL LOIOLA", "9999 - ISABELA"
    ]
elif periodo_selecionado == "2025-04":
    colaboradores = [
        "0001 - EBER", "0002 - ÉVERTON", "0003 - GABRIEL", "0004 - HENRIQUE", 
        "0005 - JOSANA", "0007 - MATEUS", "0008 - RAFAEL LOIOLA", "0009 - RODRIGO",
        "0010 - PEDRO DOURADO", "9997 - SAMUEL", "9999 - ISABELA"
    ]
else:
    colaboradores = [
        "0001 - EBER", "0002 - ÉVERTON", "0003 - GABRIEL", "0004 - HENRIQUE", 
        "0005 - JOSANA", "0007 - MATEUS", "0008 - RAFAEL LOIOLA", 
        "0010 - PEDRO DOURADO", "9997 - SAMUEL", "9999 - ISABELA"
    ]

# Função para atualizar nome selecionado
def atualizar_nome():
    st.session_state.selected_nome = st.session_state.nome_temp

# Selectbox para escolher colaborador
Nome = st.sidebar.selectbox(
    "Nome", 
    colaboradores, 
    index=colaboradores.index(st.session_state.get("selected_nome")) if st.session_state.get("selected_nome") in colaboradores else 0,
    key="nome_temp",
    on_change=atualizar_nome
)

# Atualizar o nome salvo no estado ao selecionar
st.session_state.selected_nome = Nome

# Gerar todas as datas do período selecionado
todas_datas = pd.date_range(start=data_inicio, end=data_fim, freq='D')

# Filtrar dados do colaborador selecionado
db_Nome = db_filtered[db_filtered["Nome"] == Nome]
# Criar gráficos
db_agrupado = db_filtered.groupby("Nome")["Horas Dedicadas no atendimento"].sum().reset_index()
db_agrup = db_agrupado.sort_values("Horas Dedicadas no atendimento", ascending=True)

# Gráfico de equipe com matplotlib
fig_eq, ax = plt.subplots(figsize=(8, 6))
ax.barh(db_agrup["Nome"], db_agrup["Horas Dedicadas no atendimento"], color="powderblue")
ax.set_title(f"Horas por analista no período: {periodo_titulo}", fontsize=14)  # Atualizado
ax.set_xlabel("Horas Dedicadas no Atendimento")
ax.set_ylabel("Analista")
ax.grid(True, axis='x', linestyle='--', alpha=0.6)
plt.tight_layout()

# Gráfico de equipe com plotly
# Calcular porcentagens
db_filtered2 = db_filtered.groupby(["Nome", "Tipo de Atividade"])["Horas Dedicadas no atendimento"].sum().reset_index()
apenasAnalistas = st.sidebar.toggle("Apenas Analistas")

if apenasAnalistas:
    db_filtered2 = db_filtered2[~db_filtered2["Nome"].isin(["9997 - SAMUEL", "9999 - ISABELA", "9998 - PEDRO"])]
    db_AdminVsAtend = db_filtered2.groupby("Tipo de Atividade")["Horas Dedicadas no atendimento"].sum().reset_index()
else:
    db_AdminVsAtend = db_filtered2.groupby("Tipo de Atividade")["Horas Dedicadas no atendimento"].sum().reset_index()
totais = db_filtered2.groupby("Nome")["Horas Dedicadas no atendimento"].sum().reset_index()
db_filtered2 = db_filtered2.merge(totais, on="Nome", suffixes=('', '_total'))
db_filtered2['Porcentagem'] = (db_filtered2['Horas Dedicadas no atendimento'] / db_filtered2['Horas Dedicadas no atendimento_total']) * 100

# Criar texto customizado com horas e porcentagem
db_filtered2['Texto'] = db_filtered2.apply(
    lambda x: f"{x['Horas Dedicadas no atendimento']:.1f}h<br>({x['Porcentagem']:.1f}%)", 
    axis=1
)

fig_equipe = px.bar(db_filtered2, 
                   y='Nome', 
                   x="Horas Dedicadas no atendimento", 
                   color="Tipo de Atividade",
                   title=f'Horas por analista no período: {periodo_titulo}', 
                   orientation='h',
                   text='Texto',  # Usando o texto customizado
                   category_orders={"Nome": colaboradores},  # Isso força a ordem e inclusão de todos os analistas
                   hover_data={
                       'Horas Dedicadas no atendimento': ':.1f',
                       'Porcentagem': ':.1f%',
                       'Texto': False
                   })

fig_equipe.update_layout(
    title={
        'text': f"Horas por analista no período: {periodo_titulo}",
        'y':0.95,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'
    },
    barmode='stack',
    xaxis_title='Horas',
    yaxis_title='Analista',
    height=600,
    hovermode='y unified',
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    ),
    margin=dict(l=100, r=50, t=100, b=50),
    uniformtext_minsize=10,
    uniformtext_mode='hide'
)

# Ajustar posição e estilo do texto
fig_equipe.update_traces(
    textposition='inside',
    insidetextanchor='middle',
    textfont=dict(
        size=12,
        color='white'
    ),
    hovertemplate="<b>%{y}</b><br>Tipo: %{fullData.name}<br>Horas: %{x:.1f}<br>% do total: %{customdata[1]:.1f}%<extra></extra>"
)

# Gráfico do colaborador
db_Nome_grouped = db_Nome.groupby("Data")["Horas Dedicadas no atendimento"].sum()
db_completo = db_Nome_grouped.reindex(todas_datas, fill_value=0)
feriados_br = holidays.Brazil(years=range(data_inicio.year, data_fim.year + 1))

# Cores para os dias
cores = []
count_dia = 0
for dia in db_completo.index:
    if dia in feriados_br:
        cores.append('gold')
    elif dia.weekday() >= 5:
        cores.append('lightcoral')
    else:
        cores.append('teal')
        count_dia += 1

total_horas_periodo = db_completo.sum()
x_labels = db_completo.index.strftime('%d/%m')  # Formato mais informativo para períodos longos

fig_colaborador, ax = plt.subplots(figsize=(10, 4))

ax.bar(x_labels, db_completo.values, color=cores)

ax.set_title(f"Colaborador: {nomes[Nome]} \nPeríodo: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} \nHoras totais: {total_horas_periodo}")

ax.set_xlabel("Data")

ax.set_ylabel("Horas Dedicadas no Atendimento")

ax.grid(True, axis='y', linestyle='--', alpha=0.6)

# Configurar cores dos rótulos

label_colors = []

for dia in db_completo.index:

    if dia in feriados_br:

        label_colors.append('blue')

    elif dia.weekday() >= 5:

        label_colors.append('red')

    else:

        label_colors.append('black')



ax.set_xticks(range(len(x_labels)))

ax.set_xticklabels(x_labels, rotation=90 if len(x_labels) > 20 else 45)  # Rotações diferentes conforme quantidade de dias

for tick_label, color in zip(ax.get_xticklabels(), label_colors):

    tick_label.set_color(color)

# Linhas verticais após cada domingo (só para visualização mensal)

if len(x_labels) <= 31:  # Apenas para períodos curtos

    for i, dia in enumerate(db_completo.index):

        if dia.weekday() == 6:

            ax.axvline(x=i + 0.5, color='gray', linestyle='--', alpha=0.3)

# Mostrar carga horária mensal (ajustado para períodos longos)
if Nome == "9999 - ISABELA":
    carga_horaria = f"Carga horária do período: {count_dia*6} horas."
elif Nome == "9997 - SAMUEL":
    carga_horaria = f"Carga horária do período: {count_dia*4} horas."
else:
    carga_horaria = f"Carga horária do período: {count_dia*8} horas."

# Seção de copiar gráficos para clipboard - só aparece na página Colaboradores
if pagina == "Colaboradores":
    st.sidebar.markdown("---")
    st.sidebar.title("Copiar Gráficos para Clipboard")
    
    if st.sidebar.button("Gráficos Equipe"):
        if grandeIntervalo:
            st.warning("Selecione um mês específico")
        else:
            i = 0

            # Código para copiar gráficos da equipe...
            while(i < len(colaboradores)):
                
                if colaboradores[i] == Nome:
                    fig, ax = plt.subplots(figsize=(10, 6))
                    pivot_data = db_filtered2.pivot(index='Nome', columns='Tipo de Atividade', values='Horas Dedicadas no atendimento')
                    pivot_data.plot.barh(stacked=True, ax=ax)
                    ax.set_title(f'Horas por analista no período: {periodo_titulo}', fontsize=14)
                    ax.set_xlabel('Horas Dedicadas no atendimento')
                    ax.set_ylabel('Nome')
                    ax.legend(title='Tipo de Atividade', bbox_to_anchor=(1.05, 1), loc='upper left')
                    plt.tight_layout()
                    plt.grid()

                i += 1
            
            copy_matplotlib_fig_to_clipboard(fig)
            st.toast('Gráfico geral no Clipboard! Cole-o onde desejar com CTRL V', icon='🎉')

    if st.sidebar.button("Gráfico Colaborador"):
        if grandeIntervalo:
            st.warning("Selecione um mês específico")
        else:
            copy_matplotlib_fig_to_clipboard(fig_colaborador)
            st.toast("Gráfico do colaborador foi gravado no Clipboard. Aperte CTRL V para colá-lo onde desejar.", icon='🎉')

    with st.sidebar.expander("Melhor qualidade Gráfico Equipe"):
        st.write("""Para obter uma qualidade melhor do Gráfico Equipe, basta clicar no ícone de câmera ao passar o mouse pelo gráfico""")


    st.title(f"Horas lançadas no Período: {periodo_titulo}")

    fig_equipe.update_layout(showlegend=True)

# Página Colaboradores
if pagina == "Colaboradores":
    m0, col1, m1 = st.columns([5, 10, 5])
    m3, col2, m2 = st.columns([5, 10, 5])
    
    with col1:
        st.plotly_chart(fig_equipe, use_container_width=True)
    
    with col2:
        if grandeIntervalo == False:
            st.pyplot(fig_colaborador)
    
    st.write(carga_horaria)
    
    # Mostrar dataframe de todos os colaboradores (removido o toggle apenascolab)
    st.write(db_filtered2[['Nome', 'Tipo de Atividade', 'Horas Dedicadas no atendimento', 'Horas Dedicadas no atendimento_total', 'Porcentagem']])

    #contagem1 = db_filtered["NB+P?"].value_counts().reindex(["Sim", "Não"], fill_value=0).reset_index()
    fig_AdmVsAtend = px.pie(
        db_AdminVsAtend, 
        values='Horas Dedicadas no atendimento',  # Valores numéricos (horas)
        names='Tipo de Atividade',               # Categorias (ADMINISTRATIVO/ATENDIMENTO)
        title=f"Comparação entre Horas Atendidas e de Administração ({periodo_titulo})",
        hole=0.4,
        color_discrete_map={  # Mapeamento manual das cores
        'ATENDIMENTO': '#1E3F66',     # Azul escuro
        'ADMINISTRATIVO': '#6E9ECF'   # Azul claro
        },
        labels={'Horas Dedicadas no atendimento': 'Horas', 'Tipo de Atividade': 'Tipo'}
    )

    # Melhorar formatação (opcional)
    fig_AdmVsAtend.update_traces(
        textposition='outside',
        texttemplate=(
            '<b>%{label}</b><br>'      # Nome da categoria (ATENDIMENTO/ADMINISTRATIVO)
            '%{value:.1f}h<br>'        # Valor em horas (1 decimal)
            '(%{percent:.1%})'         # Porcentagem (formato 0.0%)
        ),
        insidetextfont=dict(color='white', size=12),
        insidetextorientation='horizontal',
        marker=dict(line=dict(color='black', width=1))
    )

    fig_AdmVsAtend.update_layout(
        showlegend=True,
        legend_title_text='Tipo de Atividade'
    )

    fig_AdmVsAtend