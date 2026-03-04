import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# ═══════════════════════════════════════════════════════════════
# ROSENI TORN STIILIS TEEMA
# ═══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title='Roseni Catering | Andmete Töötlus',
    page_icon='🍽️',
    layout='wide',
    initial_sidebar_state='expanded'
)

# Custom CSS - Roseni Torn stiil
st.markdown('''
<style>
    :root {
        --roseni-dark: #1a1a2e;
        --roseni-darker: #0f0f1a;
        --roseni-gold: #c9a961;
        --roseni-gold-light: #e5d4a1;
        --roseni-cream: #f5f0e6;
    }
    
    .stApp {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f0f1a 100%);
    }
    
    .main-header {
        background: linear-gradient(90deg, #c9a961, #e5d4a1, #c9a961);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 3rem;
        font-weight: 700;
        text-align: center;
        padding: 1rem 0;
        letter-spacing: 3px;
    }
    
    .sub-header {
        color: #e5d4a1;
        text-align: center;
        font-size: 1.2rem;
        margin-bottom: 2rem;
        font-style: italic;
    }
    
    .stat-card {
        background: linear-gradient(145deg, #1e1e3a, #252550);
        border: 1px solid #c9a961;
        border-radius: 15px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 8px 32px rgba(201, 169, 97, 0.15);
        margin-bottom: 1rem;
    }
    
    .stat-number {
        font-size: 2.5rem;
        font-weight: 700;
        color: #c9a961;
    }
    
    .stat-label {
        color: #e5d4a1;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #c9a961 0%, #a88c4a 100%);
        color: #1a1a2e;
        border: none;
        border-radius: 25px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        letter-spacing: 1px;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #e5d4a1 0%, #c9a961 100%);
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(201, 169, 97, 0.4);
    }
    
    [data-testid='stFileUploader'] {
        background: rgba(201, 169, 97, 0.1);
        border: 2px dashed #c9a961;
        border-radius: 15px;
        padding: 1rem;
    }
    
    section[data-testid='stSidebar'] {
        background: linear-gradient(180deg, #0f0f1a 0%, #1a1a2e 100%);
        border-right: 1px solid #c9a961;
    }
    
    .gold-divider {
        height: 2px;
        background: linear-gradient(90deg, transparent, #c9a961, transparent);
        margin: 2rem 0;
    }
</style>
'', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# TRANSFORMATSIOONI FUNKTSIOONID
# ═══════════════════════════════════════════════════════════════

def transform_data(df):
    '''Transformeerib Worksheet andmed Tulemus formaati'''
    
    # Filtreeri välja subheading read
    df_filtered = df[df['is_subheading'] != 1].copy()
    
    # Arvuta veerg C (amount × additional_amount)
    df_filtered['calc_c'] = df_filtered['amount'] * df_filtered['additional_amount']
    
    # Arvuta hindbuumile (hind enne allahindlust)
    def calc_hindbuumile(row):
        if 0 < row['line_discount_percent'] < 100:
            return row['price'] / (1 - row['line_discount_percent']/100)
        return row['price']
    
    df_filtered['hindbuumile'] = df_filtered.apply(calc_hindbuumile, axis=1)
    
    # Sorteeri maksjate kaupa
    df_filtered = df_filtered.sort_values('payer')
    
    # Loo väljund tühjade ridadega maksjate vahel
    result_rows = []
    current_payer = None
    columns = ['payer', 'product_code', 'calc_c', 'product_name',
               'amount', 'unit', 'additional_amount',
               'line_discount_percent', 'hindbuumile', 'price', 'line_sum']
    
    for _, row in df_filtered.iterrows():
        if current_payer and current_payer != row['payer']:
            result_rows.append({col: '' for col in columns})
        current_payer = row['payer']
        result_rows.append({col: row.get(col, '') for col in columns})
    
    return pd.DataFrame(result_rows)

def create_summary_stats(df):
    '''Loob kokkuvõtva statistika'''
    stats = {}
    df_clean = df[df['payer'] != '']
    stats['total_rows'] = len(df_clean)
    stats['unique_payers'] = df_clean['payer'].nunique()
    stats['unique_products'] = df_clean['product_code'].nunique()
    try:
        stats['total_sum'] = df_clean['line_sum'].astype(float).sum()
    except:
        stats['total_sum'] = 0
    return stats

# ═══════════════════════════════════════════════════════════════
# KASUTAJALIIDES
# ═══════════════════════════════════════════════════════════════

# Header
st.markdown('<h1 class="main-header">🍽️ ROSENI CATERING</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Andmete Transformatsioon & Analüüs</p>', unsafe_allow_html=True)
st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown('## 🍽️ ROSENI TORN')
    st.markdown('---')
    st.markdown('### ⚙️ Seaded')
    show_raw_data = st.checkbox('Näita algandmeid', value=False)
    show_charts = st.checkbox('Näita graafikuid', value=True)
    st.markdown('---')
    st.markdown('### 📋 Juhised')
    st.markdown('''
    1. Lae üles Excel fail
    2. Vali 'Worksheet' leht
    3. Vajuta 'Transformeeri'
    4. Lae tulemus alla
    ''')
    st.markdown('---')
    st.markdown('*Roseni Majad OÜ*')

# Faili üleslaadimine
st.markdown('### 📁 Lae üles Excel fail')
uploaded_file = st.file_uploader(
    'Vali fail (.xlsx)',
    type=['xlsx', 'xls'],
    help='Lae üles Excel fail, mis sisaldab Worksheet lehte'
)

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        st.success(f'✅ Fail laetud! Leitud {len(sheet_names)} lehte.')
        
        selected_sheet = st.selectbox('Vali leht:', sheet_names, index=0)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        if show_raw_data:
            st.markdown('### 📊 Algandmed')
            st.dataframe(df.head(20), use_container_width=True)
        
        st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            transform_button = st.button('🔄 TRANSFORMEERI ANDMED', use_container_width=True)
        
        if transform_button:
            with st.spinner('Töötlen andmeid...'):
                required_cols = ['is_subheading', 'payer', 'product_code', 'product_name',
                                 'amount', 'unit', 'additional_amount',
                                 'line_discount_percent', 'price', 'line_sum']
                
                missing_cols = [col for col in required_cols if col not in df.columns]
                
                if missing_cols:
                    st.error(f'❌ Puuduvad veerud: {missing_cols}')
                else:
                    result_df = transform_data(df)
                    stats = create_summary_stats(result_df)
                    
                    # Statistika
                    st.markdown('### 📈 Kokkuvõte')
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f'''
                        <div class="stat-card">
                            <div class="stat-number">{stats['total_rows']}</div>
                            <div class="stat-label">Tooterida</div>
                        </div>
                        ''', unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f'''
                        <div class="stat-card">
                            <div class="stat-number">{stats['unique_payers']}</div>
                            <div class="stat-label">Maksjat</div>
                        </div>
                        ''', unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f'''
                        <div class="stat-card">
                            <div class="stat-number">{stats['unique_products']}</div>
                            <div class="stat-label">Toodet</div>
                        </div>
                        ''', unsafe_allow_html=True)
                    
                    with col4:
    total_sum_str = f"EUR {stats['total_sum']:,.2f}"
    st.markdown(f'''
    <div class="stat-card">
        <div class="stat-number">{total_sum_str}</div>
        <div class="stat-label">Kogusumma</div>
    </div>
    ''', unsafe_allow_html=True)
                    
                    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
                    
                    # Graafikud
                    if show_charts:
                        st.markdown('### 📊 Analüüs')
                        
                        df_clean = result_df[result_df['payer'] != ''].copy()
                        df_clean['line_sum'] = pd.to_numeric(df_clean['line_sum'], errors='coerce')
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            payer_sum = df_clean.groupby('payer')['line_sum'].sum().reset_index()
                            payer_sum = payer_sum.sort_values('line_sum', ascending=True)
                            
                            fig1 = px.bar(payer_sum, x='line_sum', y='payer',
                                          orientation='h',
                                          title='💰 Summa maksjate kaupa',
                                          color='line_sum',
                                          color_continuous_scale=['#1a1a2e', '#c9a961'])
                            fig1.update_layout(
                                plot_bgcolor='rgba(0,0,0,0)',
                                paper_bgcolor='rgba(0,0,0,0)',
                                font_color='#e5d4a1',
                                showlegend=False
                            )
                            st.plotly_chart(fig1, use_container_width=True)
                        
                        with col2:
                            fig2 = px.pie(payer_sum, values='line_sum', names='payer',
                                          title='📈 Osakaal maksjate kaupa',
                                          color_discrete_sequence=['#c9a961', '#a88c4a', '#e5d4a1',
                                                                    '#8b7355', '#d4b896', '#6b5344'])
                            fig2.update_layout(
                                plot_bgcolor='rgba(0,0,0,0)',
                                paper_bgcolor='rgba(0,0,0,0)',
                                font_color='#e5d4a1'
                            )
                            st.plotly_chart(fig2, use_container_width=True)
                    
                    st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
                    
                    # Tabel
                    st.markdown('### 📋 Transformeeritud Andmed')
                    st.dataframe(result_df, use_container_width=True, height=400)
                    
                    # Download
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, sheet_name='Tulemus', index=False)
                    
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        st.download_button(
                            label='⬇️ LADI TULEMUS ALLA',
                            data=output.getvalue(),
                            file_name='roseni_tulemus.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            use_container_width=True
                        )
    
    except Exception as e:
        st.error(f'❌ Viga: {str(e)}')

else:
    st.info('👆 Lae üles Excel fail alustamiseks')
    
    st.markdown('### 🎯 Mida see rakendus teeb?')
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('''
        <div class="stat-card">
            <div class="stat-number">🔄</div>
            <div class="stat-label">Transformeerimine</div>
            <p style="color:#a0a0a0;font-size:0.85rem;">
            Filtreerib ja arvutab väärtused
            </p>
        </div>
        ''', unsafe_allow_html=True)
    
    with col2:
        st.markdown('''
        <div class="stat-card">
            <div class="stat-number">📊</div>
            <div class="stat-label">Analüüs</div>
            <p style="color:#a0a0a0;font-size:0.85rem;">
            Grupeerib ja loob graafikud
            </p>
        </div>
        ''', unsafe_allow_html=True)
    
    with col3:
        st.markdown('''
        <div class="stat-card">
            <div class="stat-number">⬇️</div>
            <div class="stat-label">Eksport</div>
            <p style="color:#a0a0a0;font-size:0.85rem;">
            Lae tulemus alla
            </p>
        </div>
        ''', unsafe_allow_html=True)

# Footer
st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)
st.markdown('''
<p style="text-align:center; color:#c9a961; font-size:0.9rem;">
    🍽️ Roseni Catering | Parimad emotsioonid ja hõrgud maitsed<br>
    <span style="color:#666;">© 2024 Roseni Torn</span>
</p>
'', unsafe_allow_html=True)
