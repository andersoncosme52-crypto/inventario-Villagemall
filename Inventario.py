import streamlit as st, pandas as pd, sqlite3, plotly.express as px

st.set_page_config(page_title="Dashboard TI", layout="wide")
import streamlit as st, pandas as pd, sqlite3, plotly.express as px

st.set_page_config(page_title="Dashboard TI", layout="wide")

# =========================================================================
# 🎬 TITULO ANIMADO - ADICIONADO EXATAMENTE AQUI:
# =========================================================================
st.markdown(
    """
    <div style="background-color: #1F4E78; padding: 10px; border-radius: 5px; margin-bottom: 20px;">
        <marquee behavior="scroll" direction="left" scrollamount="7" style="color: #FFFFFF; font-size: 28px; font-weight: bold; font-family: Arial;">
            🚀 Villagemall Inventario
        </marquee>
    </div>
    """,
    unsafe_allow_html=True
)
# =========================================================================

# Conectando ao banco de dados de forma estável
conn = sqlite3.connect('dashboard_dados.db', check_same_thread=False)
c = conn.cursor()

# Conectando ao banco de dados de forma estável
conn = sqlite3.connect('dashboard_dados.db', check_same_thread=False)
c = conn.cursor()

# Criação das tabelas no banco de dados
c.execute('CREATE TABLE IF NOT EXISTS servidores (nome TEXT, rack TEXT, ip TEXT, localidade TEXT)')
c.execute('CREATE TABLE IF NOT EXISTS rede_ip (nome_computador TEXT, local TEXT, status TEXT, ip TEXT)')
c.execute('CREATE TABLE IF NOT EXISTS pinpad (local TEXT, modelo TEXT, num_serie TEXT)')
c.execute('CREATE TABLE IF NOT EXISTS ferramentas (nome TEXT, quantidade INTEGER, num_serie TEXT)')
c.execute('CREATE TABLE IF NOT EXISTS bobinas (tamanho TEXT PRIMARY KEY, quantidade INTEGER DEFAULT 0)')
c.execute('CREATE TABLE IF NOT EXISTS controle_ht (local TEXT PRIMARY KEY, quantidade INTEGER DEFAULT 0, status TEXT)')
c.execute(
    'CREATE TABLE IF NOT EXISTS chamados (id INTEGER PRIMARY KEY AUTOINCREMENT, num_chamado TEXT, empresa TEXT, status TEXT, descricao TEXT)')

bobinas = ['58MMX300M ENTRADA', '76MMX300 PAGAMENTO', '79MM80 GUICHE', '57MMX08 PAPA FILA']
terminais = ["EL1", "EVIP R1", "EVIP R2", "ESS1 R1", "ESS1 R2", "ESS1 R3", "ESS2 R1", "ESS2 R2", "ESS2 R3",
             "PAGAMENTO PORTARIA C", "PAGAMENTO PORTARIA B", "VIP PANORAMICO", "SS1 AMERICAS", "SS1 PANORAMICO",
             "SS1 LAGOA1", "SS1 LAGOA2", "SS1 LAGOA3", "GUICHE SS1 LAGOA", "SS2 PANORAMICO", "TPC SS2 AMERICAS",
             "TPC SS2 PANORAMICO", "TPC SS2 LAGOA"]
hts = ["ACEESO SSS1 AMERICAS", "SVIP", "E3", "E4", "E5", "E6", "E8", "E9", "E10", "E1"]

# Popula tabelas base se estiverem vazias
for b in bobinas: c.execute("INSERT OR IGNORE INTO bobinas VALUES (?,0)", (b,))
for h in hts: c.execute("INSERT OR IGNORE INTO controle_ht VALUES (?,0,'Disponível')", (h,))
conn.commit()

menu = st.sidebar.radio("Menu:",
                        ["📊 Dashboard", "🖥️ Servidores", "🌐 Rede/IP", "💳 Pinpad", "🔧 Ferramentas", "🧻 Bobinas", "📻 HT",
                         "🎫 Chamados"])

# ----------------- 1. DASHBOARD -----------------
if menu == "📊 Dashboard":
    st.title("📊 Painel de Indicadores Gerais")

    total_servidores = len(pd.read_sql("SELECT * FROM servidores", conn))
    total_pinpads = len(pd.read_sql("SELECT * FROM pinpad", conn))

    ferr_df = pd.read_sql("SELECT SUM(quantidade) as t FROM ferramentas", conn)
    total_ferramentas = int(ferr_df['t'].fillna(0).iloc[0])

    bob_df = pd.read_sql("SELECT SUM(quantidade) as t FROM bobinas", conn)
    total_bobinas = int(bob_df['t'].fillna(0).iloc[0])

    ht_df = pd.read_sql("SELECT SUM(quantidade) as t FROM controle_ht", conn)
    total_hts = int(ht_df['t'].fillna(0).iloc[0])

    total_chamados = len(pd.read_sql("SELECT * FROM chamados", conn))
    total_redes = len(pd.read_sql("SELECT * FROM rede_ip", conn))

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Servidores", total_servidores)
    m2.metric("Pinpads", total_pinpads)
    m3.metric("Ferramentas (Qtd Total)", total_ferramentas)
    m4.metric("Chamados Registrados", total_chamados)
    st.write("---")


    def gerar_excel_customizado():
        import io
        from openpyxl import Workbook
        from openpyxl.chart import (BarChart,
                                    PieChart, Reference)
        from openpyxl.styles import PatternFill, Font, Border, Side

        wb = Workbook()

        # Estilos Visuais Gerais
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
            top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
        )

        # ---------------------------------------------------------
        # ORELHA 1: 🧻 ESTOQUE DE BOBINAS (Com Gráfico de Barras)
        # ---------------------------------------------------------
        ws1 = wb.active
        ws1.title = "Estoque de Bobinas"
        ws1.sheet_properties.tabColor = "1F4E78"  # Azul Escuro
        ws1.sheet_view.showGridLines = True

        df_b_excel = pd.read_sql("SELECT tamanho, quantidade FROM bobinas", conn)
        ws1.append(["Modelo de Bobina", "Quantidade"])
        for idx, row in df_b_excel.iterrows():
            ws1.append([row['tamanho'], row['quantidade']])

        for row in ws1.iter_rows(min_row=1, max_row=len(df_b_excel) + 1, min_col=1, max_col=2):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill;
                    cell.font = header_font

        chart1 = BarChart()
        chart1.type = "col"
        chart1.title = "Gráfico de Barras: Estoque de Bobinas"
        data1 = Reference(ws1, min_col=2, min_row=1, max_row=len(df_b_excel) + 1)
        cats1 = Reference(ws1, min_col=1, min_row=2, max_row=len(df_b_excel) + 1)
        chart1.add_data(data1, titles_from_data=True)
        chart1.set_categories(cats1)
        ws1.add_chart(chart1, "D2")

        # ---------------------------------------------------------

        # ---------------------------------------------------------
        # ORELHA 3: 🖥️ SERVIDORES
        # ---------------------------------------------------------
        ws3 = wb.create_sheet(title="Servidores")
        ws3.sheet_properties.tabColor = "2ECC71"  # Verde
        ws3.sheet_view.showGridLines = True

        df_serv_excel = pd.read_sql("SELECT nome, rack, ip, localidade FROM servidores", conn)
        ws3.append(["Nome do Servidor", "Rack", "Endereço IP", "Localidade"])
        for idx, row in df_serv_excel.iterrows():
            ws3.append([row['nome'], row['rack'], row['ip'], row['localidade']])

        for row in ws3.iter_rows(min_row=1, max_row=len(df_serv_excel) + 1, min_col=1, max_col=4):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill;
                    cell.font = header_font

        # ---------------------------------------------------------
        # ORELHA 4: 💳 PINPADS
        # ---------------------------------------------------------
        ws4 = wb.create_sheet(title="Pinpads")
        ws4.sheet_properties.tabColor = "9B59B6"  # Roxo
        ws4.sheet_view.showGridLines = True

        df_pin_excel = pd.read_sql("SELECT local, modelo, num_serie FROM pinpad", conn)
        ws4.append(["Local de Instalação", "Modelo", "Número de Série"])
        for idx, row in df_pin_excel.iterrows():
            ws4.append([row['local'], row['modelo'], row['num_serie']])

        for row in ws4.iter_rows(min_row=1, max_row=len(df_pin_excel) + 1, min_col=1, max_col=3):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill;
                    cell.font = header_font

        # ---------------------------------------------------------
        # ORELHA 5: 🔧 FERRAMENTAS
        # ---------------------------------------------------------
        ws5 = wb.create_sheet(title="Ferramentas")
        ws5.sheet_properties.tabColor = "F39C12"  # Laranja
        ws5.sheet_view.showGridLines = True

        df_ferr_excel = pd.read_sql("SELECT nome, quantidade, num_serie FROM ferramentas", conn)
        ws5.append(["Nome do Item", "Quantidade", "Número de Série"])
        for idx, row in df_ferr_excel.iterrows():
            ws5.append([row['nome'], row['quantidade'], row['num_serie']])

        for row in ws5.iter_rows(min_row=1, max_row=len(df_ferr_excel) + 1, min_col=1, max_col=3):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill;
                    cell.font = header_font

        # ---------------------------------------------------------
        # ORELHA 6: 📻 RÁDIOS HT
        # ---------------------------------------------------------
        ws6 = wb.create_sheet(title="Rádios HT")
        ws6.sheet_properties.tabColor = "34495E"  # Cinza
        ws6.sheet_view.showGridLines = True

        df_ht_excel = pd.read_sql("SELECT local, quantidade, status FROM controle_ht", conn)
        ws6.append(["Posto / Local", "Quantidade", "Status Operacional"])
        for idx, row in df_ht_excel.iterrows():
            ws6.append([row['local'], row['quantidade'], row['status']])

        for row in ws6.iter_rows(min_row=1, max_row=len(df_ht_excel) + 1, min_col=1, max_col=3):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill;
                    cell.font = header_font

        # ---------------------------------------------------------
        # ORELHA 7: 🎫 CHAMADOS
        # ---------------------------------------------------------
        ws7 = wb.create_sheet(title="Chamados")
        ws7.sheet_properties.tabColor = "16A085"  # Verde Água
        ws7.sheet_view.showGridLines = True

        df_ch_excel = pd.read_sql("SELECT num_chamado, empresa, status, descricao FROM chamados", conn)
        ws7.append(["Número do Chamado", "Empresa Atendida", "Situação", "Descrição Detalhada"])
        for idx, row in df_ch_excel.iterrows():
            ws7.append([row['num_chamado'], row['empresa'], row['status'], row['descricao']])

        for row in ws7.iter_rows(min_row=1, max_row=len(df_ch_excel) + 1, min_col=1, max_col=4):
            for cell in row:
                cell.border = thin_border
                if cell.row == 1:
                    cell.fill = header_fill;
                    cell.font = header_font

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()


    st.download_button(
        label="📥 Gerar Relatório Excel com Gráficos",
        data=gerar_excel_customizado(),
        file_name="Inventario_TI_Formatado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_excel_dashboard"
    )
    st.write("---")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 1. Gráfico de Barras: Detalhado de Bobinas")
        df_b = pd.read_sql("SELECT * FROM bobinas", conn)
        if df_b['quantidade'].sum() > 0:
            st.plotly_chart(px.bar(df_b, x='tamanho', y='quantidade', color='tamanho',
                                   labels={'quantidade': 'Qtd', 'tamanho': 'Modelo'}), use_container_width=True)
        else:
            st.info("Sem estoque de bobinas gravado.")

    with col2:
        st.subheader("🍕 2. Gráfico de Pizza: Distribuição Total do Inventário")
        dados_gerais = {
            "Item do Inventário": ["Servidores", "IPs de Rede", "Pinpads", "Ferramentas", "Bobinas em Estoque",
                                   "Rádios HT", "Chamados"],
            "Quantidade Gravada": [total_servidores, total_redes, total_pinpads, total_ferramentas, total_bobinas,
                                   total_hts, total_chamados]
        }
        df_inventario_total = pd.DataFrame(dados_gerais)

        if df_inventario_total["Quantidade Gravada"].sum() > 0:
            fig_pizza_total = px.pie(df_inventario_total, values="Quantidade Gravada", names="Item do Inventário",
                                     hole=0.15, color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_pizza_total, use_container_width=True)
        else:
            st.info("Nenhum dado gravado no sistema ainda para gerar o gráfico.")

# ----------------- 2. SERVIDORES -----------------
elif menu == "🖥️ Servidores":
    st.title("🖥️ Servidores")
    with st.form("f_serv", clear_on_submit=True):
        n = st.text_input("Nome")
        r = st.text_input("Rack")
        i = st.text_input("IP")
        l = st.text_input("Localidade")
        if st.form_submit_button("Gravar") and n and i:
            c.execute("INSERT INTO servidores (nome, rack, ip, localidade) VALUES (?,?,?,?)", (n, r, i, l))
            conn.commit()
            st.success("Salvo!")
            st.rerun()

    df_serv = pd.read_sql("SELECT * FROM servidores", conn)
    st.dataframe(df_serv, use_container_width=True)
    df_serv = pd.read_sql("SELECT * FROM servidores", conn)
    st.dataframe(df_serv, use_container_width=True)

    # O código entra exatamente aqui embaixo da tabela:
    if not df_serv.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        serv_para_excluir = st.selectbox("Selecione o Servidor para remover:", df_serv['nome'].unique())
        if st.button("Excluir Servidor", type="primary"):
            c.execute("DELETE FROM servidores WHERE nome=?", (serv_para_excluir,))
            conn.commit()
            st.success("Servidor removido!")
            st.rerun()

    if not df_serv.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        serv_para_excluir = st.selectbox("Selecione o Servidor para remover:", df_serv['nome'].unique())
        if st.button("Excluir Servidor", type="primary"):
            c.execute("DELETE FROM servidores WHERE nome=?", (serv_para_excluir,))
            conn.commit()
            st.success("Servidor removido!")
            st.rerun()

# ----------------- 3. REDE / IP -----------------
elif menu == "🌐 Rede/IP":
    st.title("🌐 Rede / IP")
    with st.form("f_rede", clear_on_submit=True):
        n = st.text_input("Computador")
        l = st.text_input("Local")
        s = st.selectbox("Status", ["ATIVO", "VAGO"])
        i = st.text_input("IP")
        if st.form_submit_button("Gravar"):
            if n or i:
                c.execute("INSERT INTO rede_ip (nome_computador, local, status, ip) VALUES (?,?,?,?)", (n, l, s, i))
                conn.commit()
                st.success("Registro de rede salvo!")
                st.rerun()
            else:
                st.error("Por favor, preencha pelo menos o Computador ou o IP.")

    df_rede = pd.read_sql("SELECT * FROM rede_ip", conn)
    st.dataframe(df_rede, use_container_width=True)
    df_rede = pd.read_sql("SELECT * FROM rede_ip", conn)
    st.dataframe(df_rede, use_container_width=True)

    # O código entra exatamente aqui:
    if not df_rede.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        rede_para_excluir = st.selectbox("Selecione o Computador para remover:", df_rede['nome_computador'].unique(), key="del_rede")
        if st.button("Excluir Registro de Rede", type="primary", key="btn_del_rede"):
            c.execute("DELETE FROM rede_ip WHERE nome_computador=?", (rede_para_excluir,))
            conn.commit()
            st.success("Registro removido!")
            st.rerun()


    if not df_rede.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        rede_para_excluir = st.selectbox("Selecione o Computador para remover:", df_rede['nome_computador'].unique())
        if st.button("Excluir Registro de Rede", type="primary"):
            c.execute("DELETE FROM rede_ip WHERE nome_computador=?", (rede_para_excluir,))
            conn.commit()
            st.success("Registro removido!")
            st.rerun()

# ----------------- 4. PINPAD -----------------
elif menu == "💳 Pinpad":
    st.title("💳 Pinpad")
    with st.form("f_pin", clear_on_submit=True):
        l = st.text_input("Local")
        m = st.text_input("Modelo")
        s = st.text_input("N° Série")
        if st.form_submit_button("Gravar") and s:
            c.execute("INSERT INTO pinpad (local, modelo, num_serie) VALUES (?,?,?)", (l, m, s))
            conn.commit()
            st.success("Salvo com sucesso!")
            st.rerun()

    df_pin = pd.read_sql("SELECT * FROM pinpad", conn)
    st.dataframe(df_pin, use_container_width=True)
    df_pin = pd.read_sql("SELECT * FROM pinpad", conn)
    st.dataframe(df_pin, use_container_width=True)

    # O código entra exatamente aqui embaixo da tabela do Pinpad:
    if not df_pin.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        pin_para_excluir = st.selectbox("Selecione o N° Série para remover:", df_pin['num_serie'].unique(), key="del_pin")
        if st.button("Excluir Pinpad", type="primary", key="btn_del_pin"):
            c.execute("DELETE FROM pinpad WHERE num_serie=?", (pin_para_excluir,))
            conn.commit()
            st.success("Pinpad removido!")
            st.rerun()


    if not df_pin.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        pin_para_excluir = st.selectbox("Selecione o N° Série para remover:", df_pin['num_serie'].unique())
        if st.button("Excluir Pinpad", type="primary"):
            c.execute("DELETE FROM pinpad WHERE num_serie=?", (pin_para_excluir,))
            conn.commit()
            st.success("Pinpad removido!")
            st.rerun()

# ----------------- 5. FERRAMENTAS -----------------
elif menu == "🔧 Ferramentas":
    st.title("🔧 Ferramentas")
    with st.form("f_ferr", clear_on_submit=True):
        n = st.text_input("Nome")
        q = st.number_input("Qtd", min_value=1)
        s = st.text_input("N° Série")
        if st.form_submit_button("Gravar") and n:
            c.execute("INSERT INTO ferramentas (nome, quantidade, num_serie) VALUES (?,?,?)", (n, q, s))
            conn.commit()
            st.success("Salvo!")
            st.rerun()

    df_ferr = pd.read_sql("SELECT * FROM ferramentas", conn)
    st.dataframe(df_ferr, use_container_width=True)
    if not df_ferr.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        ferr_para_excluir = st.selectbox("Selecione a Ferramenta para remover:", df_ferr['nome'].unique(), key="del_ferr")
        if st.button("Excluir Ferramenta", type="primary", key="btn_del_ferr"):
            c.execute("DELETE FROM ferramentas WHERE nome=?", (ferr_para_excluir,))
            conn.commit()
            st.success("Ferramenta removida!")
            st.rerun()


    if not df_ferr.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        ferr_para_excluir = st.selectbox("Selecione a Ferramenta para remover:", df_ferr['nome'].unique())
        if st.button("Excluir Ferramenta", type="primary"):
            c.execute("DELETE FROM ferramentas WHERE nome=?", (ferr_para_excluir,))
            conn.commit()
            st.success("Ferramenta removida!")
            st.rerun()

# ----------------- 6. BOBINAS -----------------
elif menu == "🧻 Bobinas":
    st.title("🧻 Estoque de Bobinas")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📝 Ajuste Manual")
        t_sel = st.selectbox("Tamanho:", bobinas)
        v_atual = int(
            pd.read_sql("SELECT quantidade FROM bobinas WHERE tamanho=?", conn, params=[t_sel])['quantidade'].iloc[0])
        q_nova = st.number_input("Qtd exata:", value=v_atual, min_value=0)
        if st.button("Atualizar Estoque"):
            c.execute("UPDATE bobinas SET quantidade=? WHERE tamanho=?", (q_nova, t_sel))
            conn.commit()
            st.success("Ajustado!")
            st.rerun()
    with col2:
        st.subheader("🔄 Registrar Troca (-1)")
        term_sel = st.selectbox("Terminal:", terminais)
        t_troca = st.selectbox("Bobina Usada:", bobinas, key="swp")
        if st.button("Confirmar Troca"):
            qtd = int(pd.read_sql("SELECT quantidade FROM bobinas WHERE tamanho=?", conn, params=[t_troca])[
                          'quantidade'].iloc[0])
            if qtd > 0:
                c.execute("UPDATE bobinas SET quantidade = quantidade - 1 WHERE tamanho=?", (t_troca,))
                conn.commit()
                st.success("Trocado!")
                st.rerun()
            else:
                st.error("Zerado!")
    st.write("---")
    st.table(pd.read_sql("SELECT * FROM bobinas", conn))

# ----------------- 7. HT -----------------
elif menu == "📻 HT":
    st.title("📻 Rádios HT")
    with st.form("f_ht"):
        l_sel = st.selectbox("Local:", hts)
        q = st.number_input("Qtd HTs", min_value=0)
        s = st.selectbox("Status", ["Disponível", "Em Uso", "Manutenção"])
        if st.form_submit_button("Atualizar"):
            c.execute("UPDATE controle_ht SET quantidade=?, status=? WHERE local=?", (q, s, l_sel))
            conn.commit()
            st.success("Atualizado!")
            st.rerun()
    st.dataframe(pd.read_sql("SELECT * FROM controle_ht", conn), use_container_width=True)

# ----------------- 8. CHAMADOS -----------------
elif menu == "🎫 Chamados":
    st.title("🎫 Chamados")
    with st.form("f_ch", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1: num_ch = st.text_input("N° Chamado")
        with c2: empresa_sel = st.selectbox("Empresa:", ["Prime", "WPS", "PERTO", "WG", "CLUB CAR"])
        with c3: status_ch = st.selectbox("Status:", ["ABERTO", "FECHADO"])
        desc_ch = st.text_area("Descrição do Chamado")
        if st.form_submit_button("Gravar Chamado") and num_ch:
            c.execute("INSERT INTO chamados (num_chamado, empresa, status, descricao) VALUES (?,?,?,?)",
                      (num_ch, empresa_sel, status_ch, desc_ch))
            conn.commit()
            st.success("Chamado gravado!")
            st.rerun()

    df_chamados = pd.read_sql("SELECT * FROM chamados", conn)
    st.dataframe(df_chamados, use_container_width=True)
    df_chamados = pd.read_sql("SELECT * FROM chamados", conn)
    st.dataframe(df_chamados, use_container_width=True)

    # O código da lixeira entra exatamente aqui:
    if not df_chamados.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        chamado_para_excluir = st.selectbox("Selecione o N° do Chamado para remover:", df_chamados['num_chamado'].unique(), key="del_ch")
        if st.button("Excluir Chamado", type="primary", key="btn_del_ch"):
            c.execute("DELETE FROM chamados WHERE num_chamado=?", (chamado_para_excluir,))
            conn.commit()
            st.success("Chamado removido!")
            st.rerun()


    if not df_chamados.empty:
        st.write("---")
        st.subheader("🗑️ Excluir Registro")
        chamado_para_excluir = st.selectbox("Selecione o N° do Chamado para remover:",
                                            df_chamados['num_chamado'].unique())
        if st.button("Excluir Chamado", type="primary"):
            c.execute("DELETE FROM chamados WHERE num_chamado=?", (chamado_para_excluir,))
            conn.commit()
            st.success("Chamado removido!")
            st.rerun()
