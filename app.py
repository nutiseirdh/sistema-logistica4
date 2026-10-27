# Sistema de Logística 4.0 — aplicação demonstrativa/local
# Tecnologias open-source: Streamlit, SQLite, pandas, openpyxl, Plotly, ReportLab.
# Execute: pip install -r requirements.txt && streamlit run app.py

from pathlib import Path
from datetime import datetime, date, time
import sqlite3
import io
import uuid
import pandas as pd
import streamlit as st
import plotly.express as px
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "dados"
UPLOAD_DIR = DATA_DIR / "documentos"
PHOTO_DIR = DATA_DIR / "fotos"
DB_PATH = DATA_DIR / "logistica.db"
for folder in (DATA_DIR, UPLOAD_DIR, PHOTO_DIR):
    folder.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="Logística 4.0", page_icon="📦", layout="wide")
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
[data-testid="stMetric"] {border: 1px solid #e5e7eb; padding: 14px; border-radius: 12px;}
.small-muted {color:#6b7280;font-size:.9rem}
</style>
""", unsafe_allow_html=True)

def connect():
    con = sqlite3.connect(DB_PATH, timeout=20)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    return con

def init_db():
    with connect() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS items (
            id INTEGER PRIMARY KEY AUTOINCREMENT, sku TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL, category TEXT DEFAULT '', unit TEXT DEFAULT 'un',
            quantity REAL NOT NULL DEFAULT 0, min_quantity REAL NOT NULL DEFAULT 0,
            location TEXT DEFAULT '', photo_path TEXT DEFAULT '', notes TEXT DEFAULT '',
            active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT, item_id INTEGER NOT NULL,
            movement_type TEXT NOT NULL CHECK(movement_type IN ('Entrada','Saída')),
            quantity REAL NOT NULL CHECK(quantity > 0), reason TEXT NOT NULL,
            responsible TEXT NOT NULL, event_at TEXT NOT NULL,
            document_path TEXT DEFAULT '',
            FOREIGN KEY(item_id) REFERENCES items(id)
        );
        CREATE TABLE IF NOT EXISTS vehicles (
            id INTEGER PRIMARY KEY AUTOINCREMENT, plate TEXT NOT NULL UNIQUE,
            model TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Disponível',
            odometer REAL DEFAULT 0, notes TEXT DEFAULT '', active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS trips (
            id INTEGER PRIMARY KEY AUTOINCREMENT, vehicle_id INTEGER NOT NULL,
            destination TEXT NOT NULL, driver TEXT NOT NULL, departure_at TEXT NOT NULL,
            return_at TEXT DEFAULT '', distance_km REAL DEFAULT 0, status TEXT DEFAULT 'Em rota',
            notes TEXT DEFAULT '', FOREIGN KEY(vehicle_id) REFERENCES vehicles(id)
        );
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
            registration TEXT UNIQUE, role TEXT NOT NULL, shift TEXT DEFAULT '',
            permission TEXT NOT NULL DEFAULT 'Consulta', email TEXT DEFAULT '',
            active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL,
            doc_type TEXT NOT NULL, file_path TEXT NOT NULL, uploaded_by TEXT DEFAULT '',
            uploaded_at TEXT NOT NULL, notes TEXT DEFAULT ''
        );
        """)
init_db()

def query_df(sql, params=()):
    with connect() as con:
        return pd.read_sql_query(sql, con, params=params)

def execute(sql, params=()):
    with connect() as con:
        cur = con.execute(sql, params)
        con.commit()
        return cur.lastrowid

def save_upload(upload, folder):
    if upload is None:
        return ""
    suffix = Path(upload.name).suffix.lower()[:12]
    safe_name = f"{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}{suffix}"
    target = folder / safe_name
    target.write_bytes(upload.getvalue())
    return str(target.relative_to(APP_DIR))

def full_path(relative):
    return APP_DIR / relative if relative else None

def excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()

def pdf_bytes(title, df):
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 12)]
    if df.empty:
        story.append(Paragraph("Nenhum registro encontrado.", styles["Normal"]))
    else:
        view = df.fillna("").astype(str)
        data = [list(view.columns)] + view.values.tolist()
        widths = [max(55, (landscape(A4)[0] - 48) / max(1, len(data[0]))) for _ in data[0]]
        table = Table(data, repeatRows=1, colWidths=widths)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1f4e5f")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), 7),
            ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#cbd5e1")),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f1f5f9")]),
        ]))
        story.append(table)
    doc.build(story)
    return output.getvalue()

def export_buttons(prefix, title, df):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Exportar Excel", excel_bytes({title[:31]: df}), f"{prefix}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c2:
        st.download_button("⬇️ Exportar PDF", pdf_bytes(title, df), f"{prefix}.pdf",
                           "application/pdf", use_container_width=True)

def current_items():
    return query_df("SELECT id, sku, name, unit, quantity, min_quantity, location FROM items WHERE active=1 ORDER BY name")

def active_people():
    return query_df("SELECT name FROM employees WHERE active=1 ORDER BY name")["name"].tolist()

st.sidebar.title("📦 Logística 4.0")
st.sidebar.caption("Gestão local • código aberto")
page = st.sidebar.radio("Navegação", [
    "Dashboard", "Estoque", "Movimentações", "Veículos e viagens",
    "Colaboradores", "Documentos", "Relatórios"
])
st.sidebar.divider()
st.sidebar.caption("Banco SQLite local: dados/logistica.db")
st.sidebar.caption("Faça cópias de segurança da pasta `dados`.")

if page == "Dashboard":
    st.title("Dashboard operacional")
    items = query_df("SELECT * FROM items WHERE active=1")
    low = items[items["quantity"] <= items["min_quantity"]] if not items.empty else items
    vehicles = query_df("SELECT * FROM vehicles WHERE active=1")
    moving = query_df("SELECT * FROM movements")
    trips = query_df("SELECT * FROM trips")
    a,b,c,d = st.columns(4)
    a.metric("Itens cadastrados", len(items))
    b.metric("Itens com estoque baixo", len(low))
    c.metric("Veículos em rota", int((vehicles["status"] == "Em rota").sum()) if not vehicles.empty else 0)
    d.metric("Movimentações registradas", len(moving))
    left, right = st.columns(2)
    with left:
        st.subheader("Saldo por item")
        if not items.empty:
            fig = px.bar(items.sort_values("quantity", ascending=False).head(12), x="name", y="quantity",
                         labels={"name":"Material","quantity":"Saldo"}, title="12 maiores saldos")
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("Cadastre materiais para visualizar os gráficos.")
    with right:
        st.subheader("Movimentações por tipo")
        if not moving.empty:
            by_type = moving.groupby("movement_type", as_index=False)["quantity"].sum()
            fig = px.pie(by_type, names="movement_type", values="quantity", hole=.45)
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("Ainda não há movimentações registradas.")
    st.subheader("⚠️ Alertas de estoque")
    if low.empty:
        st.success("Nenhum item abaixo ou igual ao estoque mínimo.")
    else:
        st.dataframe(low[["sku","name","quantity","min_quantity","location"]].rename(columns={
            "sku":"SKU","name":"Material","quantity":"Saldo","min_quantity":"Mínimo","location":"Localização"
        }), use_container_width=True, hide_index=True)
    if not trips.empty:
        st.subheader("Viagens recentes")
        st.dataframe(trips.sort_values("id", ascending=False).head(5), use_container_width=True, hide_index=True)

elif page == "Estoque":
    st.title("Controle de estoque")
    tab1, tab2 = st.tabs(["Cadastrar material", "Consultar / atualizar"])
    with tab1:
        with st.form("new_item", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            sku = c1.text_input("SKU *")
            name = c2.text_input("Nome do material *")
            category = c3.text_input("Categoria")
            c4,c5,c6 = st.columns(3)
            unit = c4.selectbox("Unidade", ["un","cx","kg","g","L","m","m²","pct","rolo","par"])
            qty = c5.number_input("Quantidade inicial", min_value=0.0, step=1.0)
            minimum = c6.number_input("Quantidade mínima", min_value=0.0, step=1.0)
            location = st.text_input("Localização no galpão (ex.: A-02-Prateleira 3)")
            notes = st.text_area("Observações")
            photo = st.file_uploader("Foto do material (opcional)", type=["png","jpg","jpeg","webp"], key="item_photo")
            submitted = st.form_submit_button("Salvar material", type="primary")
            if submitted:
                if not sku.strip() or not name.strip():
                    st.error("Preencha SKU e nome do material.")
                else:
                    try:
                        photo_path = save_upload(photo, PHOTO_DIR)
                        execute("""INSERT INTO items(sku,name,category,unit,quantity,min_quantity,location,photo_path,notes,created_at)
                                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                                 (sku.strip().upper(),name.strip(),category.strip(),unit,qty,minimum,location.strip(),photo_path,notes.strip(),datetime.now().isoformat(timespec="seconds")))
                        st.success("Material cadastrado.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Este SKU já existe. Use um SKU único.")
    with tab2:
        df = query_df("SELECT id,sku,name,category,unit,quantity,min_quantity,location,photo_path,notes FROM items WHERE active=1 ORDER BY name")
        if df.empty: st.info("Nenhum material cadastrado.")
        else:
            search = st.text_input("Buscar por SKU, nome ou localização")
            if search:
                df = df[df.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)]
            st.dataframe(df.drop(columns=["photo_path","notes"]), use_container_width=True, hide_index=True)
            selected_id = st.selectbox("Ver detalhes do item", df["id"].tolist(), format_func=lambda x: f'{df.loc[df.id==x, "sku"].iloc[0]} — {df.loc[df.id==x, "name"].iloc[0]}')
            selected = df[df["id"]==selected_id].iloc[0]
            if selected["photo_path"] and full_path(selected["photo_path"]).exists():
                st.image(str(full_path(selected["photo_path"])), caption=selected["name"], width=300)
            with st.form("update_item"):
                st.subheader("Editar dados do material")
                e1,e2,e3 = st.columns(3)
                new_name = e1.text_input("Nome", selected["name"])
                new_min = e2.number_input("Estoque mínimo", min_value=0.0, value=float(selected["min_quantity"]), step=1.0)
                new_loc = e3.text_input("Localização", selected["location"] or "")
                new_notes = st.text_area("Observações", selected["notes"] or "")
                new_photo = st.file_uploader("Substituir foto (opcional)", type=["png","jpg","jpeg","webp"], key="edit_photo")
                save = st.form_submit_button("Atualizar material")
                if save:
                    photo_path = save_upload(new_photo, PHOTO_DIR) if new_photo else selected["photo_path"]
                    execute("UPDATE items SET name=?,min_quantity=?,location=?,notes=?,photo_path=? WHERE id=?",
                            (new_name.strip(),new_min,new_loc.strip(),new_notes.strip(),photo_path,int(selected_id)))
                    st.success("Material atualizado.")
                    st.rerun()
            if st.button("Inativar material selecionado", type="secondary"):
                execute("UPDATE items SET active=0 WHERE id=?", (int(selected_id),))
                st.warning("Material inativado (histórico preservado).")
                st.rerun()

elif page == "Movimentações":
    st.title("Entrada e saída de materiais")
    items = current_items()
    if items.empty:
        st.warning("Cadastre pelo menos um material antes de registrar movimentações.")
    else:
        with st.form("movement", clear_on_submit=True):
            item_id = st.selectbox("Material *", items["id"].tolist(),
                format_func=lambda x: f'{items.loc[items.id==x, "sku"].iloc[0]} — {items.loc[items.id==x, "name"].iloc[0]} (saldo: {items.loc[items.id==x, "quantity"].iloc[0]:g})')
            typ = st.radio("Tipo", ["Entrada","Saída"], horizontal=True)
            quantity = st.number_input("Quantidade *", min_value=0.01, step=1.0)
            reason = st.text_input("Motivo / referência *", placeholder="Ex.: recebimento NF 123 ou requisição interna")
            people = active_people()
            responsible = st.selectbox("Responsável", people + ["Outro / não cadastrado"]) if people else st.text_input("Responsável *")
            if people and responsible == "Outro / não cadastrado":
                responsible = st.text_input("Informe o responsável", key="other_responsible")
            document = st.file_uploader("Documento relacionado (opcional)", type=["pdf","png","jpg","jpeg"], key="movement_doc")
            submitted = st.form_submit_button("Registrar movimentação", type="primary")
            if submitted:
                if not reason.strip() or not str(responsible).strip():
                    st.error("Informe o motivo e o responsável.")
                else:
                    item = items[items.id==item_id].iloc[0]
                    current = float(item["quantity"])
                    if typ == "Saída" and quantity > current:
                        st.error(f"Saldo insuficiente. Saldo atual: {current:g} {item['unit']}.")
                    else:
                        doc_path = save_upload(document, UPLOAD_DIR)
                        with connect() as con:
                            con.execute("INSERT INTO movements(item_id,movement_type,quantity,reason,responsible,event_at,document_path) VALUES(?,?,?,?,?,?,?)",
                                        (int(item_id),typ,float(quantity),reason.strip(),str(responsible).strip(),datetime.now().isoformat(timespec="seconds"),doc_path))
                            delta = quantity if typ == "Entrada" else -quantity
                            con.execute("UPDATE items SET quantity=quantity+? WHERE id=?", (delta,int(item_id)))
                            con.commit()
                        st.success("Movimentação registrada e saldo atualizado.")
                        st.rerun()
    st.subheader("Histórico de movimentações")
    hist = query_df("""SELECT m.id, i.sku AS SKU, i.name AS Material, m.movement_type AS Tipo,
                       m.quantity AS Quantidade, i.unit AS Unidade, m.reason AS Motivo,
                       m.responsible AS Responsável, m.event_at AS Data_hora, m.document_path AS Documento
                       FROM movements m JOIN items i ON i.id=m.item_id ORDER BY m.id DESC""")
    st.dataframe(hist.drop(columns=["Documento"]) if not hist.empty else hist, use_container_width=True, hide_index=True)
    if not hist.empty:
        chosen = st.selectbox("Documento da movimentação", hist["id"].tolist(), format_func=lambda x: f"Mov. {x} — {hist.loc[hist.id==x, 'Material'].iloc[0]}")
        path = hist.loc[hist.id==chosen, "Documento"].iloc[0]
        if path and full_path(path).exists():
            p = full_path(path)
            st.download_button("Baixar documento da movimentação", p.read_bytes(), p.name)

elif page == "Veículos e viagens":
    st.title("Controle de frota e viagens")
    tab1,tab2,tab3 = st.tabs(["Cadastrar veículo","Registrar viagem","Frota / histórico"])
    with tab1:
        with st.form("vehicle", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            plate = c1.text_input("Placa *").upper()
            model = c2.text_input("Modelo *")
            status = c3.selectbox("Status", ["Disponível","Em rota","Manutenção"])
            odometer = st.number_input("Odômetro (km)", min_value=0.0, step=1.0)
            notes = st.text_area("Observações")
            if st.form_submit_button("Salvar veículo"):
                try:
                    if not plate.strip() or not model.strip(): st.error("Placa e modelo são obrigatórios.")
                    else:
                        execute("INSERT INTO vehicles(plate,model,status,odometer,notes) VALUES(?,?,?,?,?)",
                                (plate.strip(),model.strip(),status,odometer,notes.strip()))
                        st.success("Veículo cadastrado."); st.rerun()
                except sqlite3.IntegrityError: st.error("Placa já cadastrada.")
    with tab2:
        fleet = query_df("SELECT id,plate,model,status FROM vehicles WHERE active=1 ORDER BY plate")
        if fleet.empty: st.info("Cadastre um veículo primeiro.")
        else:
            with st.form("trip", clear_on_submit=True):
                vehicle_id = st.selectbox("Veículo", fleet["id"].tolist(), format_func=lambda x: f'{fleet.loc[fleet.id==x,"plate"].iloc[0]} — {fleet.loc[fleet.id==x,"model"].iloc[0]}')
                destination = st.text_input("Destino / rota *")
                driver = st.text_input("Condutor responsável *")
                departure_date = st.date_input("Data de saída", value=date.today())
                departure_time = st.time_input("Hora de saída", value=datetime.now().time().replace(second=0, microsecond=0))
                departure = datetime.combine(departure_date, departure_time)
                distance = st.number_input("Distância estimada/real (km)", min_value=0.0, step=1.0)
                trip_notes = st.text_area("Observações")
                if st.form_submit_button("Iniciar viagem", type="primary"):
                    if not destination.strip() or not driver.strip(): st.error("Informe destino e condutor.")
                    else:
                        with connect() as con:
                            con.execute("INSERT INTO trips(vehicle_id,destination,driver,departure_at,distance_km,status,notes) VALUES(?,?,?,?,?,'Em rota',?)",
                                        (int(vehicle_id),destination.strip(),driver.strip(),departure.isoformat(timespec="seconds"),distance,trip_notes.strip()))
                            con.execute("UPDATE vehicles SET status='Em rota' WHERE id=?", (int(vehicle_id),))
                            con.commit()
                        st.success("Viagem iniciada."); st.rerun()
    with tab3:
        fleet = query_df("SELECT * FROM vehicles WHERE active=1 ORDER BY plate")
        if not fleet.empty:
            st.dataframe(fleet.drop(columns=["active"]), use_container_width=True, hide_index=True)
            vehicle_id = st.selectbox("Veículo para atualização", fleet["id"].tolist(), format_func=lambda x: fleet.loc[fleet.id==x,"plate"].iloc[0])
            vehicle = fleet[fleet.id==vehicle_id].iloc[0]
            new_status = st.selectbox("Atualizar status", ["Disponível","Em rota","Manutenção"], index=["Disponível","Em rota","Manutenção"].index(vehicle["status"]))
            if st.button("Salvar status do veículo"):
                execute("UPDATE vehicles SET status=? WHERE id=?", (new_status,int(vehicle_id)))
                st.success("Status atualizado."); st.rerun()
        trips = query_df("""SELECT t.id, v.plate AS Placa, t.destination AS Destino, t.driver AS Condutor,
                            t.departure_at AS Saída, t.return_at AS Retorno, t.distance_km AS Km,
                            t.status AS Status, t.notes AS Observações
                            FROM trips t JOIN vehicles v ON v.id=t.vehicle_id ORDER BY t.id DESC""")
        st.subheader("Histórico de viagens")
        st.dataframe(trips, use_container_width=True, hide_index=True)
        if not trips.empty:
            active_trips = trips[trips["Status"]=="Em rota"]
            if not active_trips.empty:
                trip_id = st.selectbox("Finalizar viagem", active_trips["id"].tolist(), format_func=lambda x: f"Viagem {x} — {active_trips.loc[active_trips.id==x,'Placa'].iloc[0]}")
                ret_km = st.number_input("Km percorridos na viagem", min_value=0.0, step=1.0, key="ret_km")
                if st.button("Registrar retorno"):
                    with connect() as con:
                        row = con.execute("SELECT vehicle_id FROM trips WHERE id=?", (int(trip_id),)).fetchone()
                        con.execute("UPDATE trips SET status='Concluída',return_at=?,distance_km=? WHERE id=?",
                                    (datetime.now().isoformat(timespec="seconds"),ret_km,int(trip_id)))
                        con.execute("UPDATE vehicles SET status='Disponível',odometer=odometer+? WHERE id=?", (ret_km,row["vehicle_id"]))
                        con.commit()
                    st.success("Retorno registrado."); st.rerun()

elif page == "Colaboradores":
    st.title("Cadastro de colaboradores e permissões")
    st.caption("Este módulo registra perfis, mas a autenticação real ainda precisa ser configurada antes de uso em produção.")
    with st.form("employee", clear_on_submit=True):
        c1,c2,c3 = st.columns(3)
        name = c1.text_input("Nome completo *")
        registration = c2.text_input("Matrícula / identificador")
        role = c3.text_input("Cargo / função *")
        c4,c5,c6 = st.columns(3)
        shift = c4.selectbox("Turno", ["Administrativo","Manhã","Tarde","Noite","Escala","Outro"])
        permission = c5.selectbox("Perfil de acesso", ["Administrador","Gestor","Operador","Consulta"])
        email = c6.text_input("E-mail")
        if st.form_submit_button("Cadastrar colaborador", type="primary"):
            if not name.strip() or not role.strip(): st.error("Nome e cargo são obrigatórios.")
            else:
                try:
                    execute("INSERT INTO employees(name,registration,role,shift,permission,email) VALUES(?,?,?,?,?,?)",
                            (name.strip(),registration.strip() or None,role.strip(),shift,permission,email.strip()))
                    st.success("Colaborador cadastrado."); st.rerun()
                except sqlite3.IntegrityError: st.error("Matrícula já cadastrada.")
    people = query_df("SELECT id,name,registration,role,shift,permission,email,active FROM employees ORDER BY name")
    st.subheader("Colaboradores cadastrados")
    st.dataframe(people, use_container_width=True, hide_index=True)
    if not people.empty:
        person_id = st.selectbox("Selecionar colaborador para ativar/inativar", people["id"].tolist(),
                                 format_func=lambda x: people.loc[people.id==x,"name"].iloc[0])
        is_active = int(people.loc[people.id==person_id,"active"].iloc[0])
        if st.button("Inativar colaborador" if is_active else "Reativar colaborador"):
            execute("UPDATE employees SET active=? WHERE id=?", (0 if is_active else 1,int(person_id)))
            st.rerun()

elif page == "Documentos":
    st.title("Central de documentos")
    st.write("Armazene guias de transporte, canhotos, notas fiscais, ordens de serviço e outros anexos.")
    with st.form("document_form", clear_on_submit=True):
        title = st.text_input("Título / identificação *")
        doc_type = st.selectbox("Tipo", ["Guia de transporte","Canhoto","Nota fiscal","Ordem de serviço","Comprovante","Outro"])
        uploader = st.text_input("Enviado por")
        notes = st.text_area("Observações")
        file = st.file_uploader("Arquivo *", type=["pdf","png","jpg","jpeg","xlsx","xls","docx","doc","csv","txt"])
        if st.form_submit_button("Salvar documento", type="primary"):
            if not title.strip() or file is None: st.error("Informe o título e selecione um arquivo.")
            else:
                path = save_upload(file, UPLOAD_DIR)
                execute("INSERT INTO documents(title,doc_type,file_path,uploaded_by,uploaded_at,notes) VALUES(?,?,?,?,?,?)",
                        (title.strip(),doc_type,path,uploader.strip(),datetime.now().isoformat(timespec="seconds"),notes.strip()))
                st.success("Documento armazenado."); st.rerun()
    docs = query_df("SELECT id,title,doc_type,file_path,uploaded_by,uploaded_at,notes FROM documents ORDER BY id DESC")
    st.subheader("Documentos disponíveis")
    if docs.empty: st.info("Nenhum documento anexado.")
    else:
        for _, row in docs.iterrows():
            with st.container(border=True):
                a,b = st.columns([4,1])
                a.markdown(f"**{row['title']}**  \n{row['doc_type']} • {row['uploaded_at']} • {row['uploaded_by']}")
                if row["notes"]: a.caption(row["notes"])
                path = full_path(row["file_path"])
                if path and path.exists():
                    b.download_button("Baixar arquivo", path.read_bytes(), path.name, key=f"doc_{row['id']}", use_container_width=True)

elif page == "Relatórios":
    st.title("Relatórios e exportações")
    items = query_df("SELECT sku AS SKU,name AS Material,category AS Categoria,unit AS Unidade,quantity AS Saldo,min_quantity AS Estoque_minimo,location AS Localização FROM items WHERE active=1 ORDER BY name")
    moves = query_df("""SELECT m.id AS ID,i.sku AS SKU,i.name AS Material,m.movement_type AS Tipo,m.quantity AS Quantidade,
                        i.unit AS Unidade,m.reason AS Motivo,m.responsible AS Responsável,m.event_at AS Data_hora
                        FROM movements m JOIN items i ON i.id=m.item_id ORDER BY m.event_at DESC""")
    fleet = query_df("SELECT plate AS Placa,model AS Modelo,status AS Status,odometer AS Odômetro_km,notes AS Observações FROM vehicles WHERE active=1 ORDER BY plate")
    trips = query_df("""SELECT v.plate AS Placa,t.destination AS Destino,t.driver AS Condutor,t.departure_at AS Saída,
                        t.return_at AS Retorno,t.distance_km AS Distância_km,t.status AS Status
                        FROM trips t JOIN vehicles v ON v.id=t.vehicle_id ORDER BY t.departure_at DESC""")
    tab1,tab2,tab3,tab4 = st.tabs(["Estoque","Movimentações","Frota","Viagens"])
    with tab1:
        st.dataframe(items, use_container_width=True, hide_index=True)
        export_buttons("relatorio_estoque","Relatório de estoque",items)
    with tab2:
        st.dataframe(moves, use_container_width=True, hide_index=True)
        export_buttons("relatorio_movimentacoes","Relatório de movimentações",moves)
    with tab3:
        st.dataframe(fleet, use_container_width=True, hide_index=True)
        export_buttons("relatorio_frota","Relatório da frota",fleet)
    with tab4:
        st.dataframe(trips, use_container_width=True, hide_index=True)
        export_buttons("relatorio_viagens","Relatório de viagens",trips)

st.sidebar.divider()
st.sidebar.caption("Versão inicial • uso interno/local")
