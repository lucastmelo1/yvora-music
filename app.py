
import os
import json
import time
import re
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

APP_TITLE = "Yvora Music"
BRAND_BG = "#EFE7DD"
BRAND_BLUE = "#0E2A47"


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def fmt_mmss(total_seconds: Any) -> str:
    try:
        total_seconds = int(float(total_seconds or 0))
    except Exception:
        total_seconds = 0
    total_seconds = max(total_seconds, 0)
    m = total_seconds // 60
    s = total_seconds % 60
    return f"{m:02d}:{s:02d}"


def app_bootstrap() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(
        f"""
        <style>
          html, body, [class*="css"] {{ background: {BRAND_BG} !important; }}
          .yv-title {{ color: {BRAND_BLUE}; font-family: Georgia, "Times New Roman", serif; letter-spacing: 0.5px; margin: 0; }}
          .divider {{
            height: 1px;
            background: linear-gradient(to right, rgba(14,42,71,0), rgba(14,42,71,0.35), rgba(14,42,71,0));
            margin: 10px 0 14px 0;
          }}
          .meta {{ color: rgba(14,42,71,0.70); font-size: 12px; }}
          .mobile-wrap {{ max-width: 560px; margin: 0 auto; }}
          .show-card {{
            background: rgba(255,255,255,0.55);
            border: 1px solid rgba(14,42,71,0.14);
            border-radius: 16px;
            padding: 14px 14px;
            margin-bottom: 10px;
          }}
          .show-h1 {{ font-family: Georgia, "Times New Roman", serif; color: {BRAND_BLUE}; font-size: 22px; margin: 0 0 6px 0; }}
          .show-muted {{ color: rgba(14,42,71,0.70); font-size: 13px; }}
          .stButton > button {{
            padding: 0.20rem 0.60rem !important;
            min-height: 2.0rem !important;
            border-radius: 999px !important;
            font-size: 13px !important;
            white-space: nowrap !important;
          }}
          .yv-badge {{
            display:inline-block;
            padding:4px 10px;
            border-radius:999px;
            background: rgba(14,42,71,0.08);
            color: {BRAND_BLUE};
            font-size:12px;
            margin-right:6px;
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header(subtitle: str) -> None:
    col1, col2 = st.columns([1, 4])
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")

    with col1:
        if os.path.exists(logo_path):
            st.image(logo_path, width=90)

    with col2:
        st.markdown(f'<h2 class="yv-title">{APP_TITLE}</h2>', unsafe_allow_html=True)
        st.markdown(f'<div class="meta">{subtitle}</div>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)


@st.cache_resource
def get_sheet():
    import gspread
    from google.oauth2.service_account import Credentials

    try:
        sheet_id = st.secrets["google"]["sheet_id"]
    except Exception:
        raise RuntimeError("Falta [google].sheet_id nos Secrets")

    sa_info: Dict[str, Any] = {}
    google_block = st.secrets.get("google", {})
    if google_block.get("service_account_json"):
        try:
            sa_info = json.loads(google_block["service_account_json"])
        except Exception as e:
            raise RuntimeError(f"google.service_account_json inválido: {e}")
    elif "gcp_service_account" in st.secrets:
        try:
            sa_info = dict(st.secrets["gcp_service_account"])
        except Exception as e:
            raise RuntimeError(f"[gcp_service_account] inválido: {e}")
    else:
        raise RuntimeError("Use [google].service_account_json ou [gcp_service_account] nos Secrets.")

    private_key = sa_info.get("private_key", "")
    if not private_key:
        raise RuntimeError("private_key ausente nas credenciais.")
    if "\\n" in private_key and "\n" not in private_key:
        sa_info["private_key"] = private_key.replace("\\n", "\n")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(sheet_id)


def ws_get_or_create(title: str, headers: List[str]):
    sh = get_sheet()
    try:
        ws = sh.worksheet(title)
        row1 = ws.row_values(1)
        if headers and not row1:
            ws.append_row(headers)
        return ws
    except Exception:
        ws = sh.add_worksheet(title=title, rows=5000, cols=max(12, len(headers) or 12))
        if headers:
            ws.append_row(headers)
        return ws


def ws_read_df(title: str) -> pd.DataFrame:
    ws = ws_get_or_create(title, [])
    values = ws.get_all_values()
    if not values or not values[0]:
        return pd.DataFrame()
    header = values[0]
    rows = values[1:]
    return pd.DataFrame(rows, columns=header)


def ws_append_row(title: str, headers: List[str], row: List[Any]) -> None:
    ws = ws_get_or_create(title, headers)
    ws.append_row([("" if v is None else str(v)) for v in row])


def ws_replace_session_rows(title: str, headers: List[str], session_id: str, new_rows: List[List[Any]]) -> None:
    ws = ws_get_or_create(title, headers)
    allv = ws.get_all_values()
    if not allv:
        ws.append_row(headers)
        allv = [headers]

    header = allv[0]
    if header != headers:
        ws.clear()
        ws.append_row(headers)
        header = headers

    kept = [header]
    for r in allv[1:]:
        if not r:
            continue
        if r[0] != session_id:
            kept.append(r)

    ws.clear()
    ws.append_rows(kept)
    if new_rows:
        ws.append_rows([[("" if v is None else str(v)) for v in row] for row in new_rows])


USERS_HEADERS = ["username", "password", "role", "active"]
SESS_HEADERS = ["session_id", "title", "genre", "total_duration_min", "status", "updated_at_utc"]
CH_HEADERS = [
    "session_id", "chapter_index", "moment_key", "chapter_type", "planned_duration_sec",
    "music_title", "music_artist", "link_context_input",
    "ai_story_text", "ai_music_alternatives", "ai_experience_actions", "ai_notes",
    "updated_at_utc",
]
LIVE_HEADERS = ["session_id", "current_chapter_index", "updated_at_utc"]
HIST_HEADERS = [
    "ts_utc", "session_id", "chapter_index", "moment_key", "music_title", "music_artist",
    "link_context_input", "ai_story_text", "ai_music_alternatives", "ai_experience_actions", "ai_notes",
]


def auth_login(username: str, password: str) -> Optional[Dict[str, str]]:
    df = ws_read_df("users")
    if df.empty:
        return None
    u = (username or "").strip()
    p = (password or "").strip()
    for _, r in df.iterrows():
        active = str(r.get("active", "1")).strip()
        if active in ["0", "False", "false"]:
            continue
        if str(r.get("username", "")).strip() == u and str(r.get("password", "")).strip() == p:
            role = str(r.get("role", "staff")).strip() or "staff"
            return {"username": u, "role": role}
    return None


def logout() -> None:
    st.session_state.pop("user", None)
    st.rerun()


def list_sessions_sheet() -> pd.DataFrame:
    df = ws_read_df("sessions")
    if df.empty:
        return df
    if "updated_at_utc" in df.columns:
        df = df.sort_values("updated_at_utc", ascending=False)
    return df


def get_session_row(session_id: str) -> Optional[Dict[str, Any]]:
    df = list_sessions_sheet()
    if df.empty or "session_id" not in df.columns:
        return None
    m = df[df["session_id"] == session_id]
    if m.empty:
        return None
    return m.iloc[0].to_dict()


def upsert_session(session_id: str, title: str, genre: str, total_min: int, status: str) -> None:
    ws = ws_get_or_create("sessions", SESS_HEADERS)
    values = ws.get_all_values()
    if not values:
        ws.append_row(SESS_HEADERS)
        values = [SESS_HEADERS]
    header = values[0]
    idx_key = header.index("session_id")
    found = None
    for i, r in enumerate(values[1:], start=2):
        if len(r) > idx_key and r[idx_key] == session_id:
            found = i
            break
    row = {
        "session_id": session_id,
        "title": title,
        "genre": genre,
        "total_duration_min": str(int(total_min)),
        "status": status,
        "updated_at_utc": utc_now().isoformat(),
    }
    out = [row.get(h, "") for h in header]
    if found:
        ws.update(f"A{found}", [out])
    else:
        ws.append_row(out)


def read_chapters(session_id: str) -> pd.DataFrame:
    df = ws_read_df("chapters")
    if df.empty or "session_id" not in df.columns:
        return pd.DataFrame()
    df = df[df["session_id"] == session_id].copy()
    if df.empty:
        return df
    if "chapter_index" not in df.columns:
        df["chapter_index"] = range(len(df))
    if "planned_duration_sec" not in df.columns:
        df["planned_duration_sec"] = 300
    df["chapter_index"] = pd.to_numeric(df["chapter_index"], errors="coerce").fillna(0).astype(int)
    return df.sort_values("chapter_index").reset_index(drop=True)


def save_chapters(session_id: str, df: pd.DataFrame) -> None:
    rows = []
    for _, r in df.iterrows():
        rows.append([
            session_id,
            int(pd.to_numeric(r.get("chapter_index", 0), errors="coerce") or 0),
            r.get("moment_key", "aquecimento"),
            r.get("chapter_type", "music"),
            int(pd.to_numeric(r.get("planned_duration_sec", 300), errors="coerce") or 300),
            r.get("music_title", ""),
            r.get("music_artist", ""),
            r.get("link_context_input", ""),
            r.get("ai_story_text", ""),
            r.get("ai_music_alternatives", ""),
            r.get("ai_experience_actions", ""),
            r.get("ai_notes", ""),
            utc_now().isoformat(),
        ])
    ws_replace_session_rows("chapters", CH_HEADERS, session_id, rows)


def get_live_index(session_id: str) -> int:
    df = ws_read_df("live")
    if df.empty or "session_id" not in df.columns:
        return 0
    m = df[df["session_id"] == session_id]
    if m.empty:
        return 0
    try:
        return int(m.iloc[0]["current_chapter_index"])
    except Exception:
        return 0


def set_live_index(session_id: str, idx: int) -> None:
    ws = ws_get_or_create("live", LIVE_HEADERS)
    values = ws.get_all_values()
    if not values:
        ws.append_row(LIVE_HEADERS)
        values = [LIVE_HEADERS]
    header = values[0]
    idx_key = header.index("session_id")
    found = None
    for i, r in enumerate(values[1:], start=2):
        if len(r) > idx_key and r[idx_key] == session_id:
            found = i
            break
    row = {
        "session_id": session_id,
        "current_chapter_index": str(int(idx)),
        "updated_at_utc": utc_now().isoformat(),
    }
    out = [row.get(h, "") for h in header]
    if found:
        ws.update(f"A{found}", [out])
    else:
        ws.append_row(out)


def gemini_available() -> bool:
    return bool(st.secrets.get("gemini", {}).get("api_key", ""))


def gemini_generate(payload: Dict[str, str]) -> Dict[str, str]:
    if not gemini_available():
        return {
            "story_text": "Defina gemini.api_key nos Secrets do Streamlit Cloud.",
            "music_alternatives": "",
            "experience_actions": "",
            "notes": "",
        }
    import google.generativeai as genai
    genai.configure(api_key=st.secrets["gemini"]["api_key"])
    model = genai.GenerativeModel("gemini-1.5-pro")
    prompt = f"""
Você é curador de experiências gastronômicas com música ao vivo.

Gênero: {payload.get("genre","")}
Momento: {payload.get("moment_key","")}

Música:
Título: {payload.get("music_title","")}
Artista: {payload.get("music_artist","")}

Contexto do diretor:
{payload.get("link_context_input","")}

Entregue exatamente nestas 4 seções:

1) STORY_TEXT
4 a 6 linhas conectando momento, música e a casa.

2) MUSIC_ALTERNATIVES
3 a 6 sugestões no formato:
"Título" | Artista | por que funciona (máx 10 palavras)

3) EXPERIENCE_ACTIONS
4 a 8 bullets práticos (serviço e dinâmica)

4) NOTES
Notas extras.
""".strip()
    res = model.generate_content(prompt)
    text = res.text or ""
    def extract(section: str) -> str:
        m = re.search(rf"{section}\s*(.*?)(?=\n\s*\d\)|\Z)", text, flags=re.S | re.I)
        return (m.group(1).strip() if m else "").strip()
    return {
        "story_text": extract("STORY_TEXT"),
        "music_alternatives": extract("MUSIC_ALTERNATIVES"),
        "experience_actions": extract("EXPERIENCE_ACTIONS"),
        "notes": extract("NOTES"),
    }


def append_gemini_history(row: Dict[str, Any]) -> None:
    ws_append_row("gemini_history", HIST_HEADERS, [
        utc_now().isoformat(),
        row.get("session_id", ""),
        row.get("chapter_index", ""),
        row.get("moment_key", ""),
        row.get("music_title", ""),
        row.get("music_artist", ""),
        row.get("link_context_input", ""),
        row.get("ai_story_text", ""),
        row.get("ai_music_alternatives", ""),
        row.get("ai_experience_actions", ""),
        row.get("ai_notes", ""),
    ])


def render_client_view(session_id: str, show_controls: bool = False) -> None:
    sess = get_session_row(session_id)
    if not sess:
        st.error("Sessão não encontrada na planilha.")
        return

    ch = read_chapters(session_id)
    if ch.empty:
        st.warning("Sessão sem capítulos na planilha.")
        return

    cur_idx = get_live_index(session_id)
    cur = ch[ch["chapter_index"] == cur_idx]
    if cur.empty:
        cur = ch.iloc[[0]]
        cur_idx = int(cur.iloc[0]["chapter_index"])

    row = cur.iloc[0].to_dict()

    if show_controls:
        c1, c2, c3 = st.columns([1, 1, 3])
        with c1:
            if st.button("◀ Anterior", use_container_width=True):
                set_live_index(session_id, max(cur_idx - 1, 0))
                st.rerun()
        with c2:
            if st.button("Próximo ▶", use_container_width=True):
                max_idx = int(ch["chapter_index"].max())
                set_live_index(session_id, min(cur_idx + 1, max_idx))
                st.rerun()
        with c3:
            st.caption(f"Capítulo ao vivo: {cur_idx}")

    st.markdown('<div class="mobile-wrap">', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="show-card">
          <div class="show-h1">Agora tocando</div>
          <div style="font-weight:700; color: rgba(14,42,71,0.92); font-size:16px;">
            {row.get("music_title","") or "Ao vivo"}
          </div>
          <div class="show-muted">{row.get("music_artist","")}</div>
          <div class="show-muted">Momento: {row.get("moment_key","")}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="show-card">
          <div class="show-h1">Conexão</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    story = row.get("ai_story_text", "") or "A narrativa desta música aparece aqui."
    story_safe = story.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${}")
    components.html(
        f"""
        <div style="white-space:pre-wrap; font-size:15px; color: rgba(14,42,71,0.92); line-height:1.35;" id="t"></div>
        <script>
          const el = document.getElementById("t");
          const full = `{story_safe}`;
          let i = 0;
          function tick() {{
            if (i <= full.length) {{
              el.textContent = full.slice(0, i);
              i++;
              setTimeout(tick, 16);
            }}
          }}
          tick();
        </script>
        """,
        height=160,
    )
    st.markdown("</div>", unsafe_allow_html=True)


def safe_setlist_df(ch: pd.DataFrame) -> pd.DataFrame:
    if ch.empty:
        return ch
    show = ch.copy()
    if "planned_duration_sec" not in show.columns:
        show["planned_duration_sec"] = 0
    show["planned_duration_sec"] = pd.to_numeric(show["planned_duration_sec"], errors="coerce").fillna(0).astype(int)
    keep = [c for c in ["chapter_index", "moment_key", "chapter_type", "planned_duration_sec", "music_title", "music_artist"] if c in show.columns]
    show = show[keep].copy()
    show["planned_duration_sec"] = show["planned_duration_sec"].apply(fmt_mmss)
    return show.rename(columns={"planned_duration_sec": "duração"})


def page_login() -> None:
    render_header("Login único")
    st.write("Use o mesmo link para Cliente, Banda e Operação.")
    with st.sidebar:
        st.markdown("### Acesso interno")
        u = st.text_input("Login")
        p = st.text_input("Senha", type="password")
        if st.button("Entrar"):
            user = auth_login(u, p)
            if not user:
                st.warning("Login ou senha incorretos.")
            else:
                st.session_state["user"] = user
                st.success(f"Logado como {user['role']}")
                st.rerun()

    qp_sid = st.query_params.get("sid", "")
    df = list_sessions_sheet()
    if qp_sid:
        render_client_view(qp_sid)
        time.sleep(2)
        st.rerun()
    elif not df.empty:
        st.subheader("Cliente")
        sid = st.selectbox("Escolha a sessão", df["session_id"].tolist(), key="public_sid")
        render_client_view(sid)
        time.sleep(2)
        st.rerun()
    else:
        st.info("Sem sessões ainda. Faça login como admin para criar a primeira.")


def page_band() -> None:
    render_header("Banda (somente leitura)")
    df = list_sessions_sheet()
    if df.empty:
        st.info("Sem sessões.")
        return
    sid = st.selectbox("Sessão", df["session_id"].tolist(), key="band_sid")
    ch = read_chapters(sid)
    if ch.empty:
        st.warning("Sem capítulos.")
        return
    st.dataframe(safe_setlist_df(ch), use_container_width=True, hide_index=True)
    st.subheader("Cliente")
    render_client_view(sid, show_controls=False)


def page_admin() -> None:
    render_header("Operação (admin)")
    st.subheader("Criar sessão")
    with st.form("create_session"):
        sid = st.text_input("session_id", value=datetime.now().strftime("%Y%m%d-1900"))
        title = st.text_input("Título", value="Yvora Music Session")
        genre = st.text_input("Gênero", value="Jazz")
        total = st.number_input("Duração total (min)", min_value=45, max_value=240, value=120, step=5)
        ok = st.form_submit_button("Criar")
    if ok:
        upsert_session(sid, title, genre, int(total), "draft")
        base = []
        for i in range(10):
            base.append({
                "chapter_index": i,
                "moment_key": "aquecimento" if i < 3 else ("intervalo" if i == 5 else "intensidade"),
                "chapter_type": "break" if i == 5 else "music",
                "planned_duration_sec": 600 if i == 5 else 300,
                "music_title": "",
                "music_artist": "",
                "link_context_input": "",
                "ai_story_text": "",
                "ai_music_alternatives": "",
                "ai_experience_actions": "",
                "ai_notes": "",
            })
        save_chapters(sid, pd.DataFrame(base))
        set_live_index(sid, 0)
        st.success("Sessão criada e gravada no Google Sheets.")
        st.rerun()

    sess_df = list_sessions_sheet()
    if sess_df.empty:
        st.info("Sem sessões.")
        return

    sid2 = st.selectbox("Sessão", sess_df["session_id"].tolist(), key="adm_sid")
    sess = get_session_row(sid2) or {}
    genre = sess.get("genre", "Jazz")
    ch = read_chapters(sid2)
    if ch.empty:
        st.warning("Sem capítulos.")
        return

    st.subheader("Resumo")
    st.dataframe(safe_setlist_df(ch), use_container_width=True, hide_index=True)

    st.subheader("Editar capítulos")
    edit_cols = [c for c in [
        "chapter_index", "moment_key", "chapter_type", "planned_duration_sec",
        "music_title", "music_artist", "link_context_input"
    ] if c in ch.columns]
    edited = st.data_editor(ch[edit_cols], use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns([1, 1, 1])

    with c1:
        if st.button("Salvar alterações"):
            full = ch.copy()
            for col in edit_cols:
                full[col] = edited[col]
            save_chapters(sid2, full)
            st.success("Capítulos salvos.")
            st.rerun()

    with c2:
        if st.button("Gemini: gerar histórico"):
            full = ch.copy().reset_index(drop=True)
            n = 0
            for i in range(len(full)):
                if str(full.loc[i, "chapter_type"]).strip().lower() != "music":
                    continue
                payload = {
                    "genre": genre,
                    "moment_key": str(full.loc[i, "moment_key"]),
                    "music_title": str(full.loc[i, "music_title"]),
                    "music_artist": str(full.loc[i, "music_artist"]),
                    "link_context_input": str(full.loc[i, "link_context_input"]),
                }
                out = gemini_generate(payload)
                full.loc[i, "ai_story_text"] = out["story_text"]
                full.loc[i, "ai_music_alternatives"] = out["music_alternatives"]
                full.loc[i, "ai_experience_actions"] = out["experience_actions"]
                full.loc[i, "ai_notes"] = out["notes"]
                append_gemini_history({
                    "session_id": sid2,
                    "chapter_index": int(full.loc[i, "chapter_index"]),
                    "moment_key": payload["moment_key"],
                    "music_title": payload["music_title"],
                    "music_artist": payload["music_artist"],
                    "link_context_input": payload["link_context_input"],
                    "ai_story_text": out["story_text"],
                    "ai_music_alternatives": out["music_alternatives"],
                    "ai_experience_actions": out["experience_actions"],
                    "ai_notes": out["notes"],
                })
                n += 1
            save_chapters(sid2, full)
            st.success(f"IA gerou {n} capítulos e registrou histórico.")
            st.rerun()

    with c3:
        if st.button("Abrir modo cliente"):
            st.query_params["sid"] = sid2
            st.rerun()

    st.subheader("Controle ao vivo")
    render_client_view(sid2, show_controls=True)


def main():
    app_bootstrap()
    try:
        _ = get_sheet()
    except Exception as e:
        st.error(f"Erro de configuração: {e}")
        st.info("O app aceita [google].service_account_json ou [gcp_service_account] nos Secrets.")
        return

    user = st.session_state.get("user")
    qp_sid = st.query_params.get("sid", "")

    with st.sidebar:
        st.markdown("### Navegação")
        if user:
            st.markdown(
                f"<span class='yv-badge'>{user.get('username')}</span><span class='yv-badge'>{user.get('role')}</span>",
                unsafe_allow_html=True,
            )
            if st.button("Sair"):
                logout()

    if not user:
        page_login()
        return

    role = user.get("role", "")
    if role == "admin":
        page = st.sidebar.radio("Área", ["Operação", "Banda", "Cliente"], index=0)
        if page == "Operação":
            page_admin()
        elif page == "Banda":
            page_band()
        else:
            render_header("Cliente")
            df = list_sessions_sheet()
            if df.empty:
                st.info("Sem sessões.")
            else:
                sid = qp_sid or st.selectbox("Sessão", df["session_id"].tolist(), key="admin_public_sid")
                render_client_view(sid, show_controls=True)
    elif role == "band":
        page_band()
    else:
        render_header("Cliente")
        sid = qp_sid
        if not sid:
            df = list_sessions_sheet()
            if df.empty:
                st.info("Sem sessões.")
                return
            sid = st.selectbox("Sessão", df["session_id"].tolist(), key="staff_public_sid")
        render_client_view(sid)


if __name__ == "__main__":
    main()
