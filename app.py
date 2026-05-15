"""
app.py
NUS Development Office – Prospect Profile Generator
Sidebar layout with queue and downloads in the main area.
"""

import io
import os
import time
import zipfile
import streamlit as st
import anthropic

from profile_generator import research_prospect, get_filename
from docx_builder import build_profile_docx

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="NUS Prospect Profile Generator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .stButton > button { width: 100%; }
    .status-done    { color: #198754; font-weight: 600; }
    .status-error   { color: #dc3545; font-weight: 600; }
    .status-pending { color: #6c757d; }
    .status-running { color: #003D7C; font-weight: 600; }
    h1 { color: #003D7C; }
    .conf-note {
        font-size: 0.78rem;
        color: #6c757d;
        border-top: 1px solid #dee2e6;
        padding-top: 0.5rem;
        margin-top: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# Session state
# ──────────────────────────────────────────────
for key, default in [
    ("queue",       []),
    ("results",     {}),
    ("generating",  False),
    ("next_idx",    0),
    ("api_key",     ""),
    ("tavily_key",  ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# Pre-fill API key from secrets / env
if not st.session_state.api_key:
    try:
        st.session_state.api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        pass
    if not st.session_state.api_key:
        st.session_state.api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not st.session_state.tavily_key:
        try:
            st.session_state.tavily_key = st.secrets.get("TAVILY_API_KEY", "")
        except Exception:
            pass
    if not st.session_state.tavily_key:
        st.session_state.tavily_key = os.environ.get("TAVILY_API_KEY", "")

# ──────────────────────────────────────────────
# Header (always shown)
# ──────────────────────────────────────────────
st.title("🎓 NUS Prospect Profile Generator")
st.caption("Development Office — internal use only")
st.divider()

# ──────────────────────────────────────────────
# API key gate
# ──────────────────────────────────────────────
if not st.session_state.api_key:
    st.markdown("### 🔑 Enter your API keys to get started")
    st.markdown(
        "Your keys are used only for this session and never stored permanently."
    )
    with st.form("api_key_form"):
        key_input = st.text_input(
            "Anthropic API key *",
            type="password",
            placeholder="sk-ant-...",
            help="Required. Get yours at console.anthropic.com",
        )
        tavily_input = st.text_input(
            "Tavily API key (optional — enables live web search)",
            type="password",
            placeholder="tvly-...",
            help="Optional but recommended. Free at app.tavily.com. Without this, profiles use Claude training knowledge only.",
        )
        submitted = st.form_submit_button("Continue →", type="primary")
        if submitted:
            key_input    = key_input.strip()
            tavily_input = tavily_input.strip()
            if not key_input.startswith("sk-"):
                st.error("Anthropic key should start with `sk-`. Please check and try again.")
            else:
                try:
                    test_client = anthropic.Anthropic(api_key=key_input)
                    test_client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=10,
                        messages=[{"role": "user", "content": "Hi"}]
                    )
                    st.session_state.api_key    = key_input
                    st.session_state.tavily_key = tavily_input
                    st.rerun()
                except anthropic.AuthenticationError:
                    st.error("Invalid Anthropic API key — authentication failed. Please check and try again.")
                except Exception:
                    st.session_state.api_key    = key_input
                    st.session_state.tavily_key = tavily_input
                    st.rerun()
    st.stop()

# ──────────────────────────────────────────────
# Build client from validated session key
# ──────────────────────────────────────────────
client = anthropic.Anthropic(api_key=st.session_state.api_key)

# ──────────────────────────────────────────────
# Sidebar – Add prospect
# ──────────────────────────────────────────────
with st.sidebar:
    st.header("➕ Add Prospect")
    with st.form("add_form", clear_on_submit=True):
        name_input = st.text_input(
            "Prospect name *",
            placeholder="e.g. Wee Cho Yaw  /  Temasek Holdings",
            help="Enter a person's full name or a company/organisation name."
        )
        photo_file = st.file_uploader(
            "Photo / logo (optional)",
            type=["jpg", "jpeg", "png", "webp"],
            help="Upload a photo for individuals or a logo for companies."
        )
        submitted = st.form_submit_button("Add to queue", type="primary")

        if submitted:
            name_clean = name_input.strip()
            if not name_clean:
                st.error("Please enter a prospect name.")
            elif any(p["name"].lower() == name_clean.lower()
                     for p in st.session_state.queue):
                st.warning(f"**{name_clean}** is already in the queue.")
            else:
                photo_bytes = None
                if photo_file is not None:
                    photo_bytes = photo_file.read()
                st.session_state.queue.append({
                    "name":        name_clean,
                    "photo_bytes": photo_bytes,
                    "idx":         st.session_state.next_idx,
                })
                st.session_state.next_idx += 1
                st.success(f"Added: **{name_clean}**")

    # ── Queue summary ──
    st.divider()
    n_queue = len(st.session_state.queue)
    n_done  = sum(1 for r in st.session_state.results.values()
                  if r.get("status") == "done")
    st.metric("In queue",   n_queue)
    st.metric("Generated",  n_done)

    if st.session_state.queue:
        if st.button("🗑️ Clear entire queue", type="secondary"):
            st.session_state.queue   = []
            st.session_state.results = {}
            st.rerun()

    st.divider()
    if st.session_state.tavily_key:
        st.success("🌐 Live web search: ON", icon="✅")
    else:
        st.warning("🌐 Live web search: OFF (no Tavily key)", icon="⚠️")

    st.markdown(
        '<p class="conf-note">NUS CONFIDENTIAL · for internal use only · NOT for external circulation</p>',
        unsafe_allow_html=True
    )

    st.divider()
    if st.button("🔑 Change API keys", type="secondary"):
        st.session_state.api_key    = ""
        st.session_state.tavily_key = ""
        st.session_state.queue      = []
        st.session_state.results    = {}
        st.rerun()

# ──────────────────────────────────────────────
# Main area
# ──────────────────────────────────────────────
left, right = st.columns([3, 2], gap="large")

# ── Left: Prospect queue ──
with left:
    st.subheader("📋 Prospect Queue")

    if not st.session_state.queue:
        st.info("No prospects in the queue. Add names using the sidebar.", icon="👈")
    else:
        est_minutes = len(st.session_state.queue) * 2
        st.caption(
            f"{n_queue} prospect(s) queued · "
            f"estimated time: ~{est_minutes}–{est_minutes * 2} minutes"
        )

        for prospect in st.session_state.queue:
            idx    = prospect["idx"]
            name   = prospect["name"]
            result = st.session_state.results.get(idx, {})
            status = result.get("status", "pending")

            c1, c2, c3 = st.columns([6, 3, 2])

            with c1:
                icon = {"done": "✅", "error": "❌", "generating": "⏳"}.get(status, "🔵")
                st.markdown(f"**{icon} {name}**")
                if status == "error":
                    st.caption(f"Error: {result.get('error', 'Unknown error')}")

            with c2:
                status_label = {
                    "pending":    "Pending",
                    "generating": "Generating…",
                    "done":       "Done",
                    "error":      "Failed",
                }.get(status, "Pending")
                css = {
                    "done":       "status-done",
                    "error":      "status-error",
                    "generating": "status-running",
                }.get(status, "status-pending")
                st.markdown(
                    f'<span class="{css}">{status_label}</span>',
                    unsafe_allow_html=True
                )

            with c3:
                if st.button("Remove", key=f"rm_{idx}",
                             disabled=st.session_state.generating):
                    st.session_state.queue   = [p for p in st.session_state.queue
                                                 if p["idx"] != idx]
                    st.session_state.results.pop(idx, None)
                    st.rerun()

        st.divider()

        can_generate = (
            not st.session_state.generating
            and any(
                st.session_state.results.get(p["idx"], {}).get("status") != "done"
                for p in st.session_state.queue
            )
        )

        if st.button(
            f"🚀 Generate All Profiles ({n_queue})",
            type="primary",
            disabled=not can_generate,
            use_container_width=True,
        ):
            st.session_state.generating = True
            st.rerun()

# ── Right: Downloads ──
with right:
    st.subheader("⬇️ Download Profiles")

    completed = [
        (p, st.session_state.results[p["idx"]])
        for p in st.session_state.queue
        if st.session_state.results.get(p["idx"], {}).get("status") == "done"
    ]

    if not completed:
        st.info("Generated profiles will appear here.", icon="📄")
    else:
        if len(completed) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for prospect, result in completed:
                    zf.writestr(result["filename"], result["docx_bytes"])
            zip_buf.seek(0)
            st.download_button(
                label=f"📦 Download all {len(completed)} profiles (.zip)",
                data=zip_buf.getvalue(),
                file_name="NUS_Prospect_Profiles.zip",
                mime="application/zip",
                use_container_width=True,
            )
            st.divider()

        for prospect, result in completed:
            st.download_button(
                label=f"📄 {prospect['name']}",
                data=result["docx_bytes"],
                file_name=result["filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_{prospect['idx']}",
                use_container_width=True,
            )

# ──────────────────────────────────────────────
# Generation loop
# ──────────────────────────────────────────────
if st.session_state.generating:
    pending = [
        p for p in st.session_state.queue
        if st.session_state.results.get(p["idx"], {}).get("status") != "done"
    ]

    if not pending:
        st.session_state.generating = False
        st.rerun()

    total      = len(pending)
    prog       = st.progress(0, text="Starting generation…")
    status_box = st.empty()

    for i, prospect in enumerate(pending):
        idx  = prospect["idx"]
        name = prospect["name"]

        prog.progress(i / total, text=f"Generating profile {i+1} of {total}: **{name}**")
        st.session_state.results[idx] = {"status": "generating"}

        log_lines = []

        def progress_cb(msg: str):
            log_lines.append(msg)
            status_box.info("\n".join(f"• {l}" for l in log_lines[-8:]))

        try:
            progress_cb(f"[{i+1}/{total}] Researching {name}…")
            data = research_prospect(name, client, progress_callback=progress_cb, tavily_key=st.session_state.tavily_key)

            if data is None:
                raise ValueError(
                    "Claude did not return valid JSON after 3 attempts. "
                    "Try a more specific name (e.g. include company or country)."
                )

            progress_cb("Building .docx file…")
            docx_bytes = build_profile_docx(data, prospect.get("photo_bytes"))
            filename   = get_filename(data)

            st.session_state.results[idx] = {
                "status":     "done",
                "docx_bytes": docx_bytes,
                "filename":   filename,
            }
            progress_cb(f"✅ Done: {filename}")

        except Exception as e:
            st.session_state.results[idx] = {
                "status": "error",
                "error":  str(e),
            }
            progress_cb(f"❌ Error: {e}")

        if i < total - 1:
            time.sleep(1)

    prog.progress(1.0, text="All profiles generated.")
    time.sleep(1)
    st.session_state.generating = False
    st.rerun()
