"""
ui_theme.py - modern, AI-product look for TMC Knowledge Lens.
"""
from __future__ import annotations

import html
from datetime import datetime

LOGO = "https://tmcltd.com/wp-content/uploads/2025/10/TallyMarks-Consulting-TMC-logo.webp"

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500&display=swap');

:root{
  --ink:#0F172A; --ink-2:#334155; --muted:#64748B; --line:#E2E8F0; --card:#FFFFFF; --bg:#F6F8FB;
  --brand:#0F6B45; --brand-2:#10B981; --brand-3:#14B8A6; --ai-1:#10B981; --ai-2:#3B82F6; --ai-3:#8B5CF6;
  --ring:0 0 0 3px rgba(16,185,129,.22);
  --shadow-sm:0 1px 2px rgba(15,23,42,.05),0 1px 3px rgba(15,23,42,.06);
  --shadow-md:0 8px 24px rgba(15,23,42,.08);
  --shadow-lg:0 18px 44px rgba(15,23,42,.14);
  --r:16px;
}
html,body,.stApp,[class*="css"]{font-family:'Inter','Segoe UI',system-ui,sans-serif!important}
.stApp{
  background:
    radial-gradient(1100px 480px at 8% -8%, rgba(16,185,129,.13) 0%, transparent 60%),
    radial-gradient(900px 420px at 100% 0%, rgba(59,130,246,.11) 0%, transparent 58%),
    var(--bg);
  color:var(--ink);
}
#MainMenu,footer{visibility:hidden}
header[data-testid="stHeader"]{background:transparent}

/* --- FIX: Added 130px padding-bottom so content clears the fixed footer --- */
.block-container{padding-top:1.4rem!important;padding-bottom:130px!important;max-width:1600px}

h1,h2,h3,h4{letter-spacing:-.015em;color:var(--ink)}
::-webkit-scrollbar{width:8px;height:8px}::-webkit-scrollbar-thumb{background:#CBD5E1;border-radius:8px}::-webkit-scrollbar-thumb:hover{background:#94A3B8}

/* ---------- entrance animation ---------- */
@keyframes klfade{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
@keyframes klflow{0%{background-position:0 0,0% 50%}100%{background-position:0 0,300% 50%}}
@keyframes klpulse{0%{box-shadow:0 0 0 0 rgba(52,211,153,.7)}70%{box-shadow:0 0 0 8px rgba(52,211,153,0)}100%{box-shadow:0 0 0 0 rgba(52,211,153,0)}}
div[data-testid="stVerticalBlock"]>div{animation:klfade .35s ease both}

/* ---------- hero / top bar ---------- */
.topbar{
  position:relative;overflow:hidden;display:flex;flex-wrap:wrap;align-items:center;gap:22px;padding:26px 34px;border-radius:24px;margin-bottom:18px;
  background:
    radial-gradient(620px 240px at 88% -25%, rgba(45,212,191,.50), transparent 62%),
    radial-gradient(520px 280px at -2% 125%, rgba(59,130,246,.34), transparent 60%),
    linear-gradient(118deg,#04261B 0%,#0A5237 52%,#0E7A57 100%);
  box-shadow:0 18px 44px rgba(6,60,41,.32),inset 0 1px 0 rgba(255,255,255,.12);
}
.topbar::before{content:"";position:absolute;inset:0;pointer-events:none;
  background-image:linear-gradient(rgba(255,255,255,.055) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.055) 1px,transparent 1px);
  background-size:34px 34px;-webkit-mask-image:linear-gradient(180deg,#000 0%,transparent 90%);mask-image:linear-gradient(180deg,#000 0%,transparent 90%)}
.topbar>*{position:relative;z-index:1}
.topbar .logo{background:rgba(255,255,255,.97);border-radius:18px;padding:12px 18px;display:flex;align-items:center;flex-shrink:0;
  box-shadow:0 8px 22px rgba(0,0,0,.22),0 0 0 1px rgba(255,255,255,.4)}
.topbar .logo img{height:64px;width:auto;display:block}
.topbar .titles{flex:1;min-width:260px}
.ai-badge{display:inline-flex;align-items:center;gap:8px;font-size:.72rem;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:#D1FAE5;
  background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.22);border-radius:999px;padding:5px 12px;backdrop-filter:blur(6px)}
.ai-badge .dot{width:8px;height:8px;border-radius:50%;background:#34D399;animation:klpulse 2s infinite}
.topbar h1{color:#fff;margin:10px 0 0;font-size:2.35rem;font-weight:800;line-height:1.05;letter-spacing:-.03em}
.topbar h1 .grad{background:linear-gradient(90deg,#A7F3D0 0%,#67E8F9 55%,#C4B5FD 100%);-webkit-background-clip:text;background-clip:text;color:transparent}
.topbar p{color:#CFEBDD;margin:10px 0 0;font-size:1rem;max-width:760px;line-height:1.5}
.topbar .pills{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}
.topbar .pill{color:#fff;font-size:.82rem;font-weight:600;white-space:nowrap;padding:8px 15px;border-radius:999px;
  background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.26);backdrop-filter:blur(8px);transition:.2s}
.topbar .pill:hover{background:rgba(255,255,255,.2);transform:translateY(-1px)}

/* ---------- buttons ---------- */
.stButton>button,.stDownloadButton>button,.stFormSubmitButton>button,.stLinkButton>a{
  border-radius:12px;font-weight:600;border:1px solid var(--line);background:var(--card);color:var(--ink);
  box-shadow:var(--shadow-sm);transition:transform .15s,box-shadow .15s,border-color .15s,background .15s}
.stButton>button:hover,.stDownloadButton>button:hover,.stFormSubmitButton>button:hover{
  border-color:var(--brand-2);color:var(--brand);transform:translateY(-1px);box-shadow:var(--shadow-md)}
.stButton>button:active{transform:translateY(0)}
.stButton>button[kind="primary"],.stFormSubmitButton>button[kind="primary"],.stDownloadButton>button[kind="primary"]{
  background:linear-gradient(135deg,#0B7A50 0%,#10B981 100%);border:none;color:#fff;
  box-shadow:0 6px 18px rgba(16,185,129,.35),inset 0 1px 0 rgba(255,255,255,.25)}
.stButton>button[kind="primary"]:hover,.stFormSubmitButton>button[kind="primary"]:hover{
  background:linear-gradient(135deg,#0D8A5A 0%,#18C795 100%);color:#fff;box-shadow:0 10px 26px rgba(16,185,129,.45)}
.stButton>button:disabled{opacity:.5;box-shadow:none;transform:none}

/* ---------- inputs ---------- */
div[data-baseweb="input"],div[data-baseweb="base-input"],div[data-baseweb="select"]>div,div[data-baseweb="textarea"]{
  border-radius:12px!important;border-color:var(--line)!important;background:#fff!important;transition:box-shadow .15s,border-color .15s}
div[data-baseweb="input"]:focus-within,div[data-baseweb="select"]>div:focus-within,div[data-baseweb="textarea"]:focus-within{
  border-color:var(--brand-2)!important;box-shadow:var(--ring)}
.stTextInput label,.stTextArea label,.stSelectbox label,.stMultiSelect label{font-weight:600;color:var(--ink-2);font-size:.84rem}
span[data-baseweb="tag"]{background:#E7F7F0!important;color:#0A5237!important;border-radius:8px!important;font-weight:600}

/* ---------- containers, expanders, tabs ---------- */
div[data-testid="stVerticalBlockBorderWrapper"]{border-radius:20px;border-color:var(--line)!important;background:rgba(255,255,255,.82);
  backdrop-filter:blur(10px);box-shadow:var(--shadow-md)}
div[data-testid="stExpander"]{border:1px solid var(--line);border-radius:14px;background:#fff;box-shadow:var(--shadow-sm);overflow:hidden}
div[data-testid="stExpander"] summary{font-weight:600}
div[data-testid="stExpander"] summary:hover{color:var(--brand)}
div[data-baseweb="tab-list"]{gap:6px;background:#EEF2F7;padding:5px;border-radius:14px;width:fit-content;max-width:100%;overflow-x:auto}
button[data-baseweb="tab"]{border-radius:10px!important;padding:8px 16px!important;font-weight:600;color:var(--muted)}
button[data-baseweb="tab"][aria-selected="true"]{background:#fff!important;color:var(--brand)!important;box-shadow:var(--shadow-sm)}
div[data-baseweb="tab-highlight"],div[data-baseweb="tab-border"]{display:none}
div[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:14px;overflow:hidden;box-shadow:var(--shadow-sm)}
div[data-testid="stAlert"]{border-radius:14px;border:1px solid var(--line)}
div[role="dialog"]{border-radius:22px!important}
div[data-testid="stProgress"]>div>div>div{background:linear-gradient(90deg,var(--ai-1),var(--ai-2))}

/* ---------- sidebar ---------- */
section[data-testid="stSidebar"]{background:radial-gradient(420px 260px at 0% 0%,rgba(16,185,129,.16),transparent 60%),linear-gradient(180deg,#0B1220 0%,#101A2C 100%)!important;
  border-right:1px solid #1E293B}
section[data-testid="stSidebar"] :is(h1,h2,h3,p,label,li,summary,small,span){color:#E2E8F0!important}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"]{color:#94A3B8!important}
section[data-testid="stSidebar"] hr{border-color:#243044}
section[data-testid="stSidebar"] :is(input,textarea){background:#162235!important;color:#fff!important}
section[data-testid="stSidebar"] :is(div[data-baseweb="input"],div[data-baseweb="base-input"]){background:#162235!important;border-color:#2B3A52!important}
section[data-testid="stSidebar"] div[data-testid="stExpander"]{background:rgba(255,255,255,.04);border:1px solid #243044;box-shadow:none}
section[data-testid="stSidebar"] .stButton>button{background:rgba(255,255,255,.06);color:#fff;border:1px solid #2B3A52;box-shadow:none}
section[data-testid="stSidebar"] .stButton>button:hover{background:rgba(16,185,129,.18);border-color:#34D399;color:#fff}
section[data-testid="stSidebar"] code{background:#0B1424;color:#6EE7B7}
button[kind="header"],button[data-testid="baseButton-header"]{color:#fff!important}

/* ---------- typography helpers used by the app ---------- */
.panel-title{display:flex;align-items:center;gap:8px;font-weight:700;color:var(--ink);font-size:1.02rem;margin:6px 0 8px}
.panel-title::before{content:"";width:4px;height:16px;border-radius:4px;background:linear-gradient(180deg,var(--ai-1),var(--ai-2))}
.small{font-size:.78rem;color:var(--muted)}
.chart-title{font-weight:700;color:var(--ink);margin:6px 0 2px}

.chip{display:inline-block;background:#F1F5F9;color:#334155;border:1px solid #E2E8F0;border-radius:999px;padding:2px 10px;font-size:.72rem;font-weight:500;margin:3px 4px 2px 0}
.chip.ok{background:#ECFDF5;color:#047857;border-color:#A7F3D0}
.chip.old{background:#FFFBEB;color:#B45309;border-color:#FDE68A}
.chip.p{background:#EEF2FF;color:#4338CA;border-color:#C7D2FE;font-family:'JetBrains Mono',monospace;font-size:.68rem}

.doc-title{font-weight:700;color:var(--ink);font-size:.95rem;line-height:1.3;margin-bottom:3px}
.relbar{height:6px;background:#E8EEF5;border-radius:6px;margin:8px 0 3px;overflow:hidden}
.relbar div{height:6px;border-radius:6px;background:linear-gradient(90deg,var(--ai-1),var(--ai-2));box-shadow:0 0 10px rgba(59,130,246,.35)}
.snip{font-size:.8rem;color:#475569;background:#F8FAFC;border:1px solid #EDF2F7;border-left:3px solid var(--brand-2);padding:7px 10px;margin-top:5px;border-radius:8px;line-height:1.45}

/* ---------- cards ---------- */
.hero{position:relative;background:#fff;border:1px solid var(--line);border-radius:20px;padding:22px 28px;margin-bottom:16px;box-shadow:var(--shadow-md);overflow:hidden}
.hero::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,var(--ai-1),var(--ai-2),var(--ai-3))}
.hero::after{content:"";position:absolute;right:-60px;top:-60px;width:220px;height:220px;border-radius:50%;
  background:radial-gradient(circle,rgba(16,185,129,.16),transparent 70%)}
.hero h2{margin:0 0 6px;color:var(--ink);font-size:1.6rem;font-weight:800;letter-spacing:-.02em}
.hero p{color:var(--ink-2);font-size:1.02rem;margin:0;position:relative;z-index:1}
.pcard{background:#fff;border:1px solid var(--line);border-radius:18px;padding:20px 22px;min-height:200px;box-shadow:var(--shadow-sm);position:relative;overflow:hidden;
  transition:transform .2s,box-shadow .2s,border-color .2s}
.pcard::before{content:"";position:absolute;left:0;right:0;top:0;height:3px;background:linear-gradient(90deg,var(--ai-1),var(--ai-2),var(--ai-3))}
.pcard:hover{transform:translateY(-3px);box-shadow:var(--shadow-lg);border-color:#BBF7D0}
.pcard h4{margin:4px 0 10px;color:var(--ink);font-size:1.08rem;font-weight:700}
.pcard p,.pcard li{color:var(--ink-2);font-size:.93rem;line-height:1.55}
.pcard ul{padding-left:18px;margin:4px 0 0}
.stepno{display:inline-flex;width:36px;height:36px;border-radius:12px;background:linear-gradient(135deg,var(--ai-1),var(--ai-2));color:#fff;align-items:center;justify-content:center;font-weight:800;margin-bottom:8px}
.person{background:#fff;border:1px solid var(--line);border-radius:18px;padding:20px 24px;box-shadow:var(--shadow-sm)}
.person a{color:var(--brand);font-weight:600;text-decoration:none}.person a:hover{text-decoration:underline}
.person .nm{font-size:1.2rem;font-weight:800;color:var(--ink)}
.glance{background:linear-gradient(135deg,#FFFFFF 0%,#F0FDF9 100%);border:1px solid #CFF1E3;border-radius:16px;padding:16px 20px;margin:10px 0;color:var(--ink-2);box-shadow:var(--shadow-sm);line-height:1.55}
.glance h4{margin:0 0 6px;color:var(--brand);font-weight:700}
.rolecard{background:#fff;border:1px solid var(--line);border-radius:16px;padding:16px;min-height:120px;box-shadow:var(--shadow-sm);transition:.2s}
.rolecard:hover{box-shadow:var(--shadow-md);transform:translateY(-2px)}
.rolecard .ic{font-size:1.5rem}.rolecard .s{font-size:.82rem;color:var(--muted);margin-top:4px;line-height:1.45}

.cmp{width:100%;border-collapse:separate;border-spacing:0;background:#fff;border:1px solid var(--line);border-radius:16px;overflow:hidden;font-size:.9rem;box-shadow:var(--shadow-sm)}
.cmp th{background:linear-gradient(135deg,#0A5237,#0E7A57);color:#fff;text-align:left;padding:11px 14px;font-weight:600}
.cmp td{padding:11px 14px;border-top:1px solid #EEF2F7;vertical-align:top;color:var(--ink-2)}
.cmp tr:hover td{background:#F8FFFC}.cmp td:first-child{font-weight:600;color:var(--ink)}

/* ---------- KPIs / insight ---------- */
.kpi-row{display:flex;gap:16px;margin-bottom:24px;align-items:stretch;flex-wrap:wrap}
.kpi-modern{position:relative;background:#fff;border:1px solid var(--line);border-radius:18px;padding:20px;box-shadow:var(--shadow-sm);flex:1;min-width:160px;overflow:hidden;transition:transform .2s,box-shadow .2s}
.kpi-modern::after{content:"";position:absolute;right:-30px;bottom:-30px;width:110px;height:110px;border-radius:50%;background:radial-gradient(circle,rgba(16,185,129,.12),transparent 70%)}
.kpi-modern:hover{transform:translateY(-3px);box-shadow:var(--shadow-lg)}
.kpi-icon{width:42px;height:42px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:1.25rem}
.kpi-value{font-size:2.05rem;font-weight:800;color:var(--ink);line-height:1;margin-top:16px;letter-spacing:-.03em}
.kpi-label{font-size:.88rem;font-weight:600;color:var(--ink-2);margin-top:5px}
.kpi-sub{font-size:.74rem;color:var(--muted);margin-top:6px}
.insight-banner{position:relative;overflow:hidden;background:linear-gradient(135deg,#ECFDF5 0%,#EFF6FF 60%,#F5F3FF 100%);border:1px solid #BBF7D0;border-radius:18px;padding:18px 24px;margin-bottom:24px;box-shadow:var(--shadow-sm)}
.insight-banner::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,var(--ai-1),var(--ai-2),var(--ai-3))}
.insight-banner h4{margin:0 0 4px;color:#065F46;font-weight:800}.insight-banner p{margin:0;color:#14532D;line-height:1.55}

/* ---------- chat ---------- */
div[data-testid="stChatMessage"]{border-radius:16px;padding:12px 16px;margin-bottom:8px;border:1px solid var(--line);background:#fff;box-shadow:var(--shadow-sm)}
div[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]){background:linear-gradient(135deg,#F0FDF9 0%,#F5FAFF 100%);border-color:#CFF1E3}
div[data-testid="stChatMessageAvatarAssistant"]{background:linear-gradient(135deg,var(--ai-1),var(--ai-2))!important;color:#fff}
div[data-testid="stChatMessageAvatarUser"]{background:linear-gradient(135deg,#334155,#0F172A)!important;color:#fff}
div[data-testid="stChatInput"]{background:transparent!important}
div[data-testid="stChatInput"]>div{
  background:linear-gradient(#fff,#fff) padding-box,linear-gradient(120deg,#10B981,#3B82F6,#8B5CF6,#10B981) border-box!important;
  background-size:auto,300% 300%;border:1.5px solid transparent!important;border-radius:999px!important;min-height:56px;
  box-shadow:0 8px 28px rgba(59,130,246,.14),0 2px 8px rgba(16,185,129,.12);animation:klflow 9s linear infinite}
div[data-testid="stChatInput"]>div:focus-within{box-shadow:0 10px 34px rgba(16,185,129,.28),0 0 0 4px rgba(16,185,129,.12)}
div[data-testid="stChatInput"] textarea{font-size:1rem;color:var(--ink)}
div[data-testid="stChatInput"] button{border-radius:50%!important;background:linear-gradient(135deg,#0B7A50,#10B981)!important;color:#fff!important;border:none;box-shadow:0 4px 12px rgba(16,185,129,.4)}
div[data-testid="stChatInput"] button:disabled{background:#CBD5E1!important;box-shadow:none}

/* ---------- metadata assistant ---------- */
.diff-old{background:#FEF2F2;color:#B91C1C;text-decoration:line-through;border-radius:6px;padding:1px 7px;font-size:.78rem}
.diff-new{background:#ECFDF5;color:#047857;border-radius:6px;padding:1px 7px;font-size:.78rem;font-weight:700}
.ma-doc{background:#fff;border:1px solid var(--line);border-left:4px solid var(--brand-2);border-radius:14px;padding:11px 15px;margin:7px 0;box-shadow:var(--shadow-sm)}
.ma-doc .row{margin-top:6px;font-size:.84rem;color:var(--ink-2)}
.ma-badge{display:inline-block;border-radius:999px;padding:3px 13px;font-size:.76rem;font-weight:700;margin-bottom:5px}
.ma-badge.pending{background:#FEF3C7;color:#92400E}.ma-badge.applied{background:#D1FAE5;color:#065F46}
.ma-badge.undone,.ma-badge.cancelled{background:#E2E8F0;color:#334155}.ma-badge.error{background:#FEE2E2;color:#991B1B}

/* ---------- footer ---------- */
.kl-footer{position:fixed;left:0;right:0;bottom:0;z-index:99;display:flex;justify-content:center;align-items:center;gap:10px;padding:8px 16px;font-size:.78rem;color:#CBD5E1;
  background:rgba(11,18,32,.92);backdrop-filter:blur(12px);border-top:1px solid rgba(52,211,153,.4)}
.kl-footer b{color:#fff}
.kl-footer .live{width:7px;height:7px;border-radius:50%;background:#34D399;animation:klpulse 2s infinite}

@media (max-width:900px){.topbar{padding:20px}.topbar h1{font-size:1.8rem}.topbar .pills{justify-content:flex-start}}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style>
"""


def _e(s) -> str:
    return html.escape(str(s))


def topbar(subtitle: str, pills: list[str] | None = None) -> str:
    """Hero banner. `pills` are plain strings (escaped here)."""
    p = "".join(f'<span class="pill">{_e(x)}</span>' for x in (pills or []))
    return (f'<div class="topbar"><div class="logo"><img src="{LOGO}" alt="TMC"></div>'
            f'<div class="titles"><span class="ai-badge"><i class="dot"></i>AI-powered knowledge workspace</span>'
            f'<h1>✦ Knowledge <span class="grad">Lens</span></h1><p>{_e(subtitle)}</p></div>'
            + (f'<div class="pills">{p}</div>' if p else "") + "</div>")


def footer_html(version: str) -> str:
    return (f'<div class="kl-footer"><span class="live"></span><b>TMC Knowledge Lens</b> v{_e(version)} • '
            f'Built by TMC AI Internal Digital Department © {datetime.now().year}</div>')
