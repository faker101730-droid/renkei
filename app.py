# -*- coding: utf-8 -*-
"""
RENKEI - 地域連携室 予約業務実績ダッシュボード

正式データ: Supabase public.renkei_monthly_metrics（Read RPC経由）
更新元データ: 画面からアップロードする最新Excel
必須列: 年度, 月, 月番号, 稼働日数, 予約件数
仕様: 予約件数が空欄の未来月・未入力月は 0 件扱いせず、計算対象外にする。
更新: Excel選択 → 明示的アップロード → Preview → 人間承認 → 固定Write RPC。ファイル選択だけでは正式値を変更しない。
"""

from __future__ import annotations

import io
import re
import hashlib
from typing import Iterable, Optional

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# =========================================================
# 基本設定
# =========================================================
st.set_page_config(
    page_title="RENKEI｜地域連携室 予約業務実績",
    page_icon="🔗",
    layout="wide",
)

# -----------------------------
def render_broly_page() -> None:
    st.markdown(
        r"""
<style>
  /* BROLY page-only background: synchronized with BROLY COMMAND CENTER */
  [data-testid="stAppViewContainer"],
  [data-testid="stAppViewContainer"] > .main,
  [data-testid="stAppViewContainer"] > .main > div,
  section.main,
  section.main > div.block-container{
    background-color:#050A12 !important;
  }
  [data-testid="stAppViewContainer"],
  [data-testid="stAppViewContainer"] > .main{
    background:
      radial-gradient(950px 640px at 52% 23%, rgba(11,124,255,.18), transparent 67%),
      radial-gradient(760px 520px at 94% 6%, rgba(128,93,255,.10), transparent 62%),
      radial-gradient(650px 520px at 0% 35%, rgba(57,215,255,.08), transparent 65%),
      linear-gradient(180deg,#050A12 0%,#07101C 52%,#040810 100%) !important;
  }
  [data-testid="stAppViewContainer"] > .main > div,
  section.main,
  section.main > div.block-container{
    background:transparent !important;
  }
  [data-testid="stAppViewContainer"]::after{
    content:"";
    position:fixed;
    inset:0;
    pointer-events:none;
    z-index:0;
    opacity:.18;
    background-image:
      linear-gradient(rgba(57,215,255,.06) 1px,transparent 1px),
      linear-gradient(90deg,rgba(57,215,255,.045) 1px,transparent 1px);
    background-size:34px 34px;
    -webkit-mask-image:linear-gradient(to bottom,rgba(0,0,0,.8),transparent 92%);
    mask-image:linear-gradient(to bottom,rgba(0,0,0,.8),transparent 92%);
  }
  section.main > div.block-container{
    position:relative;
    z-index:1;
  }

  .broly-page-shell{
    --bc:#39D7FF;
    --bb:#0B7CFF;
    --bv:#805DFF;
    position:relative;
    width:100%;
    min-height:735px;
    box-sizing:border-box;
    overflow:hidden;
    border-radius:0;
    border:none;
    background:transparent;
    box-shadow:none;
  }
  .broly-page-shell:before{
    content:"";
    position:absolute;
    inset:0;
    pointer-events:none;
    opacity:.27;
    background:
      linear-gradient(rgba(57,215,255,.045) 1px,transparent 1px),
      linear-gradient(90deg,rgba(57,215,255,.04) 1px,transparent 1px);
    background-size:34px 34px;
    -webkit-mask-image:radial-gradient(circle at 50% 38%,#000 0 26%,rgba(0,0,0,.55) 52%,transparent 82%);
    mask-image:radial-gradient(circle at 50% 38%,#000 0 26%,rgba(0,0,0,.55) 52%,transparent 82%);
  }
  .broly-page-head{
    position:relative;
    z-index:20;
    display:flex;
    align-items:flex-start;
    justify-content:space-between;
    gap:24px;
    padding:28px 30px 12px;
  }
  .broly-page-brand{
    min-width:0;
  }
  .broly-page-eyebrow{
    color:rgba(57,215,255,.54);
    font-size:9px;
    font-weight:900;
    letter-spacing:2.3px;
    line-height:1;
    white-space:nowrap;
  }
  .broly-page-title{
    margin-top:8px;
    color:rgba(244,252,255,.98);
    font-size:25px;
    font-weight:900;
    letter-spacing:2.0px;
    line-height:1.05;
    text-shadow:0 0 13px rgba(57,215,255,.12);
    white-space:nowrap;
  }
  .broly-page-sub{
    margin-top:7px;
    color:rgba(173,207,220,.58);
    font-size:9px;
    font-weight:750;
    letter-spacing:1.15px;
    line-height:1.4;
  }
  .broly-page-node{
    text-align:right;
    padding-top:2px;
    white-space:nowrap;
  }
  .broly-page-node-main{
    color:rgba(57,215,255,.72);
    font-size:9px;
    font-weight:900;
    letter-spacing:1.65px;
  }
  .broly-page-node-sub{
    margin-top:6px;
    color:rgba(164,198,212,.44);
    font-size:7px;
    font-weight:800;
    letter-spacing:1.1px;
  }

  .broly-page-core-stage{
    position:relative;
    z-index:10;
    width:100%;
    display:flex;
    align-items:center;
    justify-content:center;
    min-height:395px;
    padding-top:46px;
  }
  .broly-page-core-wrap{
    position:relative;
    width:320px;
    height:320px;
    display:flex;
    align-items:center;
    justify-content:center;
    isolation:isolate;
  }
  .broly-page-core-wrap:before{
    content:"";
    position:absolute;
    inset:12px;
    border-radius:50%;
    background:radial-gradient(circle,rgba(57,215,255,.09),transparent 62%);
    box-shadow:
      0 0 80px rgba(11,124,255,.13),
      0 0 130px rgba(57,215,255,.05);
  }

  .broly-page-halo{
    position:absolute;
    inset:13px;
    border-radius:50%;
    z-index:0;
    background:
      conic-gradient(from 215deg,
        transparent 0 7%,rgba(57,215,255,.60) 8% 11%,transparent 12% 29%,
        rgba(128,93,255,.38) 30% 34%,transparent 35% 60%,
        rgba(57,215,255,.46) 61% 65%,transparent 66% 86%,
        rgba(215,250,255,.72) 87% 89%,transparent 90% 100%);
    -webkit-mask:radial-gradient(farthest-side,transparent calc(100% - 3px),#000 calc(100% - 2px));
    mask:radial-gradient(farthest-side,transparent calc(100% - 3px),#000 calc(100% - 2px));
    box-shadow:0 0 32px rgba(57,215,255,.10);
    animation:brolyPageSpin 24s linear infinite;
  }

  .broly-page-ticks{
    position:absolute;
    inset:28px;
    border-radius:50%;
    z-index:0;
    opacity:.44;
    background:repeating-conic-gradient(from -2deg,rgba(188,244,255,.58) 0 1deg,transparent 1deg 6deg);
    -webkit-mask:radial-gradient(farthest-side,transparent calc(100% - 8px),#000 calc(100% - 7px),#000 calc(100% - 3px),transparent calc(100% - 2px));
    mask:radial-gradient(farthest-side,transparent calc(100% - 8px),#000 calc(100% - 7px),#000 calc(100% - 3px),transparent calc(100% - 2px));
    animation:brolyPageSpinReverse 58s linear infinite;
  }

  .broly-page-arc{
    position:absolute;
    border-radius:50%;
    z-index:1;
    pointer-events:none;
    border:2px solid transparent;
  }
  .broly-page-arc.a{
    inset:39px;
    border-top-color:rgba(57,215,255,.55);
    border-right-color:rgba(57,215,255,.10);
    transform:rotate(23deg);
    box-shadow:0 -2px 14px rgba(57,215,255,.16);
    animation:brolyPageArcA 13s ease-in-out infinite alternate;
  }
  .broly-page-arc.b{
    inset:56px;
    border-bottom-color:rgba(128,93,255,.50);
    border-left-color:rgba(128,93,255,.12);
    transform:rotate(-31deg);
    box-shadow:0 2px 14px rgba(128,93,255,.13);
    animation:brolyPageArcB 17s ease-in-out infinite alternate;
  }
  .broly-page-arc.c{
    inset:76px;
    border-top-color:rgba(214,250,255,.34);
    border-left-color:rgba(57,215,255,.18);
    transform:rotate(92deg);
    animation:brolyPageSpinReverse 21s linear infinite;
  }

  .broly-page-ring{
    position:absolute;
    inset:31px;
    border-radius:50%;
    z-index:1;
    border:1px solid rgba(57,215,255,.27);
    box-shadow:
      0 0 42px rgba(11,124,255,.17),
      inset 0 0 42px rgba(57,215,255,.075);
    animation:brolyPageSpin 21s linear infinite;
  }
  .broly-page-ring:before,
  .broly-page-ring:after{
    content:"";
    position:absolute;
    border-radius:50%;
  }
  .broly-page-ring:before{
    inset:24px;
    border:1px dashed rgba(57,215,255,.28);
    animation:brolyPageSpinReverse 14s linear infinite;
  }
  .broly-page-ring:after{
    inset:52px;
    border:1px solid rgba(128,93,255,.38);
    box-shadow:
      0 0 38px rgba(128,93,255,.12),
      inset 0 0 24px rgba(128,93,255,.045);
  }

  .broly-page-ring-inner{
    position:absolute;
    inset:87px;
    z-index:2;
    border-radius:50%;
    border:1px solid rgba(175,240,255,.18);
    box-shadow:
      0 0 24px rgba(57,215,255,.08),
      inset 0 0 20px rgba(57,215,255,.05);
  }
  .broly-page-ring-inner:before{
    content:"";
    position:absolute;
    inset:-9px;
    border-radius:50%;
    border:1px dotted rgba(57,215,255,.16);
    animation:brolyPageSpin 9s linear infinite;
  }

  .broly-page-sat{
    --sat:6px;
    position:absolute;
    left:50%;
    top:50%;
    width:var(--sat);
    height:var(--sat);
    margin:calc(var(--sat)/-2);
    border-radius:50%;
    z-index:7;
    background:#E4FCFF;
    box-shadow:
      0 0 7px #E4FCFF,
      0 0 16px rgba(57,215,255,.90),
      0 0 25px rgba(11,124,255,.42);
    transform-origin:0 0;
  }
  .broly-page-sat.a{animation:brolyPageOrbitA 8.5s linear infinite;}
  .broly-page-sat.b{--sat:4px;opacity:.72;animation:brolyPageOrbitB 12.5s linear infinite;}
  .broly-page-sat.c{--sat:3px;opacity:.52;animation:brolyPageOrbitC 17s linear infinite;}

  .broly-page-orb{
    position:relative;
    z-index:5;
    width:148px;
    height:148px;
    border-radius:50%;
    background:
      radial-gradient(circle at 50% 50%,rgba(232,253,255,.98) 0 3%,rgba(141,238,255,.88) 4% 8%,rgba(57,215,255,.50) 11%,transparent 19%),
      radial-gradient(circle at 41% 34%,rgba(222,252,255,.98) 0 1.4%,transparent 2.4%),
      radial-gradient(circle at 58% 48%,rgba(57,215,255,.92),rgba(11,124,255,.46) 31%,rgba(31,69,133,.17) 50%,rgba(5,12,25,.34) 63%,rgba(1,5,12,.98) 75%);
    border:1px solid rgba(112,230,255,.68);
    box-shadow:
      0 0 18px rgba(196,249,255,.34),
      0 0 48px rgba(57,215,255,.36),
      0 0 100px rgba(11,124,255,.29),
      inset 0 0 25px rgba(212,250,255,.18),
      inset 0 0 62px rgba(57,215,255,.14);
    animation:brolyPagePulse 3.8s ease-in-out infinite;
  }
  .broly-page-orb:before{
    content:"";
    position:absolute;
    inset:17px;
    border-radius:50%;
    border:1px solid rgba(205,249,255,.20);
    background:conic-gradient(from 90deg,transparent,rgba(57,215,255,.09),transparent 33%,rgba(128,93,255,.08),transparent 66%,rgba(57,215,255,.10),transparent);
    box-shadow:inset 0 0 24px rgba(57,215,255,.08);
    animation:brolyPageSpin 10s linear infinite;
  }
  .broly-page-orb:after{
    content:"";
    position:absolute;
    inset:44px;
    border-radius:50%;
    background:radial-gradient(circle,rgba(245,254,255,.98) 0 7%,rgba(151,241,255,.93) 9% 25%,rgba(57,215,255,.44) 31%,rgba(11,124,255,.08) 58%,transparent 70%);
    box-shadow:0 0 18px rgba(213,252,255,.58),0 0 36px rgba(57,215,255,.36);
    animation:brolyPageNucleus 2.4s ease-in-out infinite;
  }
  .broly-page-core-label{
    position:absolute;
    z-index:9;
    text-align:center;
    pointer-events:none;
    transform:translateY(2px);
  }
  .broly-page-core-name{
    color:#fff;
    font-size:18px;
    line-height:1;
    font-weight:950;
    letter-spacing:4px;
    text-shadow:
      0 0 7px rgba(238,254,255,.94),
      0 0 17px rgba(57,215,255,.82),
      0 0 28px rgba(11,124,255,.42);
  }
  .broly-page-core-state{
    margin-top:8px;
    color:rgba(215,245,253,.72);
    font-size:7px;
    line-height:1;
    font-weight:850;
    letter-spacing:1.45px;
    white-space:nowrap;
  }

  .broly-page-statebar{
    position:relative;
    z-index:15;
    display:flex;
    justify-content:center;
    align-items:center;
    gap:10px;
    margin-top:-3px;
    color:rgba(226,247,252,.88);
    font-size:10px;
    font-weight:900;
    letter-spacing:1.75px;
  }
  .broly-page-state-dot{
    width:7px;
    height:7px;
    border-radius:50%;
    background:#39D7FF;
    box-shadow:
      0 0 7px #39D7FF,
      0 0 16px rgba(57,215,255,.70);
    animation:brolyPageDot 2.3s ease-in-out infinite;
  }

  .broly-page-console-wrap{
    position:relative;
    z-index:20;
    max-width:860px;
    margin:24px auto 0;
    padding:0 28px 34px;
  }
  .broly-page-console{
    display:grid;
    grid-template-columns:58px minmax(0,1fr);
    gap:12px;
    align-items:center;
  }
  .broly-page-mic{
    width:58px;
    height:58px;
    box-sizing:border-box;
    display:flex;
    align-items:center;
    justify-content:center;
    border-radius:50%;
    border:1px solid rgba(57,215,255,.34);
    background:
      radial-gradient(circle at 50% 45%,rgba(57,215,255,.13),rgba(11,124,255,.06) 56%,rgba(3,10,20,.84) 100%);
    box-shadow:
      inset 0 0 18px rgba(57,215,255,.05),
      0 0 17px rgba(57,215,255,.08);
    color:rgba(208,247,255,.82);
    font-size:22px;
    filter:saturate(.95);
  }
  .broly-page-input-shell{
    position:relative;
    min-height:58px;
    box-sizing:border-box;
    display:flex;
    align-items:center;
    padding:0 18px;
    border-radius:16px;
    border:1px solid rgba(57,215,255,.20);
    background:linear-gradient(180deg,rgba(11,23,38,.86),rgba(7,16,28,.90));
    box-shadow:
      inset 0 0 0 1px rgba(255,255,255,.015),
      inset 0 0 19px rgba(57,215,255,.025);
  }
  .broly-page-input-shell:before{
    content:"";
    position:absolute;
    left:18px;
    right:18px;
    bottom:8px;
    height:1px;
    background:linear-gradient(90deg,rgba(57,215,255,.30),rgba(128,93,255,.18),transparent);
  }
  .broly-page-input-placeholder{
    color:rgba(184,213,224,.48);
    font-size:11px;
    font-weight:700;
    letter-spacing:.65px;
  }
  .broly-page-send{
    margin-left:auto;
    color:rgba(57,215,255,.44);
    font-size:14px;
    font-weight:900;
    letter-spacing:1px;
  }
  .broly-page-console-note{
    margin-top:10px;
    padding-left:70px;
    color:rgba(131,177,195,.38);
    font-size:7.2px;
    font-weight:800;
    letter-spacing:1.05px;
    line-height:1.5;
  }

  @keyframes brolyPageSpin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
  @keyframes brolyPageSpinReverse{from{transform:rotate(360deg)}to{transform:rotate(0deg)}}
  @keyframes brolyPagePulse{
    0%,100%{transform:scale(.985);filter:brightness(.92) saturate(.94)}
    50%{transform:scale(1.025);filter:brightness(1.16) saturate(1.10)}
  }
  @keyframes brolyPageNucleus{
    0%,100%{transform:scale(.88);opacity:.72}
    50%{transform:scale(1.13);opacity:1}
  }
  @keyframes brolyPageDot{
    0%,100%{opacity:.42}
    50%{opacity:1}
  }
  @keyframes brolyPageArcA{
    from{transform:rotate(18deg) scale(.995);opacity:.58}
    to{transform:rotate(42deg) scale(1.012);opacity:.95}
  }
  @keyframes brolyPageArcB{
    from{transform:rotate(-40deg);opacity:.48}
    to{transform:rotate(-17deg);opacity:.82}
  }
  @keyframes brolyPageOrbitA{
    from{transform:rotate(0deg) translateX(127px) rotate(0deg)}
    to{transform:rotate(360deg) translateX(127px) rotate(-360deg)}
  }
  @keyframes brolyPageOrbitB{
    from{transform:rotate(130deg) translateX(103px) rotate(-130deg)}
    to{transform:rotate(490deg) translateX(103px) rotate(-490deg)}
  }
  @keyframes brolyPageOrbitC{
    from{transform:rotate(245deg) translateX(80px) rotate(-245deg)}
    to{transform:rotate(605deg) translateX(80px) rotate(-605deg)}
  }

  @media (max-width:760px){
    .broly-page-shell{min-height:670px}
    .broly-page-head{padding:22px 20px 10px}
    .broly-page-title{font-size:20px}
    .broly-page-core-stage{min-height:340px;padding-top:30px}
    .broly-page-core-wrap{width:270px;height:270px}
    .broly-page-orb{width:126px;height:126px}
    .broly-page-console-wrap{padding:0 18px 28px}
    .broly-page-console{grid-template-columns:52px minmax(0,1fr);gap:10px}
    .broly-page-mic{width:52px;height:52px}
    .broly-page-input-shell{min-height:52px}
    .broly-page-console-note{padding-left:62px}
    .broly-page-sat.a{animation:none}
    .broly-page-sat.b{animation:none}
    .broly-page-sat.c{animation:none}
  }

  @media (prefers-reduced-motion:reduce){
    .broly-page-halo,.broly-page-ticks,.broly-page-arc,.broly-page-ring,
    .broly-page-ring:before,.broly-page-ring-inner:before,.broly-page-sat,
    .broly-page-orb,.broly-page-orb:before,.broly-page-orb:after,
    .broly-page-state-dot{animation:none !important}
  }
</style>
<div class="broly-page-shell">
  <div class="broly-page-head">
    <div class="broly-page-brand">
      <div class="broly-page-eyebrow">BROLY COGNITIVE SYSTEM</div>
      <div class="broly-page-title">BROLY // RENKEI</div>
      <div class="broly-page-sub">AI OPERATOR INTERFACE · ANALYTICS ORCHESTRATION NODE</div>
    </div>
    <div class="broly-page-node">
      <div class="broly-page-node-main">RENKEI NODE</div>
      <div class="broly-page-node-sub">LOCAL UI SHELL // PRE-API</div>
    </div>
  </div>

  <div class="broly-page-core-stage">
    <div class="broly-page-core-wrap" aria-hidden="true">
      <div class="broly-page-halo"></div>
      <div class="broly-page-ticks"></div>
      <div class="broly-page-arc a"></div>
      <div class="broly-page-arc b"></div>
      <div class="broly-page-arc c"></div>
      <div class="broly-page-ring"></div>
      <div class="broly-page-ring-inner"></div>
      <div class="broly-page-sat a"></div>
      <div class="broly-page-sat b"></div>
      <div class="broly-page-sat c"></div>
      <div class="broly-page-orb"></div>
      <div class="broly-page-core-label">
        <div class="broly-page-core-name">BROLY</div>
        <div class="broly-page-core-state">AI CORE · RENKEI NODE</div>
      </div>
    </div>
  </div>

  <div class="broly-page-statebar">
    <span class="broly-page-state-dot"></span>
    <span>STANDBY</span>
  </div>

  <div class="broly-page-console-wrap">
    <div class="broly-page-console">
      <div class="broly-page-mic" title="Voice interface — future connection">🎙</div>
      <div class="broly-page-input-shell">
        <span class="broly-page-input-placeholder">BROLYに指示する…</span>
        <span class="broly-page-send">›</span>
      </div>
    </div>
    <div class="broly-page-console-note">VOICE / CHAT INTERFACE · UI PROTOTYPE · API CONNECTION PENDING</div>

  </div>
</div>
""",
        unsafe_allow_html=True,
    )

RENKEI_READ_RPC = "renkei_get_monthly_metrics"
RENKEI_WRITE_RPC = "renkei_upsert_monthly_metrics"

# グラフ色（対象年度と比較年度を明確に区別）
COLOR_TARGET = "#F97316"      # 対象年度：オレンジ
COLOR_COMPARISON = "#38BDF8"  # 比較年度：シアン
COLOR_OTHER = "#94A3B8"       # その他：グレー
COLOR_ABOVE_PREV = "#2563EB"  # 前年度実績を上回る月：青
COLOR_BELOW_PREV = "#EF4444"  # 前年度実績を下回る月：赤
COLOR_EQUAL_PREV = "#94A3B8"  # 前年度実績と同値・比較不能：グレー
COLOR_PREV_LINE = "#CBD5E1"   # 前年度実績の折れ線：薄いグレー

FISCAL_MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
MONTH_LABELS = {m: f"{m}月" for m in FISCAL_MONTHS}
MONTH_ORDER = {m: i + 1 for i, m in enumerate(FISCAL_MONTHS)}


# =========================================================
# BROLY GLOBAL DESIGN SYSTEM / RENKEI
#   - 表示層とデータ入出力経路のみ更新。集計計算ロジックは変更しない。
# =========================================================
st.markdown(
    r"""
<style>
:root{
  --broly-bg:#050A12;
  --broly-bg-mid:#07101C;
  --broly-bg-end:#040810;
  --broly-panel:#06101C;
  --broly-panel-2:#091B2C;
  --broly-text:#E6EAF2;
  --broly-muted:rgba(190,214,229,.68);
  --broly-cyan:#39D7FF;
  --broly-blue:#0B7CFF;
  --broly-violet:#805DFF;
  --broly-border:rgba(57,215,255,.20);
  color-scheme:dark;
}
html,body,.stApp,[data-testid="stAppViewContainer"],[data-testid="stAppViewContainer"]>.main,section.main{
  background-color:var(--broly-bg)!important;
  color:var(--broly-text)!important;
}
.stApp,[data-testid="stAppViewContainer"],[data-testid="stAppViewContainer"]>.main{
  background:
    radial-gradient(980px 650px at 50% 18%,rgba(11,124,255,.17),transparent 68%),
    radial-gradient(720px 520px at 95% 4%,rgba(128,93,255,.10),transparent 64%),
    radial-gradient(620px 500px at 0% 38%,rgba(57,215,255,.07),transparent 64%),
    linear-gradient(180deg,var(--broly-bg) 0%,var(--broly-bg-mid) 52%,var(--broly-bg-end) 100%)!important;
}
[data-testid="stAppViewContainer"]>.main>div,section.main>div.block-container,[data-testid="stVerticalBlock"],[data-testid="stHorizontalBlock"]{background:transparent!important;}
[data-testid="stAppViewContainer"]::after{
  content:"";position:fixed;inset:0;pointer-events:none;z-index:0;opacity:.055;
  background-image:linear-gradient(rgba(57,215,255,.05) 1px,transparent 1px),linear-gradient(90deg,rgba(57,215,255,.035) 1px,transparent 1px);
  background-size:34px 34px;mask-image:linear-gradient(to bottom,rgba(0,0,0,.85),transparent 94%);
}
section.main>div.block-container{position:relative;z-index:1;max-width:1500px;padding-top:.72rem;padding-bottom:2rem;}
[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"]{background:transparent!important;border:none!important;box-shadow:none!important;}
#MainMenu{visibility:hidden;} footer{visibility:hidden;}
h1,h2,h3,h4,h5,h6,[data-testid="stHeadingWithActionElements"]{color:#FFFFFF!important;}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p,small{color:var(--broly-muted)!important;}

/* Sidebar */
[data-testid="stSidebar"]{background:linear-gradient(180deg,rgba(5,12,22,.98),rgba(4,10,18,.98))!important;border-right:1px solid rgba(57,215,255,.12)!important;}
[data-testid="stSidebar"] *{color:var(--broly-text);} 
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3{color:#FFFFFF!important;}

/* Controls */
[data-baseweb="input"]>div,[data-baseweb="base-input"],[data-baseweb="select"]>div,[data-baseweb="textarea"]>div,textarea,input{
  background-color:rgba(7,17,31,.92)!important;color:#FFFFFF!important;border-color:rgba(57,215,255,.18)!important;
}
[data-baseweb="select"] svg,[data-baseweb="input"] svg{color:rgba(57,215,255,.70)!important;}
[data-baseweb="popover"],[data-baseweb="menu"]{background:rgba(7,17,31,.99)!important;color:var(--broly-text)!important;}
[data-baseweb="menu"] li{background:transparent!important;color:var(--broly-text)!important;}
[data-baseweb="menu"] li:hover,[data-baseweb="menu"] li[aria-selected="true"]{background:rgba(57,215,255,.09)!important;color:#FFFFFF!important;}
[data-testid="stFileUploader"] section{background:rgba(7,17,31,.72)!important;border-color:rgba(57,215,255,.18)!important;}

/* Buttons */
.stButton>button,.stDownloadButton>button,[data-testid="stFormSubmitButton"]>button{
  color:#FFFFFF!important;background:linear-gradient(180deg,rgba(11,124,255,.22),rgba(8,20,35,.88))!important;
  border:1px solid rgba(57,215,255,.30)!important;border-radius:10px!important;box-shadow:0 8px 20px rgba(0,0,0,.16)!important;transition:all .14s ease!important;
}
.stButton>button:hover,.stDownloadButton>button:hover,[data-testid="stFormSubmitButton"]>button:hover{
  border-color:rgba(57,215,255,.58)!important;background:linear-gradient(180deg,rgba(11,124,255,.34),rgba(8,24,42,.94))!important;
  box-shadow:0 0 18px rgba(57,215,255,.12),0 10px 22px rgba(0,0,0,.20)!important;transform:translateY(-1px);
}

/* Metrics / containers */
div[data-testid="stMetric"],div[data-testid="stExpander"]{
  border:1px solid rgba(57,215,255,.16)!important;border-radius:13px!important;background:linear-gradient(180deg,rgba(6,16,28,.88),rgba(5,13,24,.72))!important;
  box-shadow:0 12px 28px rgba(0,0,0,.18)!important;
}
div[data-testid="stMetric"]{padding:.72rem .82rem!important;}
[data-testid="stMetricLabel"]{color:var(--broly-muted)!important;}
[data-testid="stMetricValue"]{color:#FFFFFF!important;}
div[data-testid="stExpander"] details,div[data-testid="stExpander"] summary{background:transparent!important;color:var(--broly-text)!important;}
div[data-testid="stExpander"] details summary p{color:#FFFFFF!important;font-weight:800!important;}
[data-testid="stAlert"]{border:1px solid rgba(57,215,255,.16)!important;border-radius:11px!important;background:rgba(6,16,28,.78)!important;color:var(--broly-text)!important;}

/* Dataframe wrapper only. Internal grid/canvas is intentionally untouched. */
[data-testid="stDataFrame"],[data-testid="stTable"]{border:1px solid rgba(57,215,255,.14)!important;border-radius:11px!important;overflow:hidden!important;background:#050A12!important;}
hr{border-color:rgba(57,215,255,.12)!important;}

.renkei-hero{position:relative;overflow:hidden;margin:.05rem 0 .95rem;padding:1.05rem 1.15rem 1rem;border:1px solid rgba(57,215,255,.20);border-radius:15px;background:linear-gradient(135deg,rgba(7,18,32,.92),rgba(5,13,24,.78));box-shadow:0 16px 36px rgba(0,0,0,.20),inset 0 1px 0 rgba(255,255,255,.025);}
.renkei-hero::before{content:"";position:absolute;left:0;top:0;width:100%;height:2px;background:linear-gradient(90deg,transparent,rgba(57,215,255,.86),rgba(128,93,255,.55),transparent);}
.renkei-kicker{font-size:.68rem;letter-spacing:.20em;font-weight:900;color:rgba(57,215,255,.67);margin-bottom:.28rem;}
.renkei-title{font-size:1.72rem;font-weight:950;letter-spacing:.045em;color:#FFFFFF;line-height:1.15;text-shadow:0 0 18px rgba(57,215,255,.10);}
.renkei-subtitle{margin-top:.35rem;color:rgba(200,220,233,.70);font-size:.88rem;}
.renkei-node{margin-top:.55rem;font-size:.70rem;letter-spacing:.14em;color:rgba(57,215,255,.56);font-weight:850;}
.small-caption{color:var(--broly-muted)!important;font-size:.88rem;}
</style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# ユーティリティ
# =========================================================
def safe_secret(key: str, default: str) -> str:
    """Streamlit secrets が未設定でも落ちないように取得。"""
    try:
        return str(st.secrets.get(key, default))
    except Exception:
        return default


def normalize_text(value: object) -> str:
    """列名・シート名の表記ゆれを吸収する。"""
    text = str(value).strip()
    text = text.replace(" ", "").replace("　", "")
    text = text.replace("_", "").replace("-", "").replace("／", "/")
    text = text.replace("（", "(").replace("）", ")")
    # Excel側で「稼働」ではなく「稼動」になっても拾えるように統一
    text = text.replace("稼動", "稼働")
    return text.lower()


def find_column(columns: Iterable[object], keywords: list[str], required_name: str) -> str:
    """列名の表記ゆれを吸収して、該当列を返す。"""
    normalized_columns = [(str(col), normalize_text(col)) for col in columns]
    normalized_keywords = [normalize_text(k) for k in keywords]

    # 完全一致を優先
    for original, norm_col in normalized_columns:
        if norm_col in normalized_keywords:
            return original

    # 部分一致
    for original, norm_col in normalized_columns:
        if any(key in norm_col for key in normalized_keywords):
            return original

    raise ValueError(
        f"必須列「{required_name}」が見つかりません。Excelの列名を確認してください。"
    )



def header_match_score(columns: Iterable[object]) -> int:
    """RENKEI様式らしいヘッダーかをざっくり判定する。"""
    norms = [normalize_text(c) for c in columns if str(c).strip() and str(c).lower() != "nan"]
    if not norms:
        return 0

    def has_any(keys: list[str]) -> bool:
        key_norms = [normalize_text(k) for k in keys]
        return any(any(k in n for k in key_norms) for n in norms)

    flags = [
        has_any(["年度", "年"]),
        has_any(["月番号", "月番"]),
        has_any(["月"]),
        has_any(["稼働日数", "稼働", "稼動日数", "稼動", "営業日数", "診療日数"]),
        has_any(["予約件数", "予約数", "予約", "件数"]),
    ]
    return int(sum(flags))


def read_excel_smart(excel_source) -> pd.DataFrame:
    """
    Excelの先頭シート固定で落ちないように、
    1) latest/データ系シートを優先
    2) RENKEI様式のヘッダーを持つシートを探索
    3) ヘッダー行が1行目でない場合も先頭15行から探索
    して読み込む。
    """
    xls = pd.ExcelFile(excel_source, engine="openpyxl")
    sheet_names = list(xls.sheet_names)

    def sheet_priority(sheet: str) -> tuple[int, str]:
        norm = normalize_text(sheet)
        preferred = ["latest", "data", "データ", "元データ", "入力", "実績"]
        if any(k in norm for k in preferred):
            return (0, sheet)
        if "ルール" in norm or "記入例" in norm or "説明" in norm:
            return (2, sheet)
        return (1, sheet)

    candidates: list[tuple[int, str, object, pd.DataFrame]] = []

    for sheet in sorted(sheet_names, key=sheet_priority):
        # 通常の1行目ヘッダー
        try:
            df0 = pd.read_excel(excel_source, sheet_name=sheet, engine="openpyxl")
            score0 = header_match_score(df0.columns)
            if score0 >= 4:
                df0.attrs["source_sheet_name"] = sheet
                df0.attrs["source_header_row"] = 1
                return df0
            candidates.append((score0, sheet, 1, df0))
        except Exception:
            pass

        # ヘッダー行が2行目以降にある場合
        try:
            preview = pd.read_excel(excel_source, sheet_name=sheet, header=None, nrows=15, engine="openpyxl")
            for idx, row in preview.iterrows():
                values = [v for v in row.tolist() if pd.notna(v)]
                score = header_match_score(values)
                if score >= 4:
                    dfh = pd.read_excel(excel_source, sheet_name=sheet, header=int(idx), engine="openpyxl")
                    dfh.attrs["source_sheet_name"] = sheet
                    dfh.attrs["source_header_row"] = int(idx) + 1
                    return dfh
        except Exception:
            pass

    # 最後の保険：一番スコアが高かった読み込み結果を返す
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        best = candidates[0][3]
        best.attrs["source_sheet_name"] = candidates[0][1]
        best.attrs["source_header_row"] = candidates[0][2]
        return best

    raise ValueError("Excel内に読み込めるシートがありません。")

def fiscal_year_sort_key(year: object) -> int:
    """R6, R7 のような年度表記を自然順にする。"""
    text = str(year)
    match = re.search(r"(\d+)", text)
    if match:
        return int(match.group(1))
    return 999


def parse_month(value: object) -> Optional[int]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    match = re.search(r"(\d{1,2})", text)
    if not match:
        return None
    month = int(match.group(1))
    return month if 1 <= month <= 12 else None


def format_number(value: Optional[float], digits: int = 0, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "-"
    if digits == 0:
        return f"{value:,.0f}{suffix}"
    return f"{value:,.{digits}f}{suffix}"


def format_delta(value: Optional[float], digits: int = 0, suffix: str = "") -> Optional[str]:
    if value is None or pd.isna(value):
        return None
    sign = "+" if value > 0 else ""
    if digits == 0:
        return f"{sign}{value:,.0f}{suffix}"
    return f"{sign}{value:,.{digits}f}{suffix}"


def get_month_range(start_month: int, end_month: int) -> list[int]:
    start_idx = FISCAL_MONTHS.index(start_month)
    end_idx = FISCAL_MONTHS.index(end_month)
    if start_idx <= end_idx:
        return FISCAL_MONTHS[start_idx : end_idx + 1]
    # 通常はUI順で起きないが、念のため年度内で一周できるようにする
    return FISCAL_MONTHS[start_idx:] + FISCAL_MONTHS[: end_idx + 1]


def get_series_color(year: str, years: list[str]) -> str:
    """対象年度と比較年度を固定色で見分けやすくする。"""
    if years and year == years[0]:
        return COLOR_TARGET
    if len(years) > 1 and year == years[1]:
        return COLOR_COMPARISON
    return COLOR_OTHER


def get_series_label(year: str, years: list[str]) -> str:
    """凡例に対象・比較の意味を補足する。"""
    if years and year == years[0]:
        return f"{year}（対象）"
    if len(years) > 1 and year == years[1]:
        return f"{year}（比較）"
    return str(year)


def sort_years_ascending(years: list[str]) -> list[str]:
    """グラフ表示・凡例表示は早い年度から並べる。"""
    return sorted(years, key=fiscal_year_sort_key)


# =========================================================
# Supabase / データ読込・更新
# =========================================================
def _supabase_config() -> tuple[str, str]:
    url = safe_secret("SUPABASE_URL", "").strip().rstrip("/")
    key = safe_secret(
        "SUPABASE_SECRET_KEY",
        safe_secret("SUPABASE_SERVICE_ROLE_KEY", ""),
    ).strip()
    return url, key


def _supabase_headers(secret_key: str) -> dict[str, str]:
    headers = {
        "apikey": secret_key,
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "streamlit-renkei",
    }
    # Legacy JWT service_role key only。新しい sb_secret_* は apikey のみで使用する。
    if secret_key.startswith("eyJ"):
        headers["Authorization"] = f"Bearer {secret_key}"
    return headers


def _supabase_rpc(rpc_name: str, payload: Optional[dict] = None, timeout: int = 30):
    url, key = _supabase_config()
    if not url or not key:
        raise RuntimeError("SUPABASE_URL / SUPABASE_SECRET_KEY が未設定です。")
    endpoint = f"{url}/rest/v1/rpc/{rpc_name}"
    response = requests.post(
        endpoint,
        headers=_supabase_headers(key),
        json=payload or {},
        timeout=timeout,
    )
    if not response.ok:
        detail = response.text[:1200]
        raise RuntimeError(f"Supabase RPC {rpc_name} エラー: HTTP {response.status_code} / {detail}")
    if not response.content:
        return None
    try:
        return response.json()
    except Exception:
        return response.text


@st.cache_data(ttl=60, show_spinner=False)
def load_from_supabase() -> pd.DataFrame:
    rows = _supabase_rpc(RENKEI_READ_RPC, {})
    if not isinstance(rows, list):
        raise RuntimeError(f"Supabase RPC {RENKEI_READ_RPC} の返却形式が不正です。")
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["年度", "月", "月番号", "稼働日数", "予約件数", "年度内順"])
    required = {"fiscal_year", "month_no", "working_days", "reservation_count"}
    missing = sorted(required.difference(df.columns))
    if missing:
        raise RuntimeError(f"Supabase RPC返却列が不足しています: {missing}")
    df = df.rename(columns={
        "fiscal_year": "年度",
        "month_no": "月番号",
        "working_days": "稼働日数",
        "reservation_count": "予約件数",
    })
    df["年度"] = df["年度"].astype(str).str.strip().str.upper()
    df["月番号"] = pd.to_numeric(df["月番号"], errors="coerce")
    df["稼働日数"] = pd.to_numeric(df["稼働日数"], errors="coerce")
    df["予約件数"] = pd.to_numeric(df["予約件数"], errors="coerce")
    df = df[df["年度"].notna() & df["月番号"].isin(FISCAL_MONTHS)].copy()
    df["月番号"] = df["月番号"].astype(int)
    df["月"] = df["月番号"].map(MONTH_LABELS)
    df["年度内順"] = df["月番号"].map(MONTH_ORDER)
    return df[["年度", "月", "月番号", "稼働日数", "予約件数", "年度内順"]].sort_values(
        ["年度", "年度内順"], key=lambda s: s.map(fiscal_year_sort_key) if s.name == "年度" else s
    ).reset_index(drop=True)


def load_from_upload(uploaded_file) -> pd.DataFrame:
    return read_excel_smart(uploaded_file)


def standardize_renkei_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df is None or raw_df.empty:
        raise ValueError("Excelにデータがありません。")

    source_sheet_name = raw_df.attrs.get("source_sheet_name", "不明")
    source_header_row = raw_df.attrs.get("source_header_row", "不明")

    # 空行・空列を整理
    df = raw_df.copy()
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]

    try:
        year_col = find_column(df.columns, ["年度", "年"], "年度")
        month_no_col = find_column(df.columns, ["月番号", "月番"], "月番号")
        # 月列はなくても月番号から作れるため任意扱い
        try:
            month_col = find_column(df.columns, ["月ツキ", "月"], "月")
        except Exception:
            month_col = None
        working_days_col = find_column(
            df.columns,
            ["稼働日数", "稼働", "稼動日数", "稼動", "営業日数", "診療日数"],
            "稼働日数",
        )
        reservation_col = find_column(
            df.columns,
            ["予約件数", "予約数", "予約", "件数"],
            "予約件数",
        )
    except ValueError as e:
        current_columns = " / ".join([str(c) for c in df.columns.tolist()])
        raise ValueError(
            f"{e}\n読み込みシート: {source_sheet_name}、ヘッダー行: {source_header_row}行目\n"
            f"現在アプリが認識している列名: {current_columns}"
        )

    rename_map = {
        year_col: "年度",
        month_no_col: "月番号",
        working_days_col: "稼働日数",
        reservation_col: "予約件数",
    }
    if month_col is not None:
        rename_map[month_col] = "月"

    df = df.rename(columns=rename_map)
    if "月" not in df.columns:
        df["月"] = pd.NA

    df = df[["年度", "月", "月番号", "稼働日数", "予約件数"]].copy()
    df["年度"] = df["年度"].astype(str).str.strip().str.upper()
    df["年度"] = df["年度"].replace({"NAN": pd.NA, "NONE": pd.NA, "": pd.NA})

    month_from_no = pd.to_numeric(df["月番号"], errors="coerce")
    month_from_label = df["月"].map(parse_month)
    df["月番号"] = month_from_no.fillna(month_from_label)
    df["月番号"] = pd.to_numeric(df["月番号"], errors="coerce")

    df["稼働日数"] = pd.to_numeric(df["稼働日数"], errors="coerce")
    df["予約件数"] = pd.to_numeric(df["予約件数"], errors="coerce")

    # 入力ミス対策：稼働日数と予約件数が逆の場合は自動補正
    auto_swapped_working_days_reservations = False
    valid_pair = df[df["稼働日数"].notna() & df["予約件数"].notna()].copy()
    if not valid_pair.empty:
        working_days_large_ratio = (valid_pair["稼働日数"] > 100).mean()
        reservations_small_ratio = (valid_pair["予約件数"] <= 31).mean()
        if working_days_large_ratio >= 0.5 and reservations_small_ratio >= 0.5:
            df[["稼働日数", "予約件数"]] = df[["予約件数", "稼働日数"]]
            auto_swapped_working_days_reservations = True

    df = df[df["年度"].notna() & df["月番号"].isin(FISCAL_MONTHS)].copy()
    df["月番号"] = df["月番号"].astype(int)
    df["月"] = df["月番号"].map(MONTH_LABELS)
    df["年度内順"] = df["月番号"].map(MONTH_ORDER)

    def sum_keep_blank(series: pd.Series):
        return series.sum(min_count=1)

    # 同じ年度・月が複数ある場合は合算。予約件数が全空欄の月は空欄を維持。
    df = (
        df.groupby(["年度", "月番号", "月", "年度内順"], as_index=False)
        .agg({"稼働日数": "max", "予約件数": sum_keep_blank})
        .sort_values(["年度", "年度内順"])
    )

    df.attrs["auto_swapped_working_days_reservations"] = auto_swapped_working_days_reservations
    df.attrs["source_sheet_name"] = source_sheet_name
    df.attrs["source_header_row"] = source_header_row
    return df


def _is_integer_like(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    return series.isna() | (numeric.notna() & numeric.mod(1).eq(0))


def validate_renkei_update_data(df: pd.DataFrame) -> list[str]:
    """正式更新前の構造・値検証。最新Excelは、含めた年度ごとに4〜3月の12行を必須とする。"""
    errors: list[str] = []
    if df is None or df.empty:
        return ["更新対象データが空です。"]

    if not df["年度"].astype(str).str.fullmatch(r"R\d+").all():
        errors.append("年度は R6 / R7 のような令和年度表記にしてください。")
    if not df["月番号"].isin(FISCAL_MONTHS).all():
        errors.append("月番号は1〜12の範囲で指定してください。")
    if df.duplicated(["年度", "月番号"]).any():
        errors.append("同じ年度・月が重複しています。")

    for col in ["稼働日数", "予約件数"]:
        numeric = pd.to_numeric(df[col], errors="coerce")
        if (~_is_integer_like(df[col])).any():
            errors.append(f"{col}は整数または空欄で入力してください。")
        if (numeric.dropna() < 0).any():
            errors.append(f"{col}に負の値は入力できません。")

    expected = set(FISCAL_MONTHS)
    for year, group in df.groupby("年度", dropna=False):
        months = set(pd.to_numeric(group["月番号"], errors="coerce").dropna().astype(int).tolist())
        if months != expected:
            missing = [m for m in FISCAL_MONTHS if m not in months]
            extra = sorted(months.difference(expected))
            detail = []
            if missing:
                detail.append("不足=" + ",".join(f"{m}月" for m in missing))
            if extra:
                detail.append("範囲外=" + ",".join(map(str, extra)))
            errors.append(f"{year} は4〜3月の12行が必要です（{' / '.join(detail)}）。")
    return errors


def build_renkei_update_diff(current_df: pd.DataFrame, candidate_df: pd.DataFrame) -> pd.DataFrame:
    """アップロード候補と現在の正式値を年度+月で比較する。欠損同士は一致扱い。"""
    current = current_df[["年度", "月番号", "稼働日数", "予約件数"]].rename(columns={
        "稼働日数": "現在_稼働日数",
        "予約件数": "現在_予約件数",
    })
    candidate = candidate_df[["年度", "月", "月番号", "稼働日数", "予約件数", "年度内順"]].rename(columns={
        "稼働日数": "更新後_稼働日数",
        "予約件数": "更新後_予約件数",
    })
    diff = candidate.merge(current, on=["年度", "月番号"], how="left", indicator=True)

    def same_values(left: pd.Series, right: pd.Series) -> pd.Series:
        return (left.eq(right)) | (left.isna() & right.isna())

    same_working = same_values(diff["更新後_稼働日数"], diff["現在_稼働日数"])
    same_reservation = same_values(diff["更新後_予約件数"], diff["現在_予約件数"])
    diff["判定"] = "変更"
    diff.loc[diff["_merge"].eq("left_only"), "判定"] = "新規"
    diff.loc[diff["_merge"].eq("both") & same_working & same_reservation, "判定"] = "変更なし"
    diff = diff.drop(columns=["_merge"])
    return diff.sort_values(
        ["年度", "年度内順"],
        key=lambda s: s.map(fiscal_year_sort_key) if s.name == "年度" else s,
    ).reset_index(drop=True)


def renkei_update_records(candidate_df: pd.DataFrame) -> list[dict]:
    records: list[dict] = []
    for row in candidate_df.sort_values(["年度", "年度内順"]).itertuples(index=False):
        working_days = getattr(row, "稼働日数")
        reservation_count = getattr(row, "予約件数")
        records.append({
            "fiscal_year": str(getattr(row, "年度")),
            "month_no": int(getattr(row, "月番号")),
            "working_days": None if pd.isna(working_days) else int(working_days),
            "reservation_count": None if pd.isna(reservation_count) else int(reservation_count),
        })
    return records


def save_renkei_update(candidate_df: pd.DataFrame):
    records = renkei_update_records(candidate_df)
    return _supabase_rpc(RENKEI_WRITE_RPC, {"p_rows": records}, timeout=60)


def filter_period(df: pd.DataFrame, years: list[str], months: list[int]) -> pd.DataFrame:
    return df[df["年度"].isin(years) & df["月番号"].isin(months)].copy()


def calc_metrics(df: pd.DataFrame) -> dict[str, Optional[float]]:
    actual = df[df["予約件数"].notna()].copy()
    if actual.empty:
        return {
            "latest_count": None,
            "latest_month": None,
            "total": None,
            "monthly_average": None,
            "daily_average": None,
            "actual_months": 0,
            "working_days": None,
        }

    actual = actual.sort_values("年度内順")
    total = float(actual["予約件数"].sum())
    actual_months = int(actual["月番号"].nunique())
    working_days = actual["稼働日数"].sum(skipna=True)
    working_days = float(working_days) if pd.notna(working_days) and working_days > 0 else None

    latest_row = actual.tail(1).iloc[0]
    return {
        "latest_count": float(latest_row["予約件数"]),
        "latest_month": str(latest_row["月"]),
        "total": total,
        "monthly_average": total / actual_months if actual_months > 0 else None,
        "daily_average": total / working_days if working_days else None,
        "actual_months": actual_months,
        "working_days": working_days,
    }


# =========================================================
# グラフ
# =========================================================
def apply_common_layout(fig: go.Figure, title: str, y_title: str) -> go.Figure:
    fig.update_layout(
        title={"text": title, "x": 0.02, "xanchor": "left", "y": 0.98, "yanchor": "top", "font": {"size": 18, "color": "#FFFFFF"}},
        height=420,
        margin=dict(l=72, r=155, t=82, b=68),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E6EAF2", size=13),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1.0,
            xanchor="left",
            x=1.02,
            traceorder="normal",
            font=dict(size=13, color="#E6EAF2"),
            bgcolor="rgba(5,10,18,0.72)",
            bordercolor="rgba(57,215,255,0.20)",
            borderwidth=1,
        ),
        hovermode="x unified",
        hoverlabel=dict(bgcolor="#07111F", font_color="#FFFFFF", bordercolor="rgba(57,215,255,0.35)"),
    )
    fig.update_xaxes(
        title_text="月",
        categoryorder="array",
        categoryarray=[MONTH_LABELS[m] for m in FISCAL_MONTHS],
        tickfont=dict(size=14, color="#D8E7F0"),
        title_font=dict(size=15, color="#FFFFFF"),
        linecolor="rgba(190,214,229,0.22)",
        gridcolor="rgba(190,214,229,0.08)",
        zerolinecolor="rgba(190,214,229,0.12)",
        automargin=True,
    )
    fig.update_yaxes(
        title_text=y_title,
        rangemode="tozero",
        tickfont=dict(size=14, color="#D8E7F0"),
        title_font=dict(size=15, color="#FFFFFF"),
        linecolor="rgba(190,214,229,0.22)",
        gridcolor="rgba(190,214,229,0.08)",
        zerolinecolor="rgba(190,214,229,0.12)",
        automargin=True,
    )
    return fig


def get_role_years(role_years: Optional[list[str]], years: list[str]) -> tuple[Optional[str], Optional[str], dict[str, int]]:
    """
    混合グラフ用の年度役割を返す。
    role_years は [対象年度, 比較年度] の想定。
    - 対象年度：棒グラフ
    - 比較年度：折れ線グラフ
    凡例は早い年度から表示する。
    """
    role_years = role_years or years
    target_year = role_years[0] if len(role_years) >= 1 else None
    comparison_year = role_years[1] if len(role_years) >= 2 else None
    ordered = sort_years_ascending([y for y in [target_year, comparison_year] if y is not None])
    legend_rank = {year: idx + 1 for idx, year in enumerate(ordered)}
    return target_year, comparison_year, legend_rank


def compare_bar_colors(
    target_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
    y_col: str,
) -> list[str]:
    """
    対象年度の棒色を、同じ月の比較年度実績と比べて自動判定する。
    - 上回る：青
    - 下回る：赤
    - 同値または比較年度データなし：グレー
    """
    comparison_values = comparison_df.set_index("月番号")[y_col].to_dict() if not comparison_df.empty else {}
    colors: list[str] = []
    for _, row in target_df.iterrows():
        target_value = row.get(y_col)
        comparison_value = comparison_values.get(row.get("月番号"))
        if pd.isna(target_value) or comparison_value is None or pd.isna(comparison_value):
            colors.append(COLOR_EQUAL_PREV)
        elif float(target_value) > float(comparison_value):
            colors.append(COLOR_ABOVE_PREV)
        elif float(target_value) < float(comparison_value):
            colors.append(COLOR_BELOW_PREV)
        else:
            colors.append(COLOR_EQUAL_PREV)
    return colors


def add_target_bar_trace(
    fig: go.Figure,
    target_df: pd.DataFrame,
    comparison_df: pd.DataFrame,
    year: str,
    y_col: str,
    role_years: list[str],
    legend_rank: dict[str, int],
) -> None:
    """対象年度を棒で表示し、前年度比に応じて棒色を青・赤にする。"""
    bar_colors = compare_bar_colors(target_df, comparison_df, y_col)
    fig.add_trace(
        go.Bar(
            x=target_df["月"],
            y=target_df[y_col],
            name=f"{get_series_label(year, role_years)}｜棒",
            marker_color=bar_colors,
            opacity=0.82,
            text=target_df[y_col],
            texttemplate="%{text:,.1f}" if y_col == "1日平均" else "%{text:,.0f}",
            textposition="outside",
            legendrank=legend_rank.get(year, 99),
            hovertemplate=(
                "年度=%{fullData.name}<br>"
                "月=%{x}<br>"
                f"{y_col}=%{{y:,.1f}}" if y_col == "1日平均" else
                "年度=%{fullData.name}<br>"
                "月=%{x}<br>"
                f"{y_col}=%{{y:,.0f}}"
            ),
        )
    )


def add_comparison_line_trace(
    fig: go.Figure,
    comparison_df: pd.DataFrame,
    year: str,
    y_col: str,
    role_years: list[str],
    legend_rank: dict[str, int],
) -> None:
    """比較年度を折れ線で表示する。"""
    fig.add_trace(
        go.Scatter(
            x=comparison_df["月"],
            y=comparison_df[y_col],
            mode="lines+markers",
            name=f"{get_series_label(year, role_years)}｜線",
            connectgaps=False,
            line=dict(color=COLOR_PREV_LINE, width=3),
            marker=dict(color=COLOR_PREV_LINE, size=8),
            legendrank=legend_rank.get(year, 99),
        )
    )


def build_mixed_year_chart(
    actual: pd.DataFrame,
    y_col: str,
    title: str,
    y_title: str,
    years: list[str],
    role_years: Optional[list[str]] = None,
) -> go.Figure:
    """
    対象年度を棒、比較年度を折れ線で表示する混合グラフ。
    棒色は同じ月の比較年度実績を上回れば青、下回れば赤。
    """
    role_years = role_years or years
    target_year, comparison_year, legend_rank = get_role_years(role_years, years)
    actual = actual[actual[y_col].notna()].sort_values(["年度", "年度内順"]).copy()

    fig = go.Figure()

    comparison_df = pd.DataFrame()
    if comparison_year in years:
        comparison_df = actual[actual["年度"] == comparison_year].copy()

    # 棒を先に描画し、折れ線を後から重ねる。凡例順は legendrank で早い年度順に制御する。
    if target_year in years:
        target_df = actual[actual["年度"] == target_year].copy()
        if not target_df.empty:
            add_target_bar_trace(fig, target_df, comparison_df, target_year, y_col, role_years, legend_rank)

    if comparison_year in years and not comparison_df.empty:
        add_comparison_line_trace(fig, comparison_df, comparison_year, y_col, role_years, legend_rank)

    # 比較年度がない場合や2年度以外の保険。対象・比較以外は従来に近い折れ線表示。
    extra_years = [y for y in years if y not in [target_year, comparison_year]]
    for year in extra_years:
        extra_df = actual[actual["年度"] == year]
        if not extra_df.empty:
            fig.add_trace(
                go.Scatter(
                    x=extra_df["月"],
                    y=extra_df[y_col],
                    mode="lines+markers",
                    name=get_series_label(year, role_years),
                    connectgaps=False,
                    line=dict(color=get_series_color(year, role_years), width=3),
                    marker=dict(color=get_series_color(year, role_years), size=7),
                    legendrank=legend_rank.get(year, 99),
                )
            )

    fig.update_layout(barmode="overlay")
    return apply_common_layout(fig, title, y_title)


def build_monthly_trend_chart(chart_df: pd.DataFrame, years: list[str], role_years: Optional[list[str]] = None) -> go.Figure:
    actual = chart_df[chart_df["予約件数"].notna()].sort_values(["年度", "年度内順"]).copy()
    return build_mixed_year_chart(
        actual=actual,
        y_col="予約件数",
        title="月推移｜予約件数（棒：対象年度／線：比較年度）",
        y_title="予約件数",
        years=years,
        role_years=role_years,
    )


def build_monthly_average_chart(chart_df: pd.DataFrame, years: list[str], role_years: Optional[list[str]] = None) -> go.Figure:
    role_years = role_years or years
    bars = []
    for year in years:
        d = chart_df[(chart_df["年度"] == year) & chart_df["予約件数"].notna()]
        avg = d["予約件数"].mean() if not d.empty else None
        bars.append({"年度": year, "月平均": avg})
    bar_df = pd.DataFrame(bars)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=bar_df["年度"],
            y=bar_df["月平均"],
            text=bar_df["月平均"],
            texttemplate="%{text:,.1f}",
            textposition="outside",
            marker_color=[get_series_color(y, role_years) for y in bar_df["年度"]],
        )
    )
    fig.update_layout(
        title={"text": "月平均｜期間内の実績月平均", "x": 0.02, "xanchor": "left", "font": {"size": 18, "color": "#FFFFFF"}},
        height=420,
        margin=dict(l=72, r=36, t=82, b=68),
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E6EAF2", size=13),
        hoverlabel=dict(bgcolor="#07111F", font_color="#FFFFFF", bordercolor="rgba(57,215,255,0.35)"),
    )
    fig.update_xaxes(
        tickfont=dict(size=14, color="#D8E7F0"),
        title_font=dict(size=15, color="#FFFFFF"),
        linecolor="rgba(190,214,229,0.22)",
        gridcolor="rgba(190,214,229,0.08)",
        automargin=True,
    )
    fig.update_yaxes(
        title_text="月平均予約件数",
        rangemode="tozero",
        tickfont=dict(size=14, color="#D8E7F0"),
        title_font=dict(size=15, color="#FFFFFF"),
        linecolor="rgba(190,214,229,0.22)",
        gridcolor="rgba(190,214,229,0.08)",
        zerolinecolor="rgba(190,214,229,0.12)",
        automargin=True,
    )
    return fig


def build_daily_average_chart(chart_df: pd.DataFrame, years: list[str], role_years: Optional[list[str]] = None) -> go.Figure:
    actual = chart_df[
        chart_df["予約件数"].notna() & chart_df["稼働日数"].notna() & (chart_df["稼働日数"] > 0)
    ].copy()
    actual["1日平均"] = actual["予約件数"] / actual["稼働日数"]
    actual = actual.sort_values(["年度", "年度内順"])
    return build_mixed_year_chart(
        actual=actual,
        y_col="1日平均",
        title="1日平均｜予約件数 ÷ 稼働日数（棒：対象年度／線：比較年度）",
        y_title="1日平均予約件数",
        years=years,
        role_years=role_years,
    )


def build_cumulative_chart(chart_df: pd.DataFrame, years: list[str], role_years: Optional[list[str]] = None) -> go.Figure:
    actual = chart_df[chart_df["予約件数"].notna()].sort_values(["年度", "年度内順"]).copy()
    actual["累計"] = actual.groupby("年度")["予約件数"].cumsum()
    return build_mixed_year_chart(
        actual=actual,
        y_col="累計",
        title="累計｜期間内予約件数（棒：対象年度／線：比較年度）",
        y_title="累計予約件数",
        years=years,
        role_years=role_years,
    )


# =========================================================
# 画面本体
# =========================================================
with st.sidebar:
    st.header("画面選択")
    renkei_view = st.radio(
        "画面",
        ["BROLY", "予約業務実績"],
        index=1,
        key="renkei_view",
    )

# BROLYページはUIのみ。API・音声・Context Managerにはまだ未接続。
# Supabaseアクセス・集計処理に入る前に停止する。
if renkei_view == "BROLY":
    render_broly_page()
    st.stop()

st.markdown(
    """
    <div class="renkei-hero">
      <div class="renkei-kicker">BROLY COGNITIVE SYSTEM // REGIONAL COORDINATION</div>
      <div class="renkei-title">RENKEI</div>
      <div class="renkei-subtitle">地域連携室 予約業務実績ダッシュボード</div>
      <div class="renkei-node">ANALYTICS NODE // SUPABASE RPC</div>
    </div>
    """,
    unsafe_allow_html=True,
)
with st.sidebar:
    st.header("データ更新")
    selected_update_file = st.file_uploader(
        "最新データExcelを選択",
        type=["xlsx"],
        help="ここではファイルを選択するだけです。下のアップロードボタンを押した後にPreviewを作成します。",
        key="renkei_update_file_selector",
    )
    upload_for_preview = st.button(
        "アップロードしてPreview",
        type="primary",
        disabled=selected_update_file is None,
        use_container_width=True,
        key="renkei_upload_for_preview_button",
    )
    if selected_update_file is not None and not upload_for_preview:
        st.caption("ファイル選択後、「アップロードしてPreview」を押してください。")

if upload_for_preview and selected_update_file is not None:
    selected_bytes = selected_update_file.getvalue()
    st.session_state["renkei_pending_upload"] = {
        "name": selected_update_file.name,
        "bytes": selected_bytes,
        "sha256": hashlib.sha256(selected_bytes).hexdigest(),
    }

try:
    df = load_from_supabase()
except Exception as e:
    st.error("SupabaseのRENKEI正式データを読み込めませんでした。Secrets・RPC・権限を確認してください。")
    st.exception(e)
    st.stop()

auto_swapped = False

# ---------------------------------------------------------
# 最新Excel → Preview → 人間承認 → 固定Write RPC
# アップロードデータを分析値へ直接使用せず、正式保存成功後にRead RPCから再取得する。
# ---------------------------------------------------------
pending_upload = st.session_state.get("renkei_pending_upload")
if pending_upload is not None:
    upload_bytes = pending_upload["bytes"]
    upload_sha256 = pending_upload["sha256"]
    upload_file_name = pending_upload["name"]
    try:
        update_raw = load_from_upload(io.BytesIO(upload_bytes))
        update_df = standardize_renkei_data(update_raw)
        update_auto_swapped = bool(update_df.attrs.get("auto_swapped_working_days_reservations", False))
        update_errors = validate_renkei_update_data(update_df)
    except Exception as e:
        update_df = pd.DataFrame()
        update_auto_swapped = False
        update_errors = [f"Excel読込・整形エラー: {e}"]

    with st.expander("最新データ更新 Preview", expanded=True):
        st.caption(f"ファイル: {upload_file_name} / SHA-256: {upload_sha256}")
        if update_auto_swapped:
            st.warning(
                "アップロードExcelで「稼働日数」と「予約件数」が逆に入力されている可能性が高いため、"
                "従来仕様どおりアプリ側で自動補正しました。正式更新前に内容を確認してください。"
            )
        if update_errors:
            for message in update_errors:
                st.error(message)
        else:
            diff_df = build_renkei_update_diff(df, update_df)
            changed_df = diff_df[diff_df["判定"].ne("変更なし")].copy()
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("アップロード行", f"{len(update_df):,}行")
            s2.metric("変更", f"{int(diff_df['判定'].eq('変更').sum()):,}行")
            s3.metric("新規", f"{int(diff_df['判定'].eq('新規').sum()):,}行")
            s4.metric("変更なし", f"{int(diff_df['判定'].eq('変更なし').sum()):,}行")

            year_preview = (
                update_df.groupby("年度", as_index=False)
                .agg(
                    行数=("月番号", "size"),
                    実績月数=("予約件数", "count"),
                    稼働日数合計=("稼働日数", lambda s: s.sum(min_count=1)),
                    予約件数合計=("予約件数", lambda s: s.sum(min_count=1)),
                )
            )
            st.markdown("**更新後の年度別サマリ**")
            st.dataframe(year_preview, use_container_width=True, hide_index=True)

            st.markdown("**現在値との差分**")
            if changed_df.empty:
                st.info("現在のSupabase正式値と同一です。更新は不要です。")
            else:
                st.dataframe(
                    changed_df[[
                        "年度", "月", "判定",
                        "現在_稼働日数", "更新後_稼働日数",
                        "現在_予約件数", "更新後_予約件数",
                    ]],
                    use_container_width=True,
                    hide_index=True,
                )

                approval_key = f"renkei_update_approval_{upload_sha256[:16]}"
                approved = st.checkbox(
                    "Preview内容と差分を確認し、このExcelの値でSupabaseを更新する",
                    value=False,
                    key=approval_key,
                )
                if st.button(
                    "Supabaseへ正式更新",
                    type="primary",
                    disabled=not approved,
                    key=f"renkei_update_button_{upload_sha256[:16]}",
                ):
                    # Preview時に固定したbytesを正式更新直前にも再ハッシュする。
                    review_sha256 = hashlib.sha256(upload_bytes).hexdigest()
                    if review_sha256 != upload_sha256:
                        st.error("Preview保持データのSHA-256が変化したため更新を中止しました。再アップロードしてください。")
                    else:
                        try:
                            result = save_renkei_update(update_df)
                            load_from_supabase.clear()
                            st.session_state.pop("renkei_pending_upload", None)
                            st.success(f"Supabase更新成功: {result}")
                            st.rerun()
                        except Exception as e:
                            st.error("Supabase更新に失敗しました。DBの正式値は更新結果を確認するまで確定扱いにしないでください。")
                            st.exception(e)

all_years = sorted(df["年度"].dropna().unique().tolist(), key=fiscal_year_sort_key)
years_with_actual = sorted(df[df["予約件数"].notna()]["年度"].dropna().unique().tolist(), key=fiscal_year_sort_key)

if not all_years:
    st.warning("年度データがありません。テンプレートに年度・月番号・稼働日数・予約件数を入力してください。")
    st.stop()

# デフォルト年度
if years_with_actual:
    default_target = years_with_actual[-1]
else:
    default_target = all_years[-1]

default_comp = all_years[max(0, all_years.index(default_target) - 1)] if default_target in all_years else all_years[0]

with st.sidebar:
    st.header("分析条件")
    target_year = st.selectbox("対象年度", all_years, index=all_years.index(default_target))

    comparison_candidates = [y for y in all_years if y != target_year]
    if not comparison_candidates:
        comparison_candidates = [target_year]
    comparison_index = comparison_candidates.index(default_comp) if default_comp in comparison_candidates else 0
    comparison_year = st.selectbox("比較年度", comparison_candidates, index=comparison_index)

    month_options = [MONTH_LABELS[m] for m in FISCAL_MONTHS]
    start_label = st.selectbox("開始月", month_options, index=0)

    # 実績がある最後の月を終了月の初期値にする
    target_actual = df[(df["年度"] == target_year) & df["予約件数"].notna()].sort_values("年度内順")
    if not target_actual.empty:
        default_end_month = int(target_actual.tail(1).iloc[0]["月番号"])
    else:
        default_end_month = 3
    end_label = st.selectbox("終了月", month_options, index=FISCAL_MONTHS.index(default_end_month))

start_month = int(start_label.replace("月", ""))
end_month = int(end_label.replace("月", ""))
selected_months = get_month_range(start_month, end_month)
selected_years = [target_year, comparison_year]
display_years = sort_years_ascending(selected_years)
period_df = filter_period(df, selected_years, selected_months)

target_df = period_df[period_df["年度"] == target_year]
comp_df = period_df[period_df["年度"] == comparison_year]
target_metrics = calc_metrics(target_df)
comp_metrics = calc_metrics(comp_df)

if target_metrics["actual_months"] == 0:
    st.warning(f"{target_year} の {start_label}〜{end_label} に予約件数の実績がありません。Excelに予約件数を入力してください。")

# KPI cards
col1, col2, col3, col4 = st.columns(4)
with col1:
    latest_label = "最新月件数"
    if target_metrics["latest_month"]:
        latest_label += f"（{target_metrics['latest_month']}）"
    latest_delta = None
    if comp_metrics["latest_count"] is not None and target_metrics["latest_count"] is not None:
        latest_delta = target_metrics["latest_count"] - comp_metrics["latest_count"]
    st.metric(latest_label, format_number(target_metrics["latest_count"], 0, "件"), format_delta(latest_delta, 0, "件"))
    st.caption(f"{comparison_year}: {format_number(comp_metrics['latest_count'], 0, '件')}")

with col2:
    total_delta = None
    if comp_metrics["total"] is not None and target_metrics["total"] is not None:
        total_delta = target_metrics["total"] - comp_metrics["total"]
    st.metric("期間累計", format_number(target_metrics["total"], 0, "件"), format_delta(total_delta, 0, "件"))
    st.caption(f"{comparison_year}: {format_number(comp_metrics['total'], 0, '件')}")

with col3:
    avg_delta = None
    if comp_metrics["monthly_average"] is not None and target_metrics["monthly_average"] is not None:
        avg_delta = target_metrics["monthly_average"] - comp_metrics["monthly_average"]
    st.metric("月平均", format_number(target_metrics["monthly_average"], 1, "件"), format_delta(avg_delta, 1, "件"))
    st.caption(f"{comparison_year}: {format_number(comp_metrics['monthly_average'], 1, '件')}")

with col4:
    day_delta = None
    if comp_metrics["daily_average"] is not None and target_metrics["daily_average"] is not None:
        day_delta = target_metrics["daily_average"] - comp_metrics["daily_average"]
    st.metric("1日平均", format_number(target_metrics["daily_average"], 1, "件"), format_delta(day_delta, 1, "件"))
    st.caption(f"{comparison_year}: {format_number(comp_metrics['daily_average'], 1, '件')}")

st.divider()

left, right = st.columns(2)
with left:
    st.plotly_chart(build_monthly_trend_chart(period_df, display_years, selected_years), use_container_width=True)
with right:
    st.plotly_chart(build_monthly_average_chart(period_df, display_years, selected_years), use_container_width=True)

left, right = st.columns(2)
with left:
    st.plotly_chart(build_daily_average_chart(period_df, display_years, selected_years), use_container_width=True)
with right:
    st.plotly_chart(build_cumulative_chart(period_df, display_years, selected_years), use_container_width=True)

with st.expander("集計データを確認", expanded=False):
    display_df = period_df.sort_values(["年度", "年度内順"])[["年度", "月", "月番号", "稼働日数", "予約件数"]]
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "集計データCSVをダウンロード",
        data=csv,
        file_name="renkei_filtered_data.csv",
        mime="text/csv",
    )

st.caption("RENKEI v7：正式データはSupabase Read RPC。最新Excelは選択→明示的アップロード→Preview→人間承認後にWrite RPCで更新。未入力月の予約件数NULLは0件扱いしない。月推移・1日平均・累計は、対象年度を棒、比較年度を折れ線で表示。")
