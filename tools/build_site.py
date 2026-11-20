r"""build_site.py - GitHub Pages 용 웹 교안을 만든다 (강사용)

    python tools\build_site.py

docs/basic/*.md (PDF 와 같은 원고)를 HTML 로 바꿔
  docs/index.html        첫 화면 (과정 소개, 목차, 내려받기)
  docs/basic/ch00.html … 장별 읽기 화면 (휴대폰에서도 읽기 좋게)
를 만든다. GitHub Pages 는 main 브랜치의 /docs 폴더를 공개하도록 설정한다.
원고(.md)를 고친 뒤 이 스크립트를 다시 실행하고 push 하면 사이트가 바뀐다.
"""
import html
import re
from pathlib import Path
from urllib.parse import quote

import markdown

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SRC = DOCS / "basic"
REPO = "sinaboro/power-bi-tutorial"
SITE = f"https://sinaboro.github.io/power-bi-tutorial/"
RAW = f"https://github.com/{REPO}/raw/main/"            # 저장소 파일 직접 내려받기 주소
ZIP = f"https://github.com/{REPO}/archive/refs/heads/main.zip"
PDF = "PowerBI_기초_학습교안.pdf"

# 원고 파일 → (웹 파일 이름, 메뉴에 보일 짧은 이름, 학습 시간)
CHAPTERS = [
    ("00_과정안내.md", "ch00", "0장 과정 안내", "1시간"),
    ("01_시작하기와데이터불러오기.md", "ch01", "1장 데이터 불러오기", "2시간"),
    ("02_PowerQuery로정리하기.md", "ch02", "2장 Power Query 정리", "3시간"),
    ("03_관계와날짜테이블.md", "ch03", "3장 관계와 날짜 테이블", "2시간"),
    ("04_측정값만들기.md", "ch04", "4장 측정값 (DAX)", "2.5시간"),
    ("05_차트만들기.md", "ch05", "5장 차트 만들기", "2.5시간"),
    ("06_슬라이서와서식.md", "ch06", "6장 슬라이서와 서식", "2시간"),
    ("07_대시보드완성과공유.md", "ch07", "7장 대시보드 완성", "1.5시간"),
    ("08_과제_편의점대시보드.md", "ch08", "8장 과제 - 편의점", "4시간"),
    ("A_오류해결.md", "appA", "부록 A 오류 해결", ""),
    ("B_용어풀이.md", "appB", "부록 B 용어 풀이", ""),
    ("C_확인문제풀이.md", "appC", "부록 C 확인 문제 풀이", ""),
]

# 장마다 대표 그림 (첫 화면 카드에 쓴다)
THUMB = {"ch00": "fig_bi_flow.png", "ch01": "pbi_01_데이터불러오기_p1.png", "ch02": "ui_pq_sales.png",
         "ch03": "ui_model.png", "ch04": "pbi_04_측정값_p1.png", "ch05": "pbi_05_차트_p1.png",
         "ch06": "pbi_06_슬라이서와서식_p1.png", "ch07": "pbi_07_대시보드완성_p1.png",
         "ch08": "pbi_과제_편의점_풀이_p1.png"}

DOWNLOADS = [  # (보이는 이름, 저장소 안 경로, 설명)
    ("카페_판매.csv", "data/카페_판매.csv", "판매 내역 36만 줄 · 27MB"),
    ("카페_메뉴.xlsx", "data/카페_메뉴.xlsx", "메뉴 22개"),
    ("카페_매장.xlsx", "data/카페_매장.xlsx", "매장 3곳"),
    ("하루커피_테마.json", "data/하루커피_테마.json", "6장 색 테마"),
    ("편의점_판매.csv", "data/편의점_판매.csv", "8장 과제 · 8MB"),
    ("편의점_상품.xlsx", "data/편의점_상품.xlsx", "8장 과제"),
    ("편의점_점포.xlsx", "data/편의점_점포.xlsx", "8장 과제"),
]
PBIX = ["01_데이터불러오기", "02_데이터정리", "03_관계와날짜", "04_측정값", "05_차트",
        "06_슬라이서와서식", "07_대시보드완성", "과제_편의점_풀이"]


def raw(path):
    """저장소 안 파일의 내려받기 주소 (한글 파일 이름은 % 로 바꾼다)"""
    return RAW + quote(path)


# ------------------------------------------------------------------ 마크다운 → HTML
def prepare(md):
    """python-markdown 이 잘 읽도록 원고를 조금 다듬는다 (원고 파일은 바꾸지 않는다)"""
    lines, out = md.split("\n"), []
    for i, line in enumerate(lines):
        line = re.sub(r"^ {2,3}([-*] )", r"    \1", line)               # 2~3칸 들여 쓴 하위 목록 → 4칸
        is_item = re.match(r"^(\d+\.|[-*]) ", line)
        prev = out[-1] if out else ""
        if is_item and prev.strip() and not re.match(r"^\s*(\d+\.|[-*]) ", prev):
            out.append("")                                               # 문단 바로 뒤 목록 앞에 빈 줄
        out.append(line)
    return "\n".join(out)


def to_html(md):
    conv = markdown.Markdown(extensions=["tables", "fenced_code", "toc", "sane_lists"],
                             extension_configs={"toc": {"slugify": slug}}, lazy_ol=False)
    body = conv.convert(prepare(md))
    # 그림: <p><img alt="설명"></p> → 설명이 붙은 figure, 누르면 원본 크기로 열기
    body = re.sub(r'<p><img alt="([^"]*)" src="([^"]+)" ?/?></p>',
                  lambda m: (f'<figure><a href="{m[2]}" target="_blank"><img src="{m[2]}" alt="{m[1]}" loading="lazy"></a>'
                             f'<figcaption>{m[1]}</figcaption></figure>'), body)
    # 인용 상자(>)는 "도움말 상자" 로 보이게
    body = body.replace("<blockquote>", '<blockquote class="note">')
    # 표는 가로로 넘치면 표만 옆으로 밀리게 감싼다
    body = body.replace("<table>", '<div class="tbl"><table>').replace("</table>", "</table></div>")
    return body, conv.toc_tokens


def slug(value, sep):
    """제목 → 주소 조각 (한글은 그대로 두고 빈칸만 - 로)"""
    return re.sub(r"[\s]+", "-", re.sub(r"[^\w\s가-힣-]", "", value)).strip("-").lower()


# ------------------------------------------------------------------ 공통 틀
CSS = """
:root{--ink:#1f2328;--sub:#59636e;--line:#e4e1d8;--bg:#faf8f3;--card:#fff;--navy:#1d2b4f;--yellow:#f2c811;
--brown:#6b4a2f;--note:#fff8dc;--noteline:#e9c84a;--code:#f4f1ea;--link:#1a5fb4}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--ink:#e8e6e1;--sub:#a7a39b;--line:#38352f;--bg:#191815;
--card:#23221e;--navy:#0f1730;--note:#2c2813;--noteline:#8a7420;--code:#2b2a25;--link:#8ab4f8}}
*{box-sizing:border-box}html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.75 "Pretendard","Malgun Gothic","Apple SD Gothic Neo",sans-serif;word-break:keep-all}
a{color:var(--link)}img{max-width:100%;height:auto}
header.top{background:var(--navy);color:#fff;position:sticky;top:0;z-index:5}
header.top .in{max-width:1100px;margin:0 auto;padding:10px 16px;display:flex;gap:12px;align-items:center}
header.top a{color:#fff;text-decoration:none}.logo{font-weight:800;white-space:nowrap}.logo b{color:var(--yellow)}
header.top .sp{flex:1}header.top .btn{font-size:14px;border:1px solid #ffffff55;border-radius:8px;padding:4px 10px;white-space:nowrap}
.wrap{max-width:1100px;margin:0 auto;padding:0 16px}
/* 장 읽기 화면: 넓은 화면에서는 왼쪽 목차 */
.doc{display:grid;grid-template-columns:240px minmax(0,1fr);gap:32px;padding:24px 16px 60px;max-width:1100px;margin:0 auto}
.side{position:sticky;top:64px;align-self:start;max-height:calc(100vh - 80px);overflow:auto;font-size:14px}
.side a{display:block;padding:5px 10px;border-radius:6px;color:var(--ink);text-decoration:none}
.side a.on{background:var(--yellow);color:#222;font-weight:700}.side a:hover:not(.on){background:var(--line)}
article{min-width:0}
article h1{font-size:28px;line-height:1.35;margin:0 0 20px;padding-bottom:12px;border-bottom:4px solid var(--yellow)}
article h2{font-size:22px;margin:40px 0 12px;padding-left:10px;border-left:5px solid var(--brown)}
article h3{font-size:18px;margin:28px 0 8px}
figure{margin:18px 0;text-align:center}figure img{border:1px solid var(--line);border-radius:8px;background:#fff}
figcaption{font-size:13px;color:var(--sub);margin-top:6px}
.tbl{overflow-x:auto;margin:14px 0}table{border-collapse:collapse;width:100%;font-size:15px}
th{background:var(--navy);color:#fff;text-align:left}th,td{padding:7px 10px;border:1px solid var(--line);vertical-align:top}
tr:nth-child(even) td{background:color-mix(in srgb,var(--card) 70%,var(--line))}
blockquote.note{margin:16px 0;padding:10px 16px;background:var(--note);border-left:5px solid var(--noteline);border-radius:6px}
blockquote.note p{margin:6px 0}
code{background:var(--code);padding:1px 5px;border-radius:4px;font-family:Consolas,"D2Coding",monospace;font-size:.92em}
pre{background:var(--code);padding:12px 14px;border-radius:8px;overflow-x:auto;border-left:4px solid var(--yellow)}
pre code{background:none;padding:0;font-size:14px;line-height:1.55}
.pager{display:flex;gap:12px;justify-content:space-between;margin-top:48px;padding-top:20px;border-top:1px solid var(--line)}
.pager a{flex:1;padding:12px 14px;border:1px solid var(--line);border-radius:10px;text-decoration:none;color:var(--ink);background:var(--card)}
.pager a small{display:block;color:var(--sub)}.pager .next{text-align:right}
select.jump{display:none;width:100%;font-size:16px;padding:8px;margin-bottom:12px;border-radius:8px;border:1px solid var(--line);background:var(--card);color:var(--ink)}
@media (max-width:820px){.doc{display:block;padding-top:16px}.side{display:none}select.jump{display:block}article h1{font-size:23px}
header.top .hide-s{display:none}}
/* 첫 화면 */
.hero{background:var(--navy);color:#fff;padding:44px 0 36px}
.hero h1{font-size:40px;margin:0 0 8px;line-height:1.2}.hero h1 b{color:var(--yellow)}
.hero p{margin:6px 0;color:#d9deea;max-width:640px}.hero .cta{display:flex;flex-wrap:wrap;gap:10px;margin-top:22px}
.cta a{padding:11px 18px;border-radius:10px;text-decoration:none;font-weight:700}
.cta .main{background:var(--yellow);color:#222}.cta .sub{border:1px solid #ffffff66;color:#fff}
.hero .shot{margin-top:28px;border-radius:10px;border:1px solid #ffffff33;display:block}
section{padding:36px 0 8px}section h2{font-size:24px;margin:0 0 6px}section .lead{color:var(--sub);margin:0 0 18px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden;text-decoration:none;color:var(--ink);display:flex;flex-direction:column}
.card:hover{border-color:var(--yellow);box-shadow:0 4px 14px #0000001a}
.card .im{aspect-ratio:16/9;overflow:hidden;background:#fff;border-bottom:1px solid var(--line)}
.card .im img{width:100%;height:100%;object-fit:cover;object-position:left top}
.card .tx{padding:12px 14px}.card b{display:block}.card small{color:var(--sub)}
.steps{counter-reset:s;list-style:none;padding:0;display:grid;gap:10px}
.steps li{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px 12px 54px;position:relative}
.steps li:before{counter-increment:s;content:counter(s);position:absolute;left:14px;top:12px;width:28px;height:28px;border-radius:50%;
background:var(--yellow);color:#222;font-weight:800;text-align:center;line-height:28px}
.dl{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:10px}
.dl a{display:block;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 14px;text-decoration:none;color:var(--ink)}
.dl a:hover{border-color:var(--yellow)}.dl small{display:block;color:var(--sub)}
footer{padding:40px 0 50px;color:var(--sub);font-size:14px;text-align:center}
@media (max-width:600px){.hero h1{font-size:30px}.hero{padding-top:30px}}
"""

# 카카오톡 안에서 열면 PDF·내려받기가 잘 안 되므로 외부 브라우저로 다시 연다
KAKAO = """<script>
if(/KAKAOTALK/i.test(navigator.userAgent)){location.href='kakaotalk://web/openExternal?url='+encodeURIComponent(location.href);}
</script>"""


def page(title, body, rel="", desc="Power BI 를 처음 배우는 학생을 위한 기초 과정 (약 20시간)"):
    """모든 화면이 함께 쓰는 HTML 틀. rel 은 docs 폴더까지의 상대 경로 ('' 또는 '../')"""
    og = SITE + "images/pbi_07_" + quote("대시보드완성") + "_p1.png"
    return f"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)}</title>
<meta name="description" content="{desc}">
<meta property="og:title" content="{html.escape(title)}"><meta property="og:description" content="{desc}">
<meta property="og:image" content="{og}"><meta property="og:type" content="website">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect x='4' y='14' width='6' height='14' fill='%23f2c811'/%3E%3Crect x='13' y='8' width='6' height='20' fill='%23e0a800'/%3E%3Crect x='22' y='3' width='6' height='25' fill='%23b07d00'/%3E%3C/svg%3E">
{KAKAO}
<style>{CSS}</style></head><body>
<header class="top"><div class="in">
<a class="logo" href="{rel}index.html">Power BI <b>기초</b></a><span class="sp"></span>
<a class="btn hide-s" href="{rel}index.html#chapters">목차</a>
<a class="btn" href="{rel}{quote(PDF)}" target="_blank">PDF</a>
<a class="btn" href="{rel}index.html#download">내려받기</a>
</div></header>
{body}
</body></html>"""


# ------------------------------------------------------------------ 장별 화면
def build_chapters():
    for i, (src, name, short, _) in enumerate(CHAPTERS):
        md = (SRC / src).read_text(encoding="utf-8")
        body, _toc = to_html(md)
        title = re.search(r"^# (.+)$", md, re.M)[1]
        side = "".join(f'<a href="{n}.html"{" class=on" if n == name else ""}>{html.escape(s)}</a>'
                       for _, n, s, _ in CHAPTERS)
        jump = "".join(f'<option value="{n}.html"{" selected" if n == name else ""}>{html.escape(s)}</option>'
                       for _, n, s, _ in CHAPTERS)
        prev = CHAPTERS[i - 1] if i > 0 else None
        nxt = CHAPTERS[i + 1] if i + 1 < len(CHAPTERS) else None
        pager = '<nav class="pager">'
        pager += f'<a href="{prev[1]}.html"><small>← 이전</small>{prev[2]}</a>' if prev else "<span></span>"
        pager += f'<a class="next" href="{nxt[1]}.html"><small>다음 →</small>{nxt[2]}</a>' if nxt else "<span></span>"
        pager += "</nav>"
        html_body = f"""<div class="doc">
<nav class="side">{side}</nav>
<article>
<select class="jump" onchange="location.href=this.value" aria-label="장 이동">{jump}</select>
{body}
{pager}
</article></div>
<footer>Power BI 기초 학습 교안 · 강사 김대철 · <a href="https://github.com/{REPO}">GitHub</a></footer>"""
        (SRC / f"{name}.html").write_text(page(f"{title} - Power BI 기초", html_body, "../"), encoding="utf-8")
        print("  장:", f"basic/{name}.html")


# ------------------------------------------------------------------ 첫 화면
def build_index():
    cards = ""
    for _, name, short, hours in CHAPTERS[:9]:
        img = quote(THUMB[name])
        cards += (f'<a class="card" href="basic/{name}.html"><div class="im"><img src="images/{img}" alt="" loading="lazy"></div>'
                  f'<div class="tx"><b>{html.escape(short)}</b><small>{hours}</small></div></a>')
    apps = "".join(f'<a class="card" href="basic/{n}.html"><div class="tx"><b>{html.escape(s)}</b></div></a>'
                   for _, n, s, _ in CHAPTERS[9:])
    data = "".join(f'<a href="{raw(p)}"><b>{html.escape(t)}</b><small>{d}</small></a>' for t, p, d in DOWNLOADS)
    pbix = "".join(f'<a href="{raw("pbix/" + n + ".pbix")}"><b>{html.escape(n)}.pbix</b>'
                   f'<small>{"8장 과제 풀이" if n.startswith("과제") else n[:2] + "장 완성 파일"}</small></a>' for n in PBIX)
    hero = quote("pbi_07_대시보드완성_p1.png")
    body = f"""<div class="hero"><div class="wrap">
<h1>Power BI <b>기초</b></h1>
<p>가상의 카페 <b>하루커피</b>의 3년치 매출 데이터를 정리하고 분석해 매출 대시보드를 완성하는 약 20시간 과정입니다.
처음 배우는 학생도 혼자 따라 할 수 있도록 모든 단계를 실제 화면과 함께 설명합니다.</p>
<div class="cta">
<a class="main" href="basic/ch00.html">교안 읽기 시작 →</a>
<a class="sub" href="{quote(PDF)}" target="_blank">PDF 교안 (49쪽)</a>
<a class="sub" href="{ZIP}">실습 파일 전체 받기 (ZIP)</a>
</div>
<img class="shot" src="images/{hero}" alt="7장에서 완성하는 하루커피 매출 대시보드">
</div></div>

<div class="wrap">
<section id="start"><h2>시작하기</h2><p class="lead">Windows PC 와 Power BI Desktop(무료)만 있으면 됩니다.</p>
<ol class="steps">
<li><b>Power BI Desktop 설치</b> - Microsoft Store 에서 "Power BI Desktop" 을 검색해 설치합니다.</li>
<li><b>실습 파일 받기</b> - <a href="{ZIP}">ZIP 파일</a>을 받아 <code>C:\\study</code> 에 풀고, 폴더 이름을 <code>power-bi-tutorial</code> 로 바꿉니다.
완성 파일이 <code>C:\\study\\power-bi-tutorial\\data</code> 의 데이터를 읽도록 만들어져 있습니다.</li>
<li><b>0장부터 차례로</b> - 각 장은 "학습 목표 → 따라 하기 → 정리 → 확인 문제" 순서입니다. 답은 부록 C 에 있습니다.</li>
</ol></section>

<section id="chapters"><h2>과정 목차</h2><p class="lead">1~7장은 카페 데이터로 대시보드를 완성하고, 8장 과제에서 편의점 데이터로 혼자 만들어 봅니다.</p>
<div class="grid">{cards}</div>
<div class="grid" style="margin-top:16px">{apps}</div></section>

<section id="download"><h2>내려받기</h2>
<p class="lead">한 번에 받으려면 <a href="{ZIP}">전체 ZIP</a> 이 편합니다. 파일을 하나씩 받을 수도 있습니다.</p>
<h3>실습 데이터</h3><div class="dl">{data}</div>
<h3>장별 완성 파일 (.pbix)</h3><p class="lead">막히면 그 장의 완성 파일을 열어 비교해 보세요.</p><div class="dl">{pbix}</div>
</section>
</div>
<footer>데이터는 모두 교육용으로 만든 가상 데이터입니다 · 강사 김대철 · <a href="https://github.com/{REPO}">GitHub 저장소</a></footer>"""
    (DOCS / "index.html").write_text(page("Power BI 기초 학습 교안", body), encoding="utf-8")
    print("  첫 화면: index.html")


if __name__ == "__main__":
    build_chapters()
    build_index()
    (DOCS / ".nojekyll").write_text("", encoding="utf-8")    # Jekyll 변환 끄기 (파일을 있는 그대로 공개)
