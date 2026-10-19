"""Build the static, file://-compatible guide from editable editorial HTML."""
import html
import hashlib
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FULL_SOURCE = ROOT.parent / 'Claude outputs' / 'ludovskaya-gubinskaya-igarskaya-svody.md'


def full_source_material():
    """Render every source paragraph, including superseded interpretations."""
    import mistune

    source_bytes = FULL_SOURCE.read_bytes()
    source = source_bytes.decode('utf-8')
    headings = re.findall(r'^(#{1,6})\s+(.+?)\s*$', source, re.MULTILINE)
    items = [(len(marks), title, 'source-' + hashlib.sha1(title.encode('utf-8')).hexdigest()[:12])
             for marks, title in headings]
    if hasattr(mistune, 'create_markdown'):
        rendered = mistune.create_markdown(escape=True, plugins=['table'])(source)
    else:
        rendered = mistune.Markdown(renderer=mistune.Renderer(escape=True))(source)
    heading_index = iter(items)
    def heading_anchor(match):
        level, title, anchor = next(heading_index)
        assert int(match.group(1)) == level
        return '<h{0} id="{1}">{2}</h{0}>'.format(level, anchor, match.group(2))
    assert len(re.findall(r'<h[1-6]>', rendered)) == len(items)
    rendered = re.sub(r'<h([1-6])>(.*?)</h\1>', heading_anchor, rendered, flags=re.DOTALL)
    rendered = rendered.replace('<table>', '<div class="table-wrap"><table>').replace('</table>', '</table></div>')
    chunks = re.split(r'(?=<h2 id=)', rendered)
    # Headings retain their own anchors; section wrappers give scroll tracking
    # an independent target and never duplicate heading IDs.
    rendered = ''.join('<section id="block-' + str(i) + '">' + chunk + '</section>'
                       for i, chunk in enumerate(chunks))
    notice = '''<section id="full-text"><p class="kicker">Полные материалы к путеводителю</p>
<h1>Исходная сводка без сокращений</h1><p>Ниже сохранён весь текст предоставленного файла: описания подразделений, разрезы, мощности, контакты, петрография, геохимия, методика, аргументы и последовательные уточнения.</p>
<p>Это рабочая авторская сводка. В ней рядом остаются первоначальные выводы и их позднейшие пересмотры, а также расчёты с настройками, отличающимися от текущих настроек портала.</p>
<div class="actions"><a href="guide.html">К обзорному путеводителю</a><a href="guide_source.md" download>Скачать исходный текст Markdown</a></div></section>'''
    nav = '<a href="guide.html">← К обзору</a><a href="#full-text">О полном тексте</a>' + ''.join(
        '<a class="toc-level-{0}" href="#{1}">{2}</a>'.format(level, anchor, html.escape(title))
        for level, title, anchor in items if level >= 2)
    (ROOT / 'guide_source.md').write_bytes(source_bytes)
    return notice + rendered, nav, items


def build_guide():
    content = (ROOT / 'guide_content.html').read_text(encoding='utf-8')
    full_content, full_nav, source_headings = full_source_material()
    topics = {
        'schemes': ('положение Губинской', 'от самостоятельной', 'Что это значит', 'Схема Кочнева', 'Госгеолкарте-200', 'Госгеолкарта-1000', 'Фундамент'),
        'ludovskaya': ('Лудовская толща', 'Медвежинская свита', 'Старомостовская серия'),
        'gubinskaya': ('Губинская толща', 'положение Губинской', 'SM26-07 —'),
        'igarskaya': ('Петрография игарской', 'фациальная зональность', 'Вывод Гёлецяна', 'Яшмоиды', 'SM25-16 —', 'Методика анализа', 'Шлифы (', 'SM26-03', 'Полная расшифровка'),
        'intrusions': ('Секущие дайки', 'Магматизм', 'Тело SM25-07', 'Руч. Спилитовый:', 'Риолиты', 'Геохимия по группам'),
    }
    def link_full_details(match):
        section_id = match.group(1)
        links = [(title, anchor) for level, title, anchor in source_headings
                 if any(token in title for token in topics.get(section_id, ()))]
        if not links:
            return match.group(0)
        details = '<details class="full-source-links"><summary>Подробные материалы из исходной сводки</summary><ul>'
        details += ''.join('<li><a href="guide-full.html#' + anchor + '">' + html.escape(title) + '</a></li>'
                           for title, anchor in links)
        details += '</ul></details>'
        return match.group(0).replace('</section>', details + '</section>')
    content = re.sub(r'<section id="([^"]+)"[^>]*>.*?</section>', link_full_details, content, flags=re.DOTALL)
    results = json.loads((ROOT / 'upb' / 'portal_data.json').read_text(encoding='utf-8'))
    def result_block(match):
        sid = match.group(1)
        data = results.get(sid, {})
        summary = html.escape(data.get('summary', 'Результат доступен в карточке образца'))
        return '<div class="result"><span>Текущий сводный расчёт портала</span><p><b>' + html.escape(sid) + '</b> · ' + summary + '</p><small>Числа относятся к настройкам сводного расчёта; интерактивные фильтры задаются отдельно.</small></div>'
    content = re.sub(r'<div class="result" data-upb="([^"]+)"></div>', result_block, content)
    sections = re.findall(r'<section id="([^"]+)" data-title="([^"]+)"', content)
    nav = ''.join('<a href="#' + sid + '">' + html.escape(title) + '</a>' for sid, title in sections)
    nav += '<a href="guide-full.html">Полная сводка →</a>'
    template = '''<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Путеводитель по геологии Игарки</title><meta name="description" content="Лудовская, губинская и игарская: геологические разрезы, образцы, источники и геохронология.">
<style>
:root{--ink:#172b24;--muted:#61726a;--accent:#315f46;--line:#dce3dd;--paper:#fff;--bg:#f3f5f1}*{box-sizing:border-box}html{scroll-behavior:smooth;scroll-padding-top:90px}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.65 -apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif}a{color:var(--accent);text-underline-offset:3px}a:focus-visible,summary:focus-visible{outline:3px solid #d39546;outline-offset:3px}header{position:sticky;top:0;z-index:3;background:#fbfcfaee;backdrop-filter:blur(12px);border-bottom:1px solid var(--line);padding:13px 30px;display:flex;align-items:center;justify-content:space-between;gap:20px}header a{text-decoration:none;font-weight:650}header span{font-size:13px;color:var(--muted)}.layout{max-width:1370px;margin:auto;display:grid;grid-template-columns:245px minmax(0,1fr);gap:40px;padding:34px 30px 80px}nav{position:sticky;top:90px;align-self:start;display:grid;gap:5px}nav a{padding:10px 13px;text-decoration:none;color:var(--muted);font-size:14px;border-radius:8px;border-left:3px solid transparent}nav a:hover,nav a.active{background:#e7eee6;color:var(--accent);border-left-color:var(--accent)}main{min-width:0;max-width:940px}section{padding:32px 40px 40px;background:var(--paper);border:1px solid var(--line);border-radius:12px;margin-bottom:26px;scroll-margin-top:5px}h1{font-size:clamp(29px,3vw,42px);line-height:1.15;letter-spacing:-.025em;margin:8px 0 20px;font-weight:650}h2{font-size:22px;line-height:1.3;margin:32px 0 12px;font-weight:650}p{margin:12px 0}.kicker{font-size:11px;font-weight:700;color:var(--accent);letter-spacing:.11em;text-transform:uppercase}.lead{font-size:19px;line-height:1.55;color:#4f6359}.source,figcaption{font-size:12px;color:var(--muted);line-height:1.6}.source{margin-top:28px;border-top:1px solid var(--line);padding-top:16px}.tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:26px 0}.tile{padding:21px 17px;text-decoration:none;border-radius:9px;border:1px solid var(--line);display:flex;flex-direction:column;gap:10px}.tile strong{font-size:19px}.tile span{font-size:13px;line-height:1.5;color:#485d51}.ludy{background:#f0e9dc}.gub{background:#fae3dc}.igar{background:#e5efdd}.actions{display:flex;flex-wrap:wrap;gap:9px;margin:22px 0}.actions a{font-size:13px;line-height:1.4;text-decoration:none;padding:10px 13px;border:1px solid #c4d4c8;border-radius:7px;background:#f3f8f2}.actions a:hover{background:#e6f0e5}.question{padding:18px 22px;border-left:3px solid #ba884a;background:#fcf8ee;margin:24px 0}.question p{font-size:15px;margin-bottom:0}.result{padding:18px 21px;background:#eef5f0;border:1px solid #d4e3d8;border-radius:8px;margin:22px 0}.result span{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}.result small{font-size:12px;color:var(--muted)}.result p{margin:7px 0}figure{margin:24px 0}figure img{display:block;width:100%;height:auto;border:1px solid var(--line);border-radius:6px}figcaption{margin-top:8px}.table-wrap{overflow-x:auto;margin:24px 0}table{border-collapse:collapse;font-size:13px;line-height:1.55;width:100%;min-width:620px}th,td{text-align:left;vertical-align:top;padding:13px 12px;border-bottom:1px solid var(--line)}th{background:#f0f4ef;color:var(--accent)}details{margin:20px 0;padding:15px 18px;background:#f5f7f3;border:1px solid var(--line);border-radius:7px}summary{cursor:pointer;font-weight:600}li{margin:9px 0}.bibliography{padding-left:22px}.bibliography li{padding-bottom:10px}footer{font-size:12px;color:var(--muted);padding:0 8px}
@media(max-width:960px){.layout{grid-template-columns:185px minmax(0,1fr);gap:20px;padding:22px 18px}section{padding:26px}.tiles{grid-template-columns:1fr}.tile{gap:3px}}
@media(max-width:640px){header{padding:13px 16px}header span{display:none}.layout{display:block;padding:14px 12px 45px}nav{position:relative;top:auto;display:flex;overflow-x:auto;gap:6px;margin-bottom:16px;padding-bottom:8px}nav a{white-space:nowrap;flex-shrink:0;padding:8px 10px;font-size:12px}section{padding:23px 19px;border-radius:9px}h2{font-size:20px}.lead{font-size:17px}.actions a{max-width:100%}html{scroll-padding-top:70px}}
@media print{header,nav,.actions{display:none}.layout{display:block;padding:0}body{background:white}section{break-inside:auto;border:0;padding:12px 0;margin:0}details{display:block}main{max-width:none}}
</style></head><body><header><a href="index.html">← Игарка · карта и образцы</a><span>Путеводитель по геологии района</span></header><div class="layout"><nav aria-label="Разделы путеводителя">__NAV__</nav><main>__CONTENT__<footer>Материалы проекта «Игарка» · Малышев, Пасенко · Рабочая научная версия</footer></main></div>
<script>
const guideLinks=[...document.querySelectorAll('nav a')];
function markSection(id){guideLinks.forEach(a=>{const active=a.getAttribute('href')==='#'+id;a.classList.toggle('active',active);if(active)a.setAttribute('aria-current','location');else a.removeAttribute('aria-current');});}
if('IntersectionObserver' in window){const observer=new IntersectionObserver(entries=>{const visible=entries.filter(e=>e.isIntersecting).sort((a,b)=>a.boundingClientRect.top-b.boundingClientRect.top);if(visible.length)markSection(visible[0].target.id);},{rootMargin:'-80px 0px -50% 0px',threshold:0});document.querySelectorAll('main section').forEach(s=>observer.observe(s));}
window.addEventListener('hashchange',()=>markSection(location.hash.slice(1)||'overview'));markSection(location.hash.slice(1)||'overview');
</script></body></html>'''
    extra_style = '''<style>
@media(min-width:641px){nav{max-height:calc(100vh - 115px);overflow-y:auto}}
nav .toc-level-3{padding-left:25px;font-size:12px}h1,h2,h3,h4{overflow-wrap:anywhere;scroll-margin-top:90px}
code{font-size:.88em;overflow-wrap:anywhere;background:#f0f3ed;padding:1px 4px;border-radius:3px}
pre{overflow:auto;padding:15px;background:#f0f3ed}section p,section li{overflow-wrap:anywhere}
.full-source-links li{font-size:14px}.full-source-links summary{color:var(--accent)}
@media print{details{content-visibility:visible}details>summary{display:none}}
</style>'''
    template = template.replace('</head>', extra_style + '</head>')
    output = ROOT / 'guide.html'
    output.write_text(template.replace('__NAV__', nav).replace('__CONTENT__', content), encoding='utf-8')
    full_template = template.replace('<title>Путеводитель по геологии Игарки</title>', '<title>Игарка — полная геологическая сводка</title>')
    full_template = full_template.replace('document.querySelectorAll(\'main section\')', 'document.querySelectorAll(\'main h2[id]\')')
    (ROOT / 'guide-full.html').write_text(full_template.replace('__NAV__', full_nav).replace('__CONTENT__', full_content), encoding='utf-8')
    return output


if __name__ == '__main__':
    print(build_guide())
