# -*- coding: utf-8 -*-
"""
Простая школьная презентация про М.И. Кутузова.
Формат: обычная школьная работа 8 класса.
Автор в презентации: Казаков Дмитрий, 8 "А" класс.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ---------- Цвета (скромные, "школьные") ----------
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BLACK  = RGBColor(0x00, 0x00, 0x00)
DARK   = RGBColor(0x1F, 0x1F, 0x1F)
BLUE   = RGBColor(0x1F, 0x49, 0x7D)   # классический "офисный" синий
LIGHT  = RGBColor(0xEA, 0xF1, 0xF8)
GRAY   = RGBColor(0x55, 0x55, 0x55)
RULE   = RGBColor(0xCC, 0xD5, 0xE0)

F_BODY = "Calibri"      # по умолчанию в Office
F_HEAD = "Calibri"
F_SERIF = "Times New Roman"

# ---------- Документ ----------
prs = Presentation()
prs.slide_width  = Inches(10)      # 4:3, как в школе обычно
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ==========================================================
#   УТИЛИТЫ
# ==========================================================

def add_bg(slide, color=WHITE):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background(); s.shadow.inherit = False
    return s

def add_text(slide, x, y, w, h, text, *,
             size=18, bold=False, color=BLACK, font=F_BODY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.15, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        f = r.font
        f.name = font
        f.size = Pt(size)
        f.bold = bold
        f.italic = italic
        f.color.rgb = color
    return tb

def add_bullets(slide, x, y, w, h, bullets, *,
                size=20, color=BLACK, font=F_BODY, line_spacing=1.25):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.word_wrap = True
    for i, item in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        # отступ уровня 0
        pPr = p._pPr
        if pPr is None:
            pPr = p._p.get_or_add_pPr()
        # настраиваем буллет "•"
        for ch in list(pPr):
            if ch.tag in (qn("a:buChar"), qn("a:buAutoNum"), qn("a:buNone")):
                pPr.remove(ch)
        buFont = etree.SubElement(pPr, qn("a:buFont"))
        buFont.set("typeface", "Arial")
        buChar = etree.SubElement(pPr, qn("a:buChar"))
        buChar.set("char", "•")
        pPr.set("marL", "285750")   # ~0.3"
        pPr.set("indent", "-285750")

        r = p.add_run()
        r.text = item
        f = r.font
        f.name = font
        f.size = Pt(size)
        f.color.rgb = color
    return tb

def add_rect(slide, x, y, w, h, *, fill=None, line=None, line_w=0.75):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    return s


def header(slide, title_text, number, total):
    # тонкая синяя полоса сверху
    add_rect(slide, 0, 0, SW, Inches(0.12), fill=BLUE)
    # заголовок слайда
    add_text(slide, Inches(0.6), Inches(0.35), Inches(8.8), Inches(0.8),
             title_text, size=32, bold=True, color=BLUE, font=F_HEAD)
    # разделитель
    add_rect(slide, Inches(0.6), Inches(1.2), Inches(8.8), Emu(12700),
             fill=RULE)
    # футер
    add_text(slide, Inches(0.6), Inches(7.05), Inches(6), Inches(0.3),
             "Кутузов М.И. — великий русский полководец",
             size=10, color=GRAY, italic=True, font=F_BODY)
    add_text(slide, Inches(8.4), Inches(7.05), Inches(1.2), Inches(0.3),
             f"{number} / {total}",
             size=10, color=GRAY, font=F_BODY, align=PP_ALIGN.RIGHT)


# ==========================================================
#   ПЕРЕХОДЫ (простые)
# ==========================================================
def set_transition(slide, kind="fade", duration_ms=700):
    nsmap_p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    trans = etree.SubElement(sld, qn("p:transition"),
                             attrib={"spd": "med", "advClick": "1"})
    trans.set("{%s}dur" % nsmap_p14, str(duration_ms))
    xml = {
        "fade": '<p:fade xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
        "push": '<p:push xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" dir="l"/>',
        "wipe": '<p:wipe xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" dir="l"/>',
    }[kind]
    trans.append(etree.fromstring(xml))


# ==========================================================
#   СЛАЙДЫ
# ==========================================================

TOTAL = 13


def slide_title():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)

    # шапка — "школьная"
    add_text(s, Inches(0.6), Inches(0.5), Inches(8.8), Inches(0.4),
             "Муниципальное бюджетное общеобразовательное учреждение",
             size=14, color=GRAY, font=F_BODY, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.6), Inches(0.9), Inches(8.8), Inches(0.4),
             "средняя общеобразовательная школа",
             size=14, color=GRAY, font=F_BODY, align=PP_ALIGN.CENTER)

    # двойная линия
    add_rect(s, Inches(2.0), Inches(1.9), Inches(6.0), Emu(25400), fill=BLUE)
    add_rect(s, Inches(2.0), Inches(1.97), Inches(6.0), Emu(12700), fill=BLUE)

    add_text(s, Inches(0.6), Inches(2.2), Inches(8.8), Inches(0.4),
             "Презентация по истории", size=18, color=DARK, font=F_BODY,
             align=PP_ALIGN.CENTER, italic=True)

    # тема
    add_text(s, Inches(0.4), Inches(2.9), Inches(9.2), Inches(0.6),
             "на тему:", size=18, color=DARK, font=F_BODY, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.4), Inches(3.4), Inches(9.2), Inches(1.6),
             "«Кутузов Михаил Илларионович —\nвеликий русский полководец»",
             size=32, bold=True, color=BLUE, font=F_HEAD,
             align=PP_ALIGN.CENTER, line_spacing=1.15)

    add_rect(s, Inches(2.0), Inches(5.0), Inches(6.0), Emu(12700), fill=BLUE)

    # автор
    add_text(s, Inches(5.0), Inches(5.3), Inches(4.6), Inches(0.4),
             "Выполнил:", size=16, color=DARK, font=F_BODY)
    add_text(s, Inches(5.0), Inches(5.65), Inches(4.6), Inches(0.4),
             "ученик 8 «А» класса", size=16, color=DARK, font=F_BODY)
    add_text(s, Inches(5.0), Inches(6.0), Inches(4.6), Inches(0.5),
             "Казаков Дмитрий",
             size=20, bold=True, color=DARK, font=F_BODY)
    add_text(s, Inches(5.0), Inches(6.55), Inches(4.6), Inches(0.4),
             "Руководитель: _______________",
             size=14, color=GRAY, font=F_BODY)

    # город-год
    add_text(s, Inches(0.4), Inches(7.0), Inches(9.2), Inches(0.4),
             "2026 г.", size=14, color=DARK, font=F_BODY, align=PP_ALIGN.CENTER)

    set_transition(s, "fade", 600)


def slide_plan():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Содержание", 2, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.6), Inches(8.6), Inches(5.4), [
        "Краткая биография",
        "Детство и образование",
        "Начало военной службы и ранения",
        "Войны с Турцией и Европейские кампании",
        "Отечественная война 1812 года",
        "Бородинское сражение",
        "Совет в Филях. Оставление Москвы",
        "Изгнание Наполеона из России",
        "Итоги и значение",
        "Список использованной литературы",
    ], size=22)
    set_transition(s, "fade", 600)


def slide_bio():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Краткая биография", 3, TOTAL)

    # карточка слева
    add_rect(s, Inches(0.6), Inches(1.5), Inches(3.8), Inches(5.0),
             fill=LIGHT, line=RULE, line_w=0.75)
    # "фото" — заглушка
    add_rect(s, Inches(1.0), Inches(1.8), Inches(3.0), Inches(3.0),
             fill=RULE, line=BLUE, line_w=1.2)
    add_text(s, Inches(1.0), Inches(1.8), Inches(3.0), Inches(3.0),
             "[ портрет\nМ.И. Кутузова ]",
             size=14, color=GRAY, font=F_BODY, italic=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, Inches(0.8), Inches(4.9), Inches(3.4), Inches(0.4),
             "Михаил Илларионович",
             size=16, bold=True, color=BLUE, font=F_BODY,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.8), Inches(5.25), Inches(3.4), Inches(0.4),
             "Голенищев-Кутузов",
             size=16, bold=True, color=BLUE, font=F_BODY,
             align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.8), Inches(5.75), Inches(3.4), Inches(0.4),
             "(1745 – 1813)",
             size=14, color=GRAY, font=F_BODY, align=PP_ALIGN.CENTER,
             italic=True)

    # справа — факты
    add_bullets(s, Inches(4.8), Inches(1.55), Inches(4.8), Inches(5.3), [
        "Годы жизни: 1745 – 1813.",
        "Родился в Санкт-Петербурге.",
        "Русский полководец, генерал-фельдмаршал.",
        "Светлейший князь Смоленский (с 1812 г.).",
        "Главнокомандующий русской армией в Отечественной войне 1812 года.",
        "Полный кавалер ордена Святого Георгия.",
        "Ученик и соратник А.В. Суворова.",
    ], size=17)

    set_transition(s, "push", 600)


def slide_childhood():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Детство и образование", 4, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "Родился 5 (16) сентября 1745 года в семье военного инженера, генерал-поручика Иллариона Матвеевича Кутузова.",
        "С детства отличался любознательностью, хорошо учился, знал несколько иностранных языков.",
        "В 1759 году поступил в Дворянскую артиллерийскую и инженерную школу в Санкт-Петербурге.",
        "Окончил школу с отличием и был оставлен преподавателем математики.",
        "В 16 лет получил первый офицерский чин — прапорщик.",
    ], size=20)
    set_transition(s, "push", 600)


def slide_service():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Начало военной службы. Ранения", 5, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "С 1770-х годов участвовал в русско-турецких войнах под командованием П.А. Румянцева и А.В. Суворова.",
        "В 1774 году был тяжело ранен в голову у деревни Шумы в Крыму — пуля прошла навылет у виска.",
        "В 1788 году при осаде Очакова получил второе ранение в голову почти в то же место. Врачи считали его безнадёжным, но Кутузов выжил.",
        "После ранений потерял глаз, однако вернулся в строй и продолжил военную карьеру.",
        "За подвиги был награждён многими орденами, в том числе орденом Святого Георгия.",
    ], size=19)
    set_transition(s, "push", 600)


def slide_europe():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Войны с Турцией и Европейские кампании", 6, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "Участвовал в штурме Измаила в 1790 году. А.В. Суворов писал о Кутузове: «Он шёл у меня на левом крыле, но был моей правой рукой».",
        "С 1792 года — на дипломатической службе, был послом в Турции.",
        "В 1805 году возглавил русскую армию в войне с Наполеоном. Совершил знаменитый марш-отступление от Браунау до Ольмюца.",
        "В сражении при Аустерлице (1805) армия потерпела поражение — план боя был навязан Кутузову императором Александром I.",
        "В 1811–1812 годах снова воевал с Турцией. Разгромил турецкую армию под Рущуком и заключил выгодный для России Бухарестский мир.",
    ], size=17)
    set_transition(s, "push", 600)


def slide_1812():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Отечественная война 1812 года", 7, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "12 (24) июня 1812 года армия Наполеона (около 600 тысяч человек) перешла реку Неман и вторглась в Россию.",
        "Русские войска были вынуждены отступать: силы противника были значительно больше.",
        "В обществе нарастало недовольство отступлением. Народ требовал назначения главнокомандующим русского, а не иностранца.",
        "8 (20) августа 1812 года Александр I назначил М.И. Кутузова главнокомандующим всеми русскими армиями.",
        "Кутузов прибыл к войскам со словами: «Ну как можно отступать с такими молодцами!»",
    ], size=18)
    set_transition(s, "push", 600)


def slide_borodino():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Бородинское сражение", 8, TOTAL)

    add_text(s, Inches(0.8), Inches(1.45), Inches(8.6), Inches(0.5),
             "26 августа (7 сентября) 1812 года",
             size=18, italic=True, color=GRAY, font=F_BODY)

    add_bullets(s, Inches(0.8), Inches(2.05), Inches(8.6), Inches(4.8), [
        "Крупнейшее сражение Отечественной войны 1812 года.",
        "Произошло в 125 км западнее Москвы, у села Бородино.",
        "Русская армия — около 120 тысяч человек, французская — около 135 тысяч.",
        "Бой длился около 12 часов. Потери с обеих сторон были огромные — десятки тысяч человек.",
        "Ни одна из сторон не добилась решающей победы, но стратегически сражение выиграл Кутузов: армия Наполеона не была разгромлена, но потеряла наступательную силу.",
        "За Бородино Кутузов получил чин генерал-фельдмаршала.",
    ], size=17)
    set_transition(s, "wipe", 700)


def slide_fili():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Совет в Филях. Оставление Москвы", 9, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "1 (13) сентября 1812 года в подмосковной деревне Фили состоялся военный совет.",
        "Кутузов принял тяжёлое решение — оставить Москву без боя, чтобы сохранить армию.",
        "Его знаменитые слова: «С потерей Москвы ещё не потеряна Россия... С потерей же армии Россия потеряна».",
        "2 (14) сентября французская армия вошла в Москву. В городе начались пожары.",
        "Наполеон напрасно ждал в Москве делегации с просьбой о мире — её не было.",
    ], size=19)
    set_transition(s, "push", 600)


def slide_expulsion():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Изгнание Наполеона из России", 10, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "Кутузов совершил знаменитый Тарутинский марш-манёвр: перевёл армию на юг, закрыв французам дорогу в хлебные губернии.",
        "Сражение у Малоярославца (12 октября 1812 г.) заставило Наполеона отступать по разорённой Смоленской дороге.",
        "Началось массовое дезертирство и гибель французской армии от голода и холода.",
        "Действия партизан (Д. Давыдов, А. Сеславин и др.) наносили огромный урон противнику.",
        "К декабрю 1812 года от «Великой армии» Наполеона осталось менее 10% — около 30 тысяч человек.",
    ], size=18)
    set_transition(s, "push", 600)


def slide_result():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Итоги и значение", 11, TOTAL)

    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(3.8), [
        "М.И. Кутузов — один из величайших полководцев в истории России.",
        "Спас страну от порабощения, сохранил русскую армию и разгромил «Великую армию» Наполеона.",
        "За победу в 1812 году получил титул светлейшего князя Смоленского.",
        "Умер 16 (28) апреля 1813 года в городе Бунцлау (Силезия) во время Заграничного похода русской армии.",
    ], size=18)

    # цитата-рамка
    add_rect(s, Inches(0.8), Inches(5.4), Inches(8.4), Inches(1.3),
             fill=LIGHT, line=BLUE, line_w=1)
    add_text(s, Inches(1.0), Inches(5.5), Inches(8.0), Inches(0.4),
             "Его имя носят улицы и проспекты, корабли и военные академии.",
             size=14, color=DARK, font=F_BODY, italic=True)
    add_text(s, Inches(1.0), Inches(5.9), Inches(8.0), Inches(0.6),
             "«Каждый воин должен понимать свой манёвр» —",
             size=14, color=DARK, font=F_SERIF, italic=True)
    add_text(s, Inches(1.0), Inches(6.25), Inches(8.0), Inches(0.4),
             "М.И. Кутузов",
             size=13, color=GRAY, font=F_BODY, italic=True)

    set_transition(s, "fade", 700)


def slide_refs():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    header(s, "Список использованной литературы", 12, TOTAL)
    add_bullets(s, Inches(0.8), Inches(1.55), Inches(8.6), Inches(5.3), [
        "Брагин М. Г. Кутузов. — М.: Молодая гвардия, 1975.",
        "Жилин П. А. Отечественная война 1812 года. — М.: Наука, 1988.",
        "Троицкий Н. А. Фельдмаршал Кутузов: мифы и факты. — М.: Центрполиграф, 2002.",
        "Учебник истории России. 8 класс.",
        "Интернет-ресурс: https://ru.wikipedia.org — статья «Кутузов, Михаил Илларионович».",
    ], size=18)
    set_transition(s, "fade", 600)


def slide_thanks():
    s = prs.slides.add_slide(BLANK); add_bg(s, WHITE)
    # крупная полоса сверху и снизу
    add_rect(s, 0, 0, SW, Inches(0.25), fill=BLUE)
    add_rect(s, 0, SH - Inches(0.25), SW, Inches(0.25), fill=BLUE)

    add_text(s, Inches(0.4), Inches(2.6), Inches(9.2), Inches(1.4),
             "Спасибо за внимание!",
             size=60, bold=True, color=BLUE, font=F_HEAD, align=PP_ALIGN.CENTER)

    add_text(s, Inches(0.4), Inches(4.3), Inches(9.2), Inches(0.5),
             "Презентацию подготовил:",
             size=18, color=DARK, font=F_BODY, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.4), Inches(4.8), Inches(9.2), Inches(0.6),
             "ученик 8 «А» класса Казаков Дмитрий",
             size=22, bold=True, color=DARK, font=F_BODY, align=PP_ALIGN.CENTER)
    add_text(s, Inches(0.4), Inches(5.6), Inches(9.2), Inches(0.4),
             "2026 г.",
             size=16, color=GRAY, font=F_BODY, align=PP_ALIGN.CENTER)
    set_transition(s, "fade", 700)


# ==========================================================
#   СБОРКА
# ==========================================================
slide_title()
slide_plan()
slide_bio()
slide_childhood()
slide_service()
slide_europe()
slide_1812()
slide_borodino()
slide_fili()
slide_expulsion()
slide_result()
slide_refs()
slide_thanks()

out = "Kutuzov.pptx"
prs.save(out)
print("OK:", out)
