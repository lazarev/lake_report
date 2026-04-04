from pathlib import Path
from xml.sax.saxutils import escape


OUT = Path("4. Презентация/Презентация_биоиндикация.fodp")


SLIDES = [
    {
        "type": "title",
        "title": "Применение метода биоиндикации для сравнения водоемов Челябинской области",
        "subtitle": "Лазарев Григорий Андреевич\nУчебный проект",
    },
    {
        "type": "outline",
        "title": "Актуальность темы",
        "bullets": [
            "Водоемы важны для природы и человека.",
            "Их состояние зависит от антропогенной нагрузки.",
            "Биоиндикация позволяет оценивать состояние водной среды по живым организмам.",
        ],
    },
    {
        "type": "outline",
        "title": "Цель, объект, предмет и задачи",
        "bullets": [
            "Цель: сравнить экологическое состояние озера Чебаркуль и Шершневского водохранилища.",
            "Объект: два водоема Челябинской области.",
            "Предмет: растения-биоиндикаторы и другие признаки состояния водной среды.",
            "Задачи: изучить метод биоиндикации, описать водоемы, проанализировать растительность, сравнить результаты.",
        ],
    },
    {
        "type": "outline",
        "title": "Что такое биоиндикация",
        "bullets": [
            "Биоиндикация — это оценка состояния среды по живым организмам.",
            "Она помогает выявлять загрязнение, эвтрофикацию и изменения условий среды.",
            "Метод отражает не только текущее, но и накопленное воздействие на водоем.",
        ],
    },
    {
        "type": "outline",
        "title": "Объекты исследования",
        "bullets": [
            "Озеро Чебаркуль — естественный водоем тектонического происхождения.",
            "Шершневское водохранилище — искусственный водоем на реке Миасс.",
            "Главное различие: один водоем естественный, другой искусственный.",
        ],
    },
    {
        "type": "outline",
        "title": "Растительный компонент озера Чебаркуль",
        "bullets": [
            "Отмечены типичные прибрежные и водные растения.",
            "Биоиндикационное значение имеют элодея канадская и ряска трехдольная.",
            "Признаки эвтрофикации выражены слабо.",
            "Состояние озера можно оценить как удовлетворительное.",
        ],
    },
    {
        "type": "outline",
        "title": "Растительный компонент Шершневского водохранилища",
        "bullets": [
            "Антропогенная нагрузка здесь выше, чем у озера Чебаркуль.",
            "Особенно важны камыш, тростник и рогоз.",
            "Они защищают прибрежную зону и участвуют в самоочищении воды.",
            "Обнаружены паровые водоросли как дополнительный биоиндикационный признак.",
        ],
    },
    {
        "type": "outline",
        "title": "Сравнение двух водоемов",
        "bullets": [
            "Озеро Чебаркуль — естественный водоем с более спокойным экологическим состоянием.",
            "Шершневское водохранилище — искусственный водоем с более высокой антропогенной нагрузкой.",
            "Биоиндикация позволила наглядно выявить различия между ними.",
        ],
    },
    {
        "type": "outline",
        "title": "Выводы",
        "bullets": [
            "Метод биоиндикации подходит для сравнения водоемов.",
            "Озеро Чебаркуль оказалось более благополучным.",
            "Шершневское водохранилище имеет более напряженное экологическое состояние.",
            "Прибрежная растительность играет важную защитную роль.",
        ],
    },
    {
        "type": "outline",
        "title": "Практический результат работы",
        "bullets": [
            "По теме исследования была подготовлена статья.",
            "Результаты оформлены не только как проект, но и как отдельная публикация.",
            "Ссылка на публикацию: добавить после размещения статьи.",
        ],
    },
    {
        "type": "title",
        "title": "Спасибо за внимание!",
        "subtitle": "Готов ответить на вопросы",
    },
]


HEADER = """<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:officeooo="http://openoffice.org/2009/office" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:css3t="http://www.w3.org/TR/css3-text/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:rpt="http://openoffice.org/2005/report" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.presentation">
 <office:scripts></office:scripts>
 <office:font-face-decls>
  <style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="roman" style:font-pitch="variable"/>
  <style:font-face style:name="Noto Sans" svg:font-family="'Noto Sans'" style:font-family-generic="roman" style:font-pitch="variable"/>
  <style:font-face style:name="Arial Unicode MS" svg:font-family="'Arial Unicode MS'" style:font-family-generic="system" style:font-pitch="variable"/>
  <style:font-face style:name="PingFang SC" svg:font-family="'PingFang SC'" style:font-family-generic="system" style:font-pitch="variable"/>
  <style:font-face style:name="Tahoma" svg:font-family="Tahoma" style:font-family-generic="system" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:automatic-styles>
  <style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:background-visible="true" presentation:background-objects-visible="true" presentation:display-footer="true" presentation:display-page-number="false" presentation:display-date-time="true"/></style:style>
  <style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:display-header="true" presentation:display-footer="true" presentation:display-page-number="false" presentation:display-date-time="true"/></style:style>
  <style:style style:name="gr1" style:family="graphic"><style:graphic-properties style:protect="size"/></style:style>
  <style:style style:name="pr1" style:family="presentation" style:parent-style-name="Beehive-title"><style:graphic-properties draw:auto-grow-height="true" fo:min-height="4.506cm"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="pr2" style:family="presentation" style:parent-style-name="Beehive-subtitle"><style:graphic-properties draw:fill-color="#ffffff" draw:auto-grow-height="true" fo:min-height="3.2cm"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="pr3" style:family="presentation" style:parent-style-name="Beehive-notes"><style:graphic-properties draw:fill-color="#ffffff" fo:min-height="13.364cm"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="pr4" style:family="presentation" style:parent-style-name="Beehive1-title"><style:graphic-properties draw:auto-grow-height="true" fo:min-height="3.506cm"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="pr5" style:family="presentation" style:parent-style-name="Beehive1-outline1"><style:graphic-properties fo:min-height="2.95cm"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="pr6" style:family="presentation" style:parent-style-name="Beehive1-notes"><style:graphic-properties draw:fill-color="#ffffff" draw:auto-grow-height="true" fo:min-height="13.364cm"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="P1" style:family="paragraph"><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <style:style style:name="P2" style:family="paragraph"><loext:graphic-properties draw:fill-color="#ffffff"/><style:paragraph-properties style:writing-mode="lr-tb"/></style:style>
  <text:list-style style:name="L1">
   <text:list-level-style-bullet text:level="1" text:bullet-char="●"><style:list-level-properties text:min-label-width="0.6cm"/><style:text-properties fo:font-family="StarSymbol" style:use-window-font-color="true" fo:font-size="45%"/></text:list-level-style-bullet>
  </text:list-style>
 </office:automatic-styles>
 <office:body>
  <office:presentation>
"""


FOOTER = """   <presentation:settings presentation:mouse-visible="false"/>
  </office:presentation>
 </office:body>
</office:document>
"""


def p(text: str) -> str:
    return f"<text:p>{escape(text)}</text:p>"


def text_box_paragraphs(text: str) -> str:
    return "".join(p(line) for line in text.splitlines())


def bullet_list(items: list[str]) -> str:
    parts = ['<text:list text:style-name="L1">']
    for item in items:
        parts.append(f"<text:list-item>{p(item)}</text:list-item>")
    parts.append("</text:list>")
    return "".join(parts)


def notes(page_no: int, pr_name: str) -> str:
    return (
        f'<presentation:notes draw:style-name="dp2">'
        f'<draw:page-thumbnail draw:style-name="gr1" draw:layer="layout" svg:width="14.848cm" svg:height="11.136cm" svg:x="3.075cm" svg:y="2.257cm" draw:page-number="{page_no}" presentation:class="page"/>'
        f'<draw:frame presentation:style-name="{pr_name}" draw:text-style-name="P2" draw:layer="layout" svg:width="16.799cm" svg:height="13.364cm" svg:x="2.1cm" svg:y="14.107cm" presentation:class="notes" presentation:placeholder="true"><draw:text-box/></draw:frame>'
        f'</presentation:notes>'
    )


def make_title_slide(idx: int, title: str, subtitle: str) -> str:
    return f"""
   <draw:page draw:name="page{idx}" draw:style-name="dp1" draw:master-page-name="Beehive" presentation:presentation-page-layout-name="AL1T0">
    <office:forms form:automatic-focus="false" form:apply-design-mode="false"/>
    <draw:frame presentation:style-name="pr1" draw:text-style-name="P1" draw:layer="layout" svg:width="25.199cm" svg:height="4.506cm" svg:x="1.401cm" svg:y="7.094cm" presentation:class="title" presentation:user-transformed="true">
     <draw:text-box>{text_box_paragraphs(title)}</draw:text-box>
    </draw:frame>
    <draw:frame presentation:style-name="pr2" draw:text-style-name="P2" draw:layer="layout" svg:width="25.199cm" svg:height="3.2cm" svg:x="1.4cm" svg:y="12.6cm" presentation:class="subtitle" presentation:placeholder="true" presentation:user-transformed="true">
     <draw:text-box>{text_box_paragraphs(subtitle)}</draw:text-box>
    </draw:frame>
    {notes(idx, "pr3")}
   </draw:page>
"""


def make_outline_slide(idx: int, title: str, bullets: list[str]) -> str:
    return f"""
   <draw:page draw:name="page{idx}" draw:style-name="dp1" draw:master-page-name="Beehive1" presentation:presentation-page-layout-name="AL2T1">
    <office:forms form:automatic-focus="false" form:apply-design-mode="false"/>
    <draw:frame presentation:style-name="pr4" draw:text-style-name="P1" draw:layer="layout" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.837cm" presentation:class="title" presentation:user-transformed="true">
     <draw:text-box>{text_box_paragraphs(title)}</draw:text-box>
    </draw:frame>
    <draw:frame presentation:style-name="pr5" draw:text-style-name="P1" draw:layer="layout" svg:width="25.199cm" svg:height="12.179cm" svg:x="1.4cm" svg:y="4.914cm" presentation:class="outline" presentation:placeholder="true" presentation:user-transformed="true">
     <draw:text-box>{bullet_list(bullets)}</draw:text-box>
    </draw:frame>
    {notes(idx, "pr6")}
   </draw:page>
"""


parts = [HEADER]
for idx, slide in enumerate(SLIDES, start=1):
    if slide["type"] == "title":
        parts.append(make_title_slide(idx, slide["title"], slide["subtitle"]))
    else:
        parts.append(make_outline_slide(idx, slide["title"], slide["bullets"]))
parts.append(FOOTER)

OUT.write_text("".join(parts), encoding="utf-8")
print(OUT)
