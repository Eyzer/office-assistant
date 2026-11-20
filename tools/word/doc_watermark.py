"""给文档批量加水印（正文页眉里的 WordArt，倾斜半透明）。"""
import os

from docx.oxml import parse_xml

from utils import batch
from utils.file_utils import WORD_EXT
from utils.path_utils import out_path

META = {
    "id": "doc_watermark",
    "cat": "格式排版",
    "name": "文档加水印",
    "desc": "给每页加上倾斜的浅灰文字水印（如「内部资料」「保密」），可设字号与颜色。",
    "fields": [
        {"key": "text", "label": "水印文字", "type": "text", "default": "内部资料"},
        {"key": "size", "label": "字号", "type": "number", "default": 1,
         "help": "WordArt 缩放，1 为默认大小，2 更大"},
        {"key": "color", "label": "颜色", "type": "text", "default": "C0C0C0"},
    ],
}

_WM = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:w10="urn:schemas-microsoft-com:office:word">
 <w:r><w:rPr><w:noProof/></w:rPr><w:pict>
  <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"
   path="m@7,l@8,m@5,21600l@11,21600e">
   <v:formulas>
    <v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>
    <v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>
    <v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>
    <v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>
    <v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/>
   </v:formulas>
   <v:path textpathok="t" o:connecttype="custom"
    o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
   <v:textpath on="t" fitshape="t"/>
   <v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>
  </v:shapetype>
  <v:shape id="PowerPlusWaterMarkObject{uid}" o:spid="_x0000_s10{uid}"
   type="#_x0000_t136"
   style="position:absolute;margin-left:0;margin-top:0;width:{w}pt;height:{h}pt;rotation:-45;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
   o:allowincell="f" fillcolor="#{color}" stroked="f">
   <v:textpath style="font-family:&quot;宋体&quot;;font-size:1pt" string="{text}"/>
  </v:shape>
 </w:pict></w:r>
</w:p>"""


def add_watermark(doc, text, size, color):
    text = (text or "内部资料").replace('"', "")
    scale = max(0.5, float(size or 1))
    xml = _WM.format(uid=abs(hash(text)) % 900 + 100, w=int(468 * scale),
                     h=int(117 * scale), color=color or "C0C0C0", text=text)
    for section in doc.sections:
        try:
            header = section.header
            header.is_linked_to_previous = False
            p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            p._p.append(parse_xml(xml))
        except Exception:
            continue


def run_task(inputs, output_dir, params, ctx):
    text = params.get("text") or "内部资料"
    size = params.get("size") or 1
    color = params.get("color") or "C0C0C0"

    def handler(f, ctx):
        import docx
        doc = docx.Document(f)
        add_watermark(doc, text, size, color)
        dst = out_path(f, output_dir, suffix="_水印")
        doc.save(dst)
        ctx.info("%s: 已加水印「%s」" % (os.path.basename(f), text))
        return [dst]

    return batch.process(inputs, ctx, handler, WORD_EXT)
