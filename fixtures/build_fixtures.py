import zipfile
from pathlib import Path

OUTPUT_DIRECTORY = Path(__file__).resolve().parent.parent / "public"

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'mc:Ignorable="wps"')

def text_box(body_wrap):
    return (
        '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>'
        '<wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>'
        '<wp:extent cx="1176020" cy="361315"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '<wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/>'
        '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1176020" cy="361315"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>'
        '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Label text</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        f'<wps:bodyPr rot="0" vert="horz" wrap="{body_wrap}" lIns="0" tIns="0" rIns="0" bIns="190500" anchor="b" anchorCtr="0"><a:spAutoFit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>'
    )

def cell(width, content):
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr><w:p>{content}<w:r><w:t>Cell</w:t></w:r></w:p></w:tc>'

def table(table_width, first_cell_content):
    return (
        f'<w:tbl><w:tblPr><w:tblW {table_width}/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="3008"/><w:gridCol w:w="3011"/><w:gridCol w:w="3008"/></w:tblGrid>'
        f'<w:tr>{cell(3007, first_cell_content)}{cell(3011, "")}{cell(3008, "")}</w:tr></w:tbl>'
    )

VARIANTS = {
    "A-textbox-nowrap-in-pct-table": table('w:w="5000" w:type="pct"', text_box("none")) + '<w:p/>',
    "B-textbox-nowrap-no-table": f'<w:p>{text_box("none")}<w:r><w:t>Footer</w:t></w:r></w:p>',
    "C-textbox-nowrap-in-dxa-table": table('w:w="9027" w:type="dxa"', text_box("none")) + '<w:p/>',
    "D-textbox-square-in-pct-table": table('w:w="5000" w:type="pct"', text_box("square")) + '<w:p/>',
}

BODY = ''.join(f'<w:p><w:r><w:t>Paragraph {i} of the agreement body text.</w:t></w:r></w:p>' for i in range(40))
DOCUMENT = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{BODY}'
            '<w:sectPr><w:footerReference w:type="default" r:id="rId1"/><w:pgSz w:w="11907" w:h="16839"/>'
            '<w:pgMar w:top="709" w:right="1440" w:bottom="709" w:left="1440" w:header="540" w:footer="540" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>')
CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOCUMENT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')

for name, footer_body in VARIANTS.items():
    with zipfile.ZipFile(OUTPUT_DIRECTORY / f"{name}.docx", "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        archive.writestr("word/footer1.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {NS}>{footer_body}</w:ftr>')
    print(OUTPUT_DIRECTORY / f"{name}.docx")
