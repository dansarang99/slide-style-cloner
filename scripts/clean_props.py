import zipfile
import shutil

fpath = r"C:\Users\note\vd\vd15_slide-style-cloner\result\팜드라이버_농기계대리운전_IR투자제안서.pptx"
temp_zip = fpath + ".tmp"

clean_core = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>팜드라이버 IR 투자제안서</dc:title>
  <dc:subject>농기계 대리운전 및 토탈 MRO 플랫폼</dc:subject>
  <dc:creator>(주)팜드라이버 IR팀</dc:creator>
  <cp:keywords>팜드라이버, 농기계대리운전, IR투자제안서, AgTech, 모빌리티</cp:keywords>
  <dc:description>주식회사 팜드라이버 공식 IR 투자제안서</dc:description>
  <cp:lastModifiedBy>(주)팜드라이버</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>"""

clean_app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Words>0</Words>
  <Application>Microsoft Macintosh PowerPoint</Application>
  <PresentationFormat>와이드스크린</PresentationFormat>
  <Paragraphs>0</Paragraphs>
  <Slides>30</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <Company>(주)팜드라이버</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>"""

with zipfile.ZipFile(fpath, 'r') as zin, zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == 'docProps/core.xml':
            zout.writestr(item, clean_core.encode('utf-8'))
        elif item.filename == 'docProps/app.xml':
            zout.writestr(item, clean_app.encode('utf-8'))
        else:
            zout.writestr(item, zin.read(item.filename))

shutil.move(temp_zip, fpath)
print("[SUCCESS] docProps completely cleaned!")
