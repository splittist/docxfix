"""Document generator for creating .docx files from specifications."""

import base64
import io
import random
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from docxfix.spec import (
    ChangeType,
    Comment,
    DocumentSpec,
    PageOrientation,
    Paragraph,
    SectionSpec,
)
from docxfix.xml_utils import XMLElement

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du"><w:zoom w:percent="100"/><w:removePersonalInformation/><w:removeDateAndTime/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="2050"/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/></w:compat><w:rsids><w:rsidRoot w:val="002B6F96"/><w:rsid w:val="00005A81"/><w:rsid w:val="00047834"/><w:rsid w:val="001A745D"/><w:rsid w:val="002B6F96"/><w:rsid w:val="0041366A"/><w:rsid w:val="004C6F26"/><w:rsid w:val="005E03F7"/><w:rsid w:val="007C54EA"/><w:rsid w:val="007E0A1F"/><w:rsid w:val="00CA009D"/><w:rsid w:val="00D46874"/><w:rsid w:val="00DC7384"/><w:rsid w:val="00F96A40"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-CH"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="2050"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/><w14:docId w14:val="4D55359B"/><w15:chartTrackingRefBased/></w:settings>
"""

WEB_SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du"><w:optimizeForBrowser/><w:allowPNG/></w:webSettings>
"""

FONT_TABLE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du"><w:font w:name="Aptos"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="20000287" w:usb1="00000003" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font><w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Aptos Display"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="20000287" w:usb1="00000003" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/></w:font></w:fonts>
"""

APP_PROPERTIES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>0</TotalTime><Pages>1</Pages><Words>35</Words><Characters>213</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>3</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>247</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>
"""

THEME_XML_B64 = """PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KPGE6dGhlbWUgeG1sbnM6YT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdpbmdtbC8yMDA2L21haW4iIG5hbWU9Ik9mZmljZSBUaGVtZSI+PGE6dGhlbWVFbGVtZW50cz48YTpjbHJTY2hlbWUgbmFtZT0iT2ZmaWNlIj48YTpkazE+PGE6c3lzQ2xyIHZhbD0id2luZG93VGV4dCIgbGFzdENscj0iMDAwMDAwIi8+PC9hOmRrMT48YTpsdDE+PGE6c3lzQ2xyIHZhbD0id2luZG93IiBsYXN0Q2xyPSJGRkZGRkYiLz48L2E6bHQxPjxhOmRrMj48YTpzcmdiQ2xyIHZhbD0iMEUyODQxIi8+PC9hOmRrMj48YTpsdDI+PGE6c3JnYkNsciB2YWw9IkU4RThFOCIvPjwvYTpsdDI+PGE6YWNjZW50MT48YTpzcmdiQ2xyIHZhbD0iMTU2MDgyIi8+PC9hOmFjY2VudDE+PGE6YWNjZW50Mj48YTpzcmdiQ2xyIHZhbD0iRTk3MTMyIi8+PC9hOmFjY2VudDI+PGE6YWNjZW50Mz48YTpzcmdiQ2xyIHZhbD0iMTk2QjI0Ii8+PC9hOmFjY2VudDM+PGE6YWNjZW50ND48YTpzcmdiQ2xyIHZhbD0iMEY5RUQ1Ii8+PC9hOmFjY2VudDQ+PGE6YWNjZW50NT48YTpzcmdiQ2xyIHZhbD0iQTAyQjkzIi8+PC9hOmFjY2VudDU+PGE6YWNjZW50Nj48YTpzcmdiQ2xyIHZhbD0iNEVBNzJFIi8+PC9hOmFjY2VudDY+PGE6aGxpbms+PGE6c3JnYkNsciB2YWw9IjQ2Nzg4NiIvPjwvYTpobGluaz48YTpmb2xIbGluaz48YTpzcmdiQ2xyIHZhbD0iOTY2MDdEIi8+PC9hOmZvbEhsaW5rPjwvYTpjbHJTY2hlbWU+PGE6Zm9udFNjaGVtZSBuYW1lPSJPZmZpY2UiPjxhOm1ham9yRm9udD48YTpsYXRpbiB0eXBlZmFjZT0iQXB0b3MgRGlzcGxheSIgcGFub3NlPSIwMjExMDAwNDAyMDIwMjAyMDIwNCIvPjxhOmVhIHR5cGVmYWNlPSIiLz48YTpjcyB0eXBlZmFjZT0iIi8+PGE6Zm9udCBzY3JpcHQ9IkpwYW4iIHR5cGVmYWNlPSLmuLjjgrTjgrfjg4Pjgq8gTGlnaHQiLz48YTpmb250IHNjcmlwdD0iSGFuZyIgdHlwZWZhY2U9IuunkeydgCDqs6DrlJUiLz48YTpmb250IHNjcmlwdD0iSGFucyIgdHlwZWZhY2U9Iuetiee6vyBMaWdodCIvPjxhOmZvbnQgc2NyaXB0PSJIYW50IiB0eXBlZmFjZT0i5paw57Sw5piO6auUIi8+PGE6Zm9udCBzY3JpcHQ9IkFyYWIiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iLz48YTpmb250IHNjcmlwdD0iSGViciIgdHlwZWZhY2U9IlRpbWVzIE5ldyBSb21hbiIvPjxhOmZvbnQgc2NyaXB0PSJUaGFpIiB0eXBlZmFjZT0iQW5nc2FuYSBOZXciLz48YTpmb250IHNjcmlwdD0iRXRoaSIgdHlwZWZhY2U9Ik55YWxhIi8+PGE6Zm9udCBzY3JpcHQ9IkJlbmciIHR5cGVmYWNlPSJWcmluZGEiLz48YTpmb250IHNjcmlwdD0iR3VqciIgdHlwZWZhY2U9IlNocnV0aSIvPjxhOmZvbnQgc2NyaXB0PSJLaG1yIiB0eXBlZmFjZT0iTW9vbEJvcmFuIi8+PGE6Zm9udCBzY3JpcHQ9IktuZGEiIHR5cGVmYWNlPSJUdW5nYSIvPjxhOmZvbnQgc2NyaXB0PSJHdXJ1IiB0eXBlZmFjZT0iUmFhdmkiLz48YTpmb250IHNjcmlwdD0iQ2FucyIgdHlwZWZhY2U9IkV1cGhlbWlhIi8+PGE6Zm9udCBzY3JpcHQ9IkNoZXIiIHR5cGVmYWNlPSJQbGFudGFnZW5ldCBDaGVyb2tlZSIvPjxhOmZvbnQgc2NyaXB0PSJZaWlpIiB0eXBlZmFjZT0iTWljcm9zb2Z0IFlpIEJhaXRpIi8+PGE6Zm9udCBzY3JpcHQ9IlRpYnQiIHR5cGVmYWNlPSJNaWNyb3NvZnQgSGltYWxheWEiLz48YTpmb250IHNjcmlwdD0iVGhhYSIgdHlwZWZhY2U9Ik1WIEJvbGkiLz48YTpmb250IHNjcmlwdD0iRGV2YSIgdHlwZWZhY2U9Ik1hbmdhbCIvPjxhOmZvbnQgc2NyaXB0PSJUZWx1IiB0eXBlZmFjZT0iR2F1dGFtaSIvPjxhOmZvbnQgc2NyaXB0PSJUYW1sIiB0eXBlZmFjZT0iTGF0aGEiLz48YTpmb250IHNjcmlwdD0iU3lyYyIgdHlwZWZhY2U9IkVzdHJhbmdlbG8gRWRlc3NhIi8+PGE6Zm9udCBzY3JpcHQ9Ik9yeWEiIHR5cGVmYWNlPSJLYWxpbmdhIi8+PGE6Zm9udCBzY3JpcHQ9Ik1seW0iIHR5cGVmYWNlPSJLYXJ0aWthIi8+PGE6Zm9udCBzY3JpcHQ9Ikxhb28iIHR5cGVmYWNlPSJEb2tDaGFtcGEiLz48YTpmb250IHNjcmlwdD0iU2luaCIgdHlwZWZhY2U9Iklza29vbGEgUG90YSIvPjxhOmZvbnQgc2NyaXB0PSJNb25nIiB0eXBlZmFjZT0iTW9uZ29saWFuIEJhaXRpIi8+PGE6Zm9udCBzY3JpcHQ9IlZpZXQiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iLz48YTpmb250IHNjcmlwdD0iVWlnaCIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBVaWdodXIiLz48YTpmb250IHNjcmlwdD0iR2VvciIgdHlwZWZhY2U9IlN5bGZhZW4iLz48YTpmb250IHNjcmlwdD0iQXJtbiIgdHlwZWZhY2U9IkFyaWFsIi8+PGE6Zm9udCBzY3JpcHQ9IkJ1Z2kiIHR5cGVmYWNlPSJMZWVsYXdhZGVlIFVJIi8+PGE6Zm9udCBzY3JpcHQ9IkJvcG8iIHR5cGVmYWNlPSJNaWNyb3NvZnQgSmhlbmdIZWkiLz48YTpmb250IHNjcmlwdD0iSmF2YSIgdHlwZWZhY2U9IkphdmFuZXNlIFRleHQiLz48YTpmb250IHNjcmlwdD0iTGlzdSIgdHlwZWZhY2U9IlNlZ29lIFVJIi8+PGE6Zm9udCBzY3JpcHQ9Ik15bXIiIHR5cGVmYWNlPSJNeWFubWFyIFRleHQiLz48YTpmb250IHNjcmlwdD0iTmtvbyIgdHlwZWZhY2U9IkVicmltYSIvPjxhOmZvbnQgc2NyaXB0PSJPbGNrIiB0eXBlZmFjZT0iTmlybWFsYSBVSSIvPjxhOmZvbnQgc2NyaXB0PSJPc21hIiB0eXBlZmFjZT0iRWJyaW1hIi8+PGE6Zm9udCBzY3JpcHQ9IlBoYWciIHR5cGVmYWNlPSJQaGFnc3BhIi8+PGE6Zm9udCBzY3JpcHQ9IlN5cm4iIHR5cGVmYWNlPSJFc3RyYW5nZWxvIEVkZXNzYSIvPjxhOmZvbnQgc2NyaXB0PSJTeXJqIiB0eXBlZmFjZT0iRXN0cmFuZ2VsbyBFZGVzc2EiLz48YTpmb250IHNjcmlwdD0iU3lyZSIgdHlwZWZhY2U9IkVzdHJhbmdlbG8gRWRlc3NhIi8+PGE6Zm9udCBzY3JpcHQ9IlNvcmEiIHR5cGVmYWNlPSJOaXJtYWxhIFVJIi8+PGE6Zm9udCBzY3JpcHQ9IlRhbGUiIHR5cGVmYWNlPSJNaWNyb3NvZnQgVGFpIExlIi8+PGE6Zm9udCBzY3JpcHQ9IlRhbHUiIHR5cGVmYWNlPSJNaWNyb3NvZnQgTmV3IFRhaSBMdWUiLz48YTpmb250IHNjcmlwdD0iVGZuZyIgdHlwZWZhY2U9IkVicmltYSIvPjwvYTptYWpvckZvbnQ+PGE6bWlub3JGb250PjxhOmxhdGluIHR5cGVmYWNlPSJBcHRvcyIgcGFub3NlPSIwMjExMDAwNDAyMDIwMjAyMDIwNCIvPjxhOmVhIHR5cGVmYWNlPSIiLz48YTpjcyB0eXBlZmFjZT0iIi8+PGE6Zm9udCBzY3JpcHQ9IkpwYW4iIHR5cGVmYWNlPSLmuLjmmI7mnJ0iLz48YTpmb250IHNjcmlwdD0iSGFuZyIgdHlwZWZhY2U9IuunkeydgCDqs6DrlJUiLz48YTpmb250IHNjcmlwdD0iSGFucyIgdHlwZWZhY2U9Iuetiee6vyIvPjxhOmZvbnQgc2NyaXB0PSJIYW50IiB0eXBlZmFjZT0i5paw57Sw5piO6auUIi8+PGE6Zm9udCBzY3JpcHQ9IkFyYWIiIHR5cGVmYWNlPSJBcmlhbCIvPjxhOmZvbnQgc2NyaXB0PSJIZWJyIiB0eXBlZmFjZT0iQXJpYWwiLz48YTpmb250IHNjcmlwdD0iVGhhaSIgdHlwZWZhY2U9IkNvcmRpYSBOZXciLz48YTpmb250IHNjcmlwdD0iRXRoaSIgdHlwZWZhY2U9Ik55YWxhIi8+PGE6Zm9udCBzY3JpcHQ9IkJlbmciIHR5cGVmYWNlPSJWcmluZGEiLz48YTpmb250IHNjcmlwdD0iR3VqciIgdHlwZWZhY2U9IlNocnV0aSIvPjxhOmZvbnQgc2NyaXB0PSJLaG1yIiB0eXBlZmFjZT0iRGF1blBlbmgiLz48YTpmb250IHNjcmlwdD0iS25kYSIgdHlwZWZhY2U9IlR1bmdhIi8+PGE6Zm9udCBzY3JpcHQ9Ikd1cnUiIHR5cGVmYWNlPSJSYWF2aSIvPjxhOmZvbnQgc2NyaXB0PSJDYW5zIiB0eXBlZmFjZT0iRXVwaGVtaWEiLz48YTpmb250IHNjcmlwdD0iQ2hlciIgdHlwZWZhY2U9IlBsYW50YWdlbmV0IENoZXJva2VlIi8+PGE6Zm9udCBzY3JpcHQ9IllpaWkiIHR5cGVmYWNlPSJNaWNyb3NvZnQgWWkgQmFpdGkiLz48YTpmb250IHNjcmlwdD0iVGlidCIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBIaW1hbGF5YSIvPjxhOmZvbnQgc2NyaXB0PSJUaGFhIiB0eXBlZmFjZT0iTVYgQm9saSIvPjxhOmZvbnQgc2NyaXB0PSJEZXZhIiB0eXBlZmFjZT0iTWFuZ2FsIi8+PGE6Zm9udCBzY3JpcHQ9IlRlbHUiIHR5cGVmYWNlPSJHYXV0YW1pIi8+PGE6Zm9udCBzY3JpcHQ9IlRhbWwiIHR5cGVmYWNlPSJMYXRoYSIvPjxhOmZvbnQgc2NyaXB0PSJTeXJjIiB0eXBlZmFjZT0iRXN0cmFuZ2VsbyBFZGVzc2EiLz48YTpmb250IHNjcmlwdD0iT3J5YSIgdHlwZWZhY2U9IkthbGluZ2EiLz48YTpmb250IHNjcmlwdD0iTWx5bSIgdHlwZWZhY2U9IkthcnRpa2EiLz48YTpmb250IHNjcmlwdD0iTGFvbyIgdHlwZWZhY2U9IkRva0NoYW1wYSIvPjxhOmZvbnQgc2NyaXB0PSJTaW5oIiB0eXBlZmFjZT0iSXNrb29sYSBQb3RhIi8+PGE6Zm9udCBzY3JpcHQ9Ik1vbmciIHR5cGVmYWNlPSJNb25nb2xpYW4gQmFpdGkiLz48YTpmb250IHNjcmlwdD0iVmlldCIgdHlwZWZhY2U9IkFyaWFsIi8+PGE6Zm9udCBzY3JpcHQ9IlVpZ2giIHR5cGVmYWNlPSJNaWNyb3NvZnQgVWlnaHVyIi8+PGE6Zm9udCBzY3JpcHQ9Ikdlb3IiIHR5cGVmYWNlPSJTeWxmYWVuIi8+PGE6Zm9udCBzY3JpcHQ9IkFybW4iIHR5cGVmYWNlPSJBcmlhbCIvPjxhOmZvbnQgc2NyaXB0PSJCdWdpIiB0eXBlZmFjZT0iTGVlbGF3YWRlZSBVSSIvPjxhOmZvbnQgc2NyaXB0PSJCb3BvIiB0eXBlZmFjZT0iTWljcm9zb2Z0IEpoZW5nSGVpIi8+PGE6Zm9udCBzY3JpcHQ9IkphdmEiIHR5cGVmYWNlPSJKYXZhbmVzZSBUZXh0Ii8+PGE6Zm9udCBzY3JpcHQ9Ikxpc3UiIHR5cGVmYWNlPSJTZWdvZSBVSSIvPjxhOmZvbnQgc2NyaXB0PSJNeW1yIiB0eXBlZmFjZT0iTXlhbm1hciBUZXh0Ii8+PGE6Zm9udCBzY3JpcHQ9Ik5rb28iIHR5cGVmYWNlPSJFYnJpbWEiLz48YTpmb250IHNjcmlwdD0iT2xjayIgdHlwZWZhY2U9Ik5pcm1hbGEgVUkiLz48YTpmb250IHNjcmlwdD0iT3NtYSIgdHlwZWZhY2U9IkVicmltYSIvPjxhOmZvbnQgc2NyaXB0PSJQaGFnIiB0eXBlZmFjZT0iUGhhZ3NwYSIvPjxhOmZvbnQgc2NyaXB0PSJTeXJuIiB0eXBlZmFjZT0iRXN0cmFuZ2VsbyBFZGVzc2EiLz48YTpmb250IHNjcmlwdD0iU3lyaiIgdHlwZWZhY2U9IkVzdHJhbmdlbG8gRWRlc3NhIi8+PGE6Zm9udCBzY3JpcHQ9IlN5cmUiIHR5cGVmYWNlPSJFc3RyYW5nZWxvIEVkZXNzYSIvPjxhOmZvbnQgc2NyaXB0PSJTb3JhIiB0eXBlZmFjZT0iTmlybWFsYSBVSSIvPjxhOmZvbnQgc2NyaXB0PSJUYWxlIiB0eXBlZmFjZT0iTWljcm9zb2Z0IFRhaSBMZSIvPjxhOmZvbnQgc2NyaXB0PSJUYWx1IiB0eXBlZmFjZT0iTWljcm9zb2Z0IE5ldyBUYWkgTHVlIi8+PGE6Zm9udCBzY3JpcHQ9IlRmbmciIHR5cGVmYWNlPSJFYnJpbWEiLz48L2E6bWlub3JGb250PjwvYTpmb250U2NoZW1lPjxhOmZtdFNjaGVtZSBuYW1lPSJPZmZpY2UiPjxhOmZpbGxTdHlsZUxzdD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjxhOmdyYWRGaWxsIHJvdFdpdGhTaGFwZT0iMSI+PGE6Z3NMc3Q+PGE6Z3MgcG9zPSIwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6bHVtTW9kIHZhbD0iMTEwMDAwIi8+PGE6c2F0TW9kIHZhbD0iMTA1MDAwIi8+PGE6dGludCB2YWw9IjY3MDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSI1MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOmx1bU1vZCB2YWw9IjEwNTAwMCIvPjxhOnNhdE1vZCB2YWw9IjEwMzAwMCIvPjxhOnRpbnQgdmFsPSI3MzAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6bHVtTW9kIHZhbD0iMTA1MDAwIi8+PGE6c2F0TW9kIHZhbD0iMTA5MDAwIi8+PGE6dGludCB2YWw9IjgxMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0PjxhOmxpbiBhbmc9IjU0MDAwMDAiIHNjYWxlZD0iMCIvPjwvYTpncmFkRmlsbD48YTpncmFkRmlsbCByb3RXaXRoU2hhcGU9IjEiPjxhOmdzTHN0PjxhOmdzIHBvcz0iMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNhdE1vZCB2YWw9IjEwMzAwMCIvPjxhOmx1bU1vZCB2YWw9IjEwMjAwMCIvPjxhOnRpbnQgdmFsPSI5NDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iNTAwMDAiPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTpzYXRNb2QgdmFsPSIxMTAwMDAiLz48YTpsdW1Nb2QgdmFsPSIxMDAwMDAiLz48YTpzaGFkZSB2YWw9IjEwMDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6bHVtTW9kIHZhbD0iOTkwMDAiLz48YTpzYXRNb2QgdmFsPSIxMjAwMDAiLz48YTpzaGFkZSB2YWw9Ijc4MDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0PjxhOmxpbiBhbmc9IjU0MDAwMDAiIHNjYWxlZD0iMCIvPjwvYTpncmFkRmlsbD48L2E6ZmlsbFN0eWxlTHN0PjxhOmxuU3R5bGVMc3Q+PGE6bG4gdz0iNjM1MCIgY2FwPSJmbGF0IiBjbXBkPSJzbmciIGFsZ249ImN0ciI+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIi8+PC9hOnNvbGlkRmlsbD48YTpwcnN0RGFzaCB2YWw9InNvbGlkIi8+PGE6bWl0ZXIgbGltPSI4MDAwMDAiLz48L2E6bG4+PGE6bG4gdz0iMTI3MDAiIGNhcD0iZmxhdCIgY21wZD0ic25nIiBhbGduPSJjdHIiPjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6cHJzdERhc2ggdmFsPSJzb2xpZCIvPjxhOm1pdGVyIGxpbT0iODAwMDAwIi8+PC9hOmxuPjxhOmxuIHc9IjE5MDUwIiBjYXA9ImZsYXQiIGNtcGQ9InNuZyIgYWxnbj0iY3RyIj48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6c29saWRGaWxsPjxhOnByc3REYXNoIHZhbD0ic29saWQiLz48YTptaXRlciBsaW09IjgwMDAwMCIvPjwvYTpsbj48L2E6bG5TdHlsZUxzdD48YTplZmZlY3RTdHlsZUxzdD48YTplZmZlY3RTdHlsZT48YTplZmZlY3RMc3QvPjwvYTplZmZlY3RTdHlsZT48YTplZmZlY3RTdHlsZT48YTplZmZlY3RMc3QvPjwvYTplZmZlY3RTdHlsZT48YTplZmZlY3RTdHlsZT48YTplZmZlY3RMc3Q+PGE6b3V0ZXJTaGR3IGJsdXJSYWQ9IjU3MTUwIiBkaXN0PSIxOTA1MCIgZGlyPSI1NDAwMDAwIiBhbGduPSJjdHIiIHJvdFdpdGhTaGFwZT0iMCI+PGE6c3JnYkNsciB2YWw9IjAwMDAwMCI+PGE6YWxwaGEgdmFsPSI2MzAwMCIvPjwvYTpzcmdiQ2xyPjwvYTpvdXRlclNoZHc+PC9hOmVmZmVjdExzdD48L2E6ZWZmZWN0U3R5bGU+PC9hOmVmZmVjdFN0eWxlTHN0PjxhOmJnRmlsbFN0eWxlTHN0PjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6c29saWRGaWxsPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTp0aW50IHZhbD0iOTUwMDAiLz48YTpzYXRNb2QgdmFsPSIxNzAwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpzb2xpZEZpbGw+PGE6Z3JhZEZpbGwgcm90V2l0aFNoYXBlPSIxIj48YTpnc0xzdD48YTpncyBwb3M9IjAiPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTp0aW50IHZhbD0iOTMwMDAiLz48YTpzYXRNb2QgdmFsPSIxNTAwMDAiLz48YTpzaGFkZSB2YWw9Ijk4MDAwIi8+PGE6bHVtTW9kIHZhbD0iMTAyMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6Z3MgcG9zPSI1MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI5ODAwMCIvPjxhOnNhdE1vZCB2YWw9IjEzMDAwMCIvPjxhOnNoYWRlIHZhbD0iOTAwMDAiLz48YTpsdW1Nb2QgdmFsPSIxMDMwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjEwMDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iNjMwMDAiLz48YTpzYXRNb2QgdmFsPSIxMjAwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48L2E6Z3NMc3Q+PGE6bGluIGFuZz0iNTQwMDAwMCIgc2NhbGVkPSIwIi8+PC9hOmdyYWRGaWxsPjwvYTpiZ0ZpbGxTdHlsZUxzdD48L2E6Zm10U2NoZW1lPjwvYTp0aGVtZUVsZW1lbnRzPjxhOm9iamVjdERlZmF1bHRzLz48YTpleHRyYUNsclNjaGVtZUxzdC8+PGE6ZXh0THN0PjxhOmV4dCB1cmk9InswNUE0QzI1Qy0wODVFLTQzNDAtODVBMy1BNTUzMUU1MTBEQjJ9Ij48dGhtMTU6dGhlbWVGYW1pbHkgeG1sbnM6dGhtMTU9Imh0dHA6Ly9zY2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3RoZW1lbWwvMjAxMi9tYWluIiBuYW1lPSJPZmZpY2UgVGhlbWUiIGlkPSJ7MkUxNDJBMkMtQ0QxNi00MkQ2LTg3M0EtQzI2RDJBMDUwNkZBfSIgdmlkPSJ7MUJEREZGNTItNkNENi00MEE1LUFCM0MtNjhFQjJGMUU0RDBBfSIvPjwvYTpleHQ+PC9hOmV4dExzdD48L2E6dGhlbWU+"""


class DocumentGenerator:
    """Generates .docx files from DocumentSpec."""

    # Core OOXML namespaces (used in code)
    NAMESPACES = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    }
    
    # Comprehensive Word namespace map (for compatibility)
    WORD_NAMESPACES = {
        "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
        "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
        "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
        "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
        "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
        "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
        "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
        "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
        "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        "o": "urn:schemas-microsoft-com:office:office",
        "oel": "http://schemas.microsoft.com/office/2019/extlst",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "v": "urn:schemas-microsoft-com:vml",
        "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "w10": "urn:schemas-microsoft-com:office:word",
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
        "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
        "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
        "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
        "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    }

    def __init__(self, spec: DocumentSpec) -> None:
        """Initialize generator with a document specification."""
        self.spec = spec
        self._revision_counter = 0
        self._comment_counter = 0
        self._comment_metadata = []  # Track comment metadata for multi-part generation
        self._section_layout = self._normalize_sections()
        self._section_header_footer_refs: dict[int, dict[str, dict[str, str]]] = {}

        # Initialize random seed if specified
        if spec.seed is not None:
            random.seed(spec.seed)

    def generate(self, output_path: str | Path) -> None:
        """
        Generate a .docx file at the specified path.

        Args:
            output_path: Path where the .docx file will be created
        """
        output_path = Path(output_path)

        # Check if document has comments
        has_comments = any(
            para.comments for para in self.spec.paragraphs
        )

        # Check if document has numbering
        has_numbering = any(
            para.numbering for para in self.spec.paragraphs
        )

        # Check if document has heading-based numbering
        has_heading_numbering = any(
            para.heading_level for para in self.spec.paragraphs
        )

        # Styles and numbering parts needed for either type
        needs_numbering = has_numbering or has_heading_numbering
        section_parts = self._build_section_part_manifest()

        # Create a ZIP file (docx is a ZIP archive)
        with zipfile.ZipFile(
            output_path, "w", zipfile.ZIP_DEFLATED
        ) as docx_zip:
            # Add required files
            docx_zip.writestr(
                "[Content_Types].xml",
                self._create_content_types(has_comments, needs_numbering, section_parts),
            )
            docx_zip.writestr("_rels/.rels", self._create_rels())
            docx_zip.writestr(
                "word/_rels/document.xml.rels",
                self._create_document_rels(has_comments, needs_numbering, section_parts),
            )
            docx_zip.writestr("word/document.xml", self._create_document())
            docx_zip.writestr("word/settings.xml", self._create_settings())
            docx_zip.writestr("word/webSettings.xml", self._create_web_settings())
            docx_zip.writestr("word/footnotes.xml", self._create_footnotes())
            docx_zip.writestr("word/endnotes.xml", self._create_endnotes())
            docx_zip.writestr("word/fontTable.xml", self._create_font_table())
            docx_zip.writestr("word/theme/theme1.xml", self._create_theme())
            docx_zip.writestr("docProps/core.xml", self._create_core_properties())
            docx_zip.writestr("docProps/app.xml", self._create_app_properties())

            for part in section_parts:
                docx_zip.writestr(part["path"], self._create_header_footer_part(part["kind"], part["text"]))

            # Add comment files if needed
            if has_comments:
                docx_zip.writestr("word/comments.xml", self._create_comments())
                docx_zip.writestr("word/commentsExtended.xml", self._create_comments_extended())
                docx_zip.writestr("word/commentsIds.xml", self._create_comments_ids())

            # Add numbering files if needed
            if needs_numbering:
                docx_zip.writestr("word/numbering.xml", self._create_numbering(has_numbering, has_heading_numbering))
                docx_zip.writestr("word/styles.xml", self._create_styles(has_heading_numbering))

    def _create_content_types(
        self,
        has_comments: bool = False,
        has_numbering: bool = False,
        section_parts: list[dict] | None = None,
    ) -> bytes:
        """Create [Content_Types].xml."""
        section_parts = section_parts or []
        types = etree.Element(
            "Types",
            xmlns="http://schemas.openxmlformats.org/package/2006/content-types",
        )
        etree.SubElement(
            types,
            "Default",
            Extension="rels",
            ContentType="application/vnd.openxmlformats-package.relationships+xml",
        )
        etree.SubElement(
            types,
            "Default",
            Extension="xml",
            ContentType="application/xml",
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/document.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document.main+xml"
            ),
        )

        # Add comment content types if needed
        if has_comments:
            etree.SubElement(
                types,
                "Override",
                PartName="/word/comments.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.comments+xml"
                ),
            )
            etree.SubElement(
                types,
                "Override",
                PartName="/word/commentsExtended.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.commentsExtended+xml"
                ),
            )
            etree.SubElement(
                types,
                "Override",
                PartName="/word/commentsIds.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.commentsIds+xml"
                ),
            )

        # Add numbering and styles content types if needed
        if has_numbering:
            etree.SubElement(
                types,
                "Override",
                PartName="/word/numbering.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.numbering+xml"
                ),
            )
            etree.SubElement(
                types,
                "Override",
                PartName="/word/styles.xml",
                ContentType=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.styles+xml"
                ),
            )

        for part in section_parts:
            etree.SubElement(
                types,
                "Override",
                PartName=f"/word/{part['path'].split('/', 1)[1]}",
                ContentType=part["content_type"],
            )

        # Add Word compatibility parts
        etree.SubElement(
            types,
            "Override",
            PartName="/word/settings.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.settings+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/webSettings.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.webSettings+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/footnotes.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.footnotes+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/endnotes.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.endnotes+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/fontTable.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.fontTable+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/word/theme/theme1.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "theme+xml"
            ),
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/docProps/core.xml",
            ContentType="application/vnd.openxmlformats-package.core-properties+xml",
        )
        etree.SubElement(
            types,
            "Override",
            PartName="/docProps/app.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument."
                "extended-properties+xml"
            ),
        )

        return etree.tostring(
            types, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_rels(self) -> bytes:
        """Create _rels/.rels."""
        rels = etree.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId1",
            Type=(
                "http://schemas.openxmlformats.org/officeDocument/"
                "2006/relationships/officeDocument"
            ),
            Target="word/document.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId2",
            Type=(
                "http://schemas.openxmlformats.org/package/2006/relationships/"
                "metadata/core-properties"
            ),
            Target="docProps/core.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id="rId3",
            Type=(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                "extended-properties"
            ),
            Target="docProps/app.xml",
        )
        return etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_document_rels(
        self,
        has_comments: bool = False,
        has_numbering: bool = False,
        section_parts: list[dict] | None = None,
    ) -> bytes:
        """Create word/_rels/document.xml.rels."""
        section_parts = section_parts or []
        rels = etree.Element(
            "Relationships",
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships",
        )

        next_id = 1

        # Add comment relationships if needed
        if has_comments:
            etree.SubElement(
                rels,
                "Relationship",
                Id=f"rId{next_id}",
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                Target="comments.xml",
            )
            next_id += 1
            etree.SubElement(
                rels,
                "Relationship",
                Id=f"rId{next_id}",
                Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                Target="commentsExtended.xml",
            )
            next_id += 1
            etree.SubElement(
                rels,
                "Relationship",
                Id=f"rId{next_id}",
                Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                Target="commentsIds.xml",
            )
            next_id += 1

        # Add numbering and styles relationships if needed
        if has_numbering:
            etree.SubElement(
                rels,
                "Relationship",
                Id=f"rId{next_id}",
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                Target="numbering.xml",
            )
            next_id += 1
            etree.SubElement(
                rels,
                "Relationship",
                Id=f"rId{next_id}",
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                Target="styles.xml",
            )
            next_id += 1

        for part in section_parts:
            rid = f"rId{next_id}"
            next_id += 1
            etree.SubElement(
                rels,
                "Relationship",
                Id=rid,
                Type=part["relationship_type"],
                Target=part["target"],
            )
            self._section_header_footer_refs[part["section_index"]][part["kind"]][part["variant"]] = rid

        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id}",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            Target="settings.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id + 1}",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
            Target="webSettings.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id + 2}",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
            Target="footnotes.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id + 3}",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
            Target="endnotes.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id + 4}",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
            Target="fontTable.xml",
        )
        etree.SubElement(
            rels,
            "Relationship",
            Id=f"rId{next_id + 5}",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            Target="theme/theme1.xml",
        )

        return etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_document(self) -> bytes:
        """Create word/document.xml with paragraphs and features."""
        document = etree.Element(
            f"{{{self.NAMESPACES['w']}}}document",
            nsmap=self.WORD_NAMESPACES,
        )
        document.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")
        body = etree.SubElement(document, f"{{{self.NAMESPACES['w']}}}body")

        sections = self._section_layout
        paragraph_count = len(self.spec.paragraphs)
        section_starts = [section.start_paragraph for section in sections] + [paragraph_count]

        boundary_to_section: dict[int, SectionSpec] = {}
        for idx in range(len(sections) - 1):
            boundary_to_section[section_starts[idx + 1] - 1] = sections[idx]

        # Add each paragraph
        for para_index, para_spec in enumerate(self.spec.paragraphs):
            para = self._add_paragraph(body, para_spec)
            if para_index in boundary_to_section:
                self._add_section_properties(para, boundary_to_section[para_index], is_body_level=False)

        self._add_section_properties(body, sections[-1], is_body_level=True)

        return etree.tostring(
            document, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _add_paragraph(self, body: XMLElement, para_spec: Paragraph) -> XMLElement:

        """Add a paragraph to the body."""
        w_ns = self.NAMESPACES["w"]
        w14_ns = self.NAMESPACES["w14"]
        para = etree.SubElement(body, f"{{{w_ns}}}p")
        
        # Generate unique paraId for paragraph
        para_id = self._generate_hex_id(8)
        para.set(f"{{{w14_ns}}}paraId", para_id)
        para.set(f"{{{w14_ns}}}textId", "77777777")  # Static textId for now
        
        # Add paragraph properties if needed (for numbering or heading)
        if para_spec.numbering:
            pPr = etree.SubElement(para, f"{{{w_ns}}}pPr")

            # Add paragraph style (ListParagraph for numbered lists)
            pStyle = etree.SubElement(pPr, f"{{{w_ns}}}pStyle")
            pStyle.set(f"{{{w_ns}}}val", "ListParagraph")

            # Add numbering properties
            numPr = etree.SubElement(pPr, f"{{{w_ns}}}numPr")

            # Set indentation level (ilvl)
            ilvl = etree.SubElement(numPr, f"{{{w_ns}}}ilvl")
            ilvl.set(f"{{{w_ns}}}val", str(para_spec.numbering.level))

            # Set numbering ID (numId)
            numId = etree.SubElement(numPr, f"{{{w_ns}}}numId")
            numId.set(f"{{{w_ns}}}val", str(para_spec.numbering.numbering_id))
        elif para_spec.heading_level:
            pPr = etree.SubElement(para, f"{{{w_ns}}}pPr")
            pStyle = etree.SubElement(pPr, f"{{{w_ns}}}pStyle")
            pStyle.set(f"{{{w_ns}}}val", f"Heading{para_spec.heading_level}")

        # Handle different content types
        if para_spec.comments and para_spec.tracked_changes:
            self._add_paragraph_with_comments_and_tracked_changes(para, para_spec)
        elif para_spec.comments:
            self._add_paragraph_with_comments(para, para_spec)
        elif para_spec.tracked_changes:
            self._add_paragraph_with_tracked_changes(para, para_spec)
        else:
            # Simple run with text (applies to both numbered and regular paragraphs)
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = para_spec.text

        return para

    def _add_text_run(self, parent: XMLElement, text: str) -> None:
        """Add a plain text run, setting xml:space='preserve' when needed."""
        w_ns = self.NAMESPACES["w"]
        run = etree.SubElement(parent, f"{{{w_ns}}}r")
        text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
        text_elem.text = text
        if text != text.strip():
            text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def _add_paragraph_with_tracked_changes(
        self, para: XMLElement, para_spec: Paragraph
    ) -> None:
        """Add paragraph content interleaving plain text with tracked changes.

        When the paragraph has base text, tracked changes are positioned within
        it: deletions are located by matching their ``text`` as a substring;
        insertions are placed after their ``insert_after`` substring.  When the
        paragraph text is empty, changes are emitted sequentially (legacy
        behaviour).
        """
        w_ns = self.NAMESPACES["w"]
        base_text = para_spec.text

        if not base_text:
            # Legacy path: no base text, just emit changes in order
            for change in para_spec.tracked_changes:
                self._emit_tracked_change(para, change)
            return

        # Build a sorted list of (position, event) markers.
        # Each event is ("ins", change) or ("del", change).
        events: list[tuple[int, str, object]] = []

        for change in para_spec.tracked_changes:
            if change.change_type == ChangeType.DELETION:
                idx = base_text.find(change.text)
                if idx == -1:
                    # Deleted text not found – append as standalone
                    events.append((len(base_text), "del", change))
                else:
                    events.append((idx, "del", change))
            else:
                # Insertion
                if change.insert_after:
                    idx = base_text.find(change.insert_after)
                    if idx != -1:
                        events.append((idx + len(change.insert_after), "ins", change))
                    else:
                        events.append((len(base_text), "ins", change))
                else:
                    events.append((len(base_text), "ins", change))

        # Sort by position, with insertions before deletions at the same pos
        events.sort(key=lambda e: (e[0], 0 if e[1] == "ins" else 1))

        # Walk through base_text, emitting plain runs and tracked changes
        cursor = 0
        for pos, kind, change in events:
            if kind == "del":
                # Emit plain text before the deletion
                if pos > cursor:
                    self._add_text_run(para, base_text[cursor:pos])
                # Emit the deletion element
                self._emit_tracked_change(para, change)
                cursor = pos + len(change.text)
            else:
                # Insertion: emit plain text up to the insertion point
                if pos > cursor:
                    self._add_text_run(para, base_text[cursor:pos])
                    cursor = pos
                self._emit_tracked_change(para, change)

        # Emit remaining plain text
        if cursor < len(base_text):
            self._add_text_run(para, base_text[cursor:])

    def _emit_tracked_change(self, para: XMLElement, change) -> None:
        """Emit a single ``<w:ins>`` or ``<w:del>`` element."""
        w_ns = self.NAMESPACES["w"]
        self._revision_counter += 1

        # Format date
        date_str = change.date.strftime("%Y-%m-%dT%H:%M:%SZ")

        if change.change_type == ChangeType.INSERTION:
            ins = etree.SubElement(
                para,
                f"{{{w_ns}}}ins",
                {
                    f"{{{w_ns}}}id": str(self._revision_counter),
                    f"{{{w_ns}}}author": change.author,
                    f"{{{w_ns}}}date": date_str,
                },
            )
            run = etree.SubElement(ins, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = change.text
            if change.text != change.text.strip():
                text_elem.set(
                    "{http://www.w3.org/XML/1998/namespace}space", "preserve"
                )

        elif change.change_type == ChangeType.DELETION:
            delete = etree.SubElement(
                para,
                f"{{{w_ns}}}del",
                {
                    f"{{{w_ns}}}id": str(self._revision_counter),
                    f"{{{w_ns}}}author": change.author,
                    f"{{{w_ns}}}date": date_str,
                },
            )
            run = etree.SubElement(delete, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}delText")
            text_elem.text = change.text
            if change.text != change.text.strip():
                text_elem.set(
                    "{http://www.w3.org/XML/1998/namespace}space", "preserve"
                )

    def _add_paragraph_with_comments_and_tracked_changes(
        self, para: XMLElement, para_spec: Paragraph
    ) -> None:
        """Add paragraph content with both comments and tracked changes.

        Merges comment anchoring and tracked-change positioning into a single
        walk over the base text.  Events (comment-start, comment-end,
        insertion, deletion) are collected with their text positions, sorted,
        and emitted in order so both features coexist in one paragraph.
        """
        w_ns = self.NAMESPACES["w"]
        base_text = para_spec.text

        # --- 1. Build tracked-change events (same logic as _add_paragraph_with_tracked_changes) ---
        tc_events: list[tuple[int, str, object]] = []
        for change in para_spec.tracked_changes:
            if change.change_type == ChangeType.DELETION:
                idx = base_text.find(change.text)
                if idx == -1:
                    tc_events.append((len(base_text), "del", change))
                else:
                    tc_events.append((idx, "del", change))
            else:
                if change.insert_after:
                    idx = base_text.find(change.insert_after)
                    if idx != -1:
                        tc_events.append((idx + len(change.insert_after), "ins", change))
                    else:
                        tc_events.append((len(base_text), "ins", change))
                else:
                    tc_events.append((len(base_text), "ins", change))

        # --- 2. Build comment events ---
        # Each comment produces a "comment_start" and "comment_end" event.
        # We also track metadata so we can emit commentRangeEnd + reference runs.
        comment_events: list[tuple[int, str, object]] = []
        comment_infos: list[dict] = []  # parallel to para_spec.comments

        for comment in para_spec.comments:
            anchor_text = comment.anchor_text
            if anchor_text in base_text:
                start = base_text.index(anchor_text)
                end = start + len(anchor_text)
            else:
                start = 0
                end = len(base_text)

            # Register comment metadata (same as _add_paragraph_with_comments)
            comment_id = str(self._comment_counter)
            parent_para_id = self._generate_hex_id(8).upper()
            durable_id = parent_para_id

            self._comment_metadata.append({
                "id": comment_id,
                "para_id": parent_para_id,
                "durable_id": durable_id,
                "author": comment.author,
                "date": comment.date,
                "text": comment.text,
                "resolved": comment.resolved,
                "parent_para_id": None,
            })
            self._comment_counter += 1

            reply_ids = []
            for reply in comment.replies:
                reply_id = str(self._comment_counter)
                reply_para_id = self._generate_hex_id(8).upper()
                reply_durable_id = self._generate_hex_id(8).upper()
                self._comment_metadata.append({
                    "id": reply_id,
                    "para_id": reply_para_id,
                    "durable_id": reply_durable_id,
                    "author": reply.author,
                    "date": reply.date,
                    "text": reply.text,
                    "resolved": comment.resolved,
                    "parent_para_id": parent_para_id,
                })
                reply_ids.append(reply_id)
                self._comment_counter += 1

            info = {"comment_id": comment_id, "reply_ids": reply_ids}
            comment_infos.append(info)

            comment_events.append((start, "comment_start", info))
            comment_events.append((end, "comment_end", info))

        # --- 3. Merge all events and sort ---
        # Priority within the same position:
        #   comment_start (0) < ins (1) < del (2) < comment_end (3)
        # This ensures markers wrap around content correctly.
        ORDER = {"comment_start": 0, "ins": 1, "del": 2, "comment_end": 3}
        all_events: list[tuple[int, str, object]] = tc_events + comment_events
        all_events.sort(key=lambda e: (e[0], ORDER.get(e[1], 5)))

        # --- 4. Walk and emit ---
        cursor = 0
        for pos, kind, payload in all_events:
            if kind == "comment_start":
                # Emit any plain text before the comment starts
                if pos > cursor:
                    self._add_text_run(para, base_text[cursor:pos])
                    cursor = pos
                # Emit commentRangeStart for main + replies
                etree.SubElement(para, f"{{{w_ns}}}commentRangeStart", {f"{{{w_ns}}}id": payload["comment_id"]})
                for rid in payload["reply_ids"]:
                    etree.SubElement(para, f"{{{w_ns}}}commentRangeStart", {f"{{{w_ns}}}id": rid})

            elif kind == "comment_end":
                # Emit any plain text before the comment ends
                if pos > cursor:
                    self._add_text_run(para, base_text[cursor:pos])
                    cursor = pos
                # Emit commentRangeEnd + reference runs
                etree.SubElement(para, f"{{{w_ns}}}commentRangeEnd", {f"{{{w_ns}}}id": payload["comment_id"]})
                for rid in payload["reply_ids"]:
                    etree.SubElement(para, f"{{{w_ns}}}commentRangeEnd", {f"{{{w_ns}}}id": rid})
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                self._add_comment_reference_run(run, payload["comment_id"])
                for rid in payload["reply_ids"]:
                    run = etree.SubElement(para, f"{{{w_ns}}}r")
                    self._add_comment_reference_run(run, rid)

            elif kind == "del":
                if pos > cursor:
                    self._add_text_run(para, base_text[cursor:pos])
                self._emit_tracked_change(para, payload)
                cursor = pos + len(payload.text)

            else:  # "ins"
                if pos > cursor:
                    self._add_text_run(para, base_text[cursor:pos])
                    cursor = pos
                self._emit_tracked_change(para, payload)

        # Emit remaining plain text
        if cursor < len(base_text):
            self._add_text_run(para, base_text[cursor:])

    def _add_paragraph_with_comments(self, para: XMLElement, para_spec: Paragraph) -> None:
        """Add a paragraph with comment anchoring."""
        w_ns = self.NAMESPACES["w"]
        
        # Split text into before, anchor, and after parts
        # For simplicity, we'll find the anchor_text in the paragraph text
        # and add comment markers around it
        
        for comment in para_spec.comments:
            anchor_text = comment.anchor_text
            full_text = para_spec.text
            
            # Find anchor position
            if anchor_text not in full_text:
                # If anchor text not found, just comment the whole paragraph
                anchor_start = 0
                anchor_end = len(full_text)
                before_text = ""
                after_text = ""
            else:
                anchor_start = full_text.index(anchor_text)
                anchor_end = anchor_start + len(anchor_text)
                before_text = full_text[:anchor_start]
                after_text = full_text[anchor_end:]
            
            # Create comment ID and metadata
            comment_id = str(self._comment_counter)
            parent_para_id = self._generate_hex_id(8).upper()
            # For main comments, durableId must equal paraId (Word requirement)
            durable_id = parent_para_id
            
            # Store metadata for later use in comment files
            self._comment_metadata.append({
                "id": comment_id,
                "para_id": parent_para_id,
                "durable_id": durable_id,
                "author": comment.author,
                "date": comment.date,
                "text": comment.text,
                "resolved": comment.resolved,
                "parent_para_id": None,  # No parent for main comment
            })
            
            self._comment_counter += 1
            
            # Handle replies
            reply_ids = []
            for reply in comment.replies:
                reply_id = str(self._comment_counter)
                reply_para_id = self._generate_hex_id(8).upper()
                reply_durable_id = self._generate_hex_id(8).upper()
                
                self._comment_metadata.append({
                    "id": reply_id,
                    "para_id": reply_para_id,
                    "durable_id": reply_durable_id,
                    "author": reply.author,
                    "date": reply.date,
                    "text": reply.text,
                    "resolved": comment.resolved,
                    "parent_para_id": parent_para_id,  # Link to parent comment
                })
                
                reply_ids.append(reply_id)
                self._comment_counter += 1
            
            # Add text before anchor
            if before_text:
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
                text_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text_elem.text = before_text
            
            # Add comment range starts: main comment first, then replies
            etree.SubElement(para, f"{{{w_ns}}}commentRangeStart", {f"{{{w_ns}}}id": comment_id})
            for reply_id in reply_ids:
                etree.SubElement(para, f"{{{w_ns}}}commentRangeStart", {f"{{{w_ns}}}id": reply_id})

            # Add the anchored text
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = anchor_text

            # Add all comment range ends together
            etree.SubElement(para, f"{{{w_ns}}}commentRangeEnd", {f"{{{w_ns}}}id": comment_id})
            for reply_id in reply_ids:
                etree.SubElement(para, f"{{{w_ns}}}commentRangeEnd", {f"{{{w_ns}}}id": reply_id})

            # Add all comment references together
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            self._add_comment_reference_run(run, comment_id)
            for reply_id in reply_ids:
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                self._add_comment_reference_run(run, reply_id)
            
            # Add text after anchor
            if after_text:
                run = etree.SubElement(para, f"{{{w_ns}}}r")
                text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
                text_elem.text = after_text

    def _generate_hex_id(self, length: int = 8) -> str:
        """Generate a random hexadecimal ID of specified length."""
        return "".join(random.choices("0123456789ABCDEF", k=length))

    def _add_comment_reference_run(self, run: XMLElement, comment_id: str) -> None:
        """Add a styled comment reference run element."""
        w_ns = self.NAMESPACES["w"]
        r_pr = etree.SubElement(run, f"{{{w_ns}}}rPr")
        r_style = etree.SubElement(r_pr, f"{{{w_ns}}}rStyle")
        r_style.set(f"{{{w_ns}}}val", "CommentReference")
        sz = etree.SubElement(r_pr, f"{{{w_ns}}}sz")
        sz.set(f"{{{w_ns}}}val", "24")
        sz_cs = etree.SubElement(r_pr, f"{{{w_ns}}}szCs")
        sz_cs.set(f"{{{w_ns}}}val", "24")
        etree.SubElement(run, f"{{{w_ns}}}commentReference", {f"{{{w_ns}}}id": comment_id})

    def _normalize_sections(self) -> list[SectionSpec]:
        """Return sections sorted by start paragraph with an initial section."""
        paragraph_count = len(self.spec.paragraphs)
        sections = sorted(self.spec.sections, key=lambda section: section.start_paragraph)
        normalized: list[SectionSpec] = []
        seen: set[int] = set()
        for section in sections:
            if section.start_paragraph >= paragraph_count and paragraph_count > 0:
                continue
            if section.start_paragraph in seen:
                continue
            seen.add(section.start_paragraph)
            normalized.append(section)

        if not normalized:
            normalized = [SectionSpec(start_paragraph=0)]
        elif normalized[0].start_paragraph != 0:
            normalized.insert(0, SectionSpec(start_paragraph=0))

        return normalized

    def _build_section_part_manifest(self) -> list[dict]:
        """Create section header/footer part metadata and relationship mapping."""
        relationship_type_map = {
            "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
            "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        }
        content_type_map = {
            "header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        }

        manifest: list[dict] = []
        self._section_header_footer_refs = {}

        for section_index, section in enumerate(self._section_layout):
            refs: dict[str, dict[str, str]] = {"header": {}, "footer": {}}
            for kind, set_obj in (("header", section.headers), ("footer", section.footers)):
                for variant in ("default", "first", "even"):
                    text = getattr(set_obj, variant)
                    if text is None:
                        continue
                    part_num = len(manifest) + 1
                    part_name = f"{kind}{part_num}.xml"
                    manifest.append(
                        {
                            "section_index": section_index,
                            "kind": kind,
                            "variant": variant,
                            "path": f"word/{part_name}",
                            "target": part_name,
                            "text": text,
                            "relationship_type": relationship_type_map[kind],
                            "content_type": content_type_map[kind],
                        }
                    )
                    refs[kind][variant] = ""
            self._section_header_footer_refs[section_index] = refs

        return manifest

    def _add_section_properties(
        self,
        parent: XMLElement,
        section: SectionSpec,
        *,
        is_body_level: bool,
    ) -> None:
        """Add section properties to paragraph boundary or body."""
        w_ns = self.NAMESPACES["w"]
        r_ns = self.NAMESPACES["r"]

        if is_body_level:
            sect_pr = etree.SubElement(parent, f"{{{w_ns}}}sectPr")
        else:
            p_pr = parent.find(f"{{{w_ns}}}pPr")
            if p_pr is None:
                p_pr = etree.SubElement(parent, f"{{{w_ns}}}pPr")
            sect_pr = etree.SubElement(p_pr, f"{{{w_ns}}}sectPr")
            etree.SubElement(sect_pr, f"{{{w_ns}}}type", {f"{{{w_ns}}}val": section.break_type})

        section_index = self._section_layout.index(section)
        refs = self._section_header_footer_refs.get(section_index, {"header": {}, "footer": {}})

        for kind, ref_type in (("header", "headerReference"), ("footer", "footerReference")):
            for variant, rel_type in (("default", "default"), ("first", "first"), ("even", "even")):
                rid = refs.get(kind, {}).get(variant)
                if rid:
                    etree.SubElement(
                        sect_pr,
                        f"{{{w_ns}}}{ref_type}",
                        {f"{{{w_ns}}}type": rel_type, f"{{{r_ns}}}id": rid},
                    )

        if refs.get("header", {}).get("first") or refs.get("footer", {}).get("first"):
            etree.SubElement(sect_pr, f"{{{w_ns}}}titlePg")

        if section.restart_page_numbering or section.page_number_start is not None:
            attrs = {}
            if section.page_number_start is not None:
                attrs[f"{{{w_ns}}}start"] = str(section.page_number_start)
            elif section.restart_page_numbering:
                attrs[f"{{{w_ns}}}start"] = "1"
            etree.SubElement(sect_pr, f"{{{w_ns}}}pgNumType", attrs)

        pg_sz = etree.SubElement(sect_pr, f"{{{w_ns}}}pgSz")
        if section.orientation == PageOrientation.LANDSCAPE:
            pg_sz.set(f"{{{w_ns}}}w", "16838")
            pg_sz.set(f"{{{w_ns}}}h", "11906")
            pg_sz.set(f"{{{w_ns}}}orient", "landscape")
        else:
            pg_sz.set(f"{{{w_ns}}}w", "11906")
            pg_sz.set(f"{{{w_ns}}}h", "16838")

        pg_mar = etree.SubElement(sect_pr, f"{{{w_ns}}}pgMar")
        pg_mar.set(f"{{{w_ns}}}top", "1440")
        pg_mar.set(f"{{{w_ns}}}right", "1440")
        pg_mar.set(f"{{{w_ns}}}bottom", "1440")
        pg_mar.set(f"{{{w_ns}}}left", "1440")
        pg_mar.set(f"{{{w_ns}}}header", "708")
        pg_mar.set(f"{{{w_ns}}}footer", "708")
        pg_mar.set(f"{{{w_ns}}}gutter", "0")
        cols = etree.SubElement(sect_pr, f"{{{w_ns}}}cols")
        cols.set(f"{{{w_ns}}}space", "708")
        doc_grid = etree.SubElement(sect_pr, f"{{{w_ns}}}docGrid")
        doc_grid.set(f"{{{w_ns}}}linePitch", "360")

    def _create_header_footer_part(self, kind: str, text: str) -> bytes:
        """Create a basic header or footer part containing one paragraph."""
        w_ns = self.NAMESPACES["w"]
        tag = "hdr" if kind == "header" else "ftr"
        part = etree.Element(f"{{{w_ns}}}{tag}", nsmap=self.WORD_NAMESPACES)
        para = etree.SubElement(part, f"{{{w_ns}}}p")
        run = etree.SubElement(para, f"{{{w_ns}}}r")
        text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
        text_elem.text = text
        return etree.tostring(part, xml_declaration=True, encoding="UTF-8", pretty_print=True)

    def _create_comments(self) -> bytes:
        """Create word/comments.xml."""
        w_ns = self.NAMESPACES["w"]
        w14_ns = self.NAMESPACES["w14"]
        
        comments = etree.Element(
            f"{{{w_ns}}}comments",
            nsmap=self.WORD_NAMESPACES,
        )
        comments.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")
        
        # Add each comment
        for metadata in self._comment_metadata:
            comment = etree.SubElement(
                comments,
                f"{{{w_ns}}}comment",
                {
                    f"{{{w_ns}}}id": metadata["id"],
                    f"{{{w_ns}}}author": metadata["author"],
                    f"{{{w_ns}}}initials": metadata["author"][0] if metadata["author"] else "A",
                },
            )
            
            # Add comment paragraph
            para = etree.SubElement(comment, f"{{{w_ns}}}p")
            para.set(f"{{{w14_ns}}}paraId", metadata["para_id"])
            para.set(f"{{{w14_ns}}}textId", "77777777")

            p_pr = etree.SubElement(para, f"{{{w_ns}}}pPr")
            p_style = etree.SubElement(p_pr, f"{{{w_ns}}}pStyle")
            p_style.set(f"{{{w_ns}}}val", "CommentText")
            
            # Add annotation reference run
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            r_pr = etree.SubElement(run, f"{{{w_ns}}}rPr")
            r_style = etree.SubElement(r_pr, f"{{{w_ns}}}rStyle")
            r_style.set(f"{{{w_ns}}}val", "CommentReference")
            etree.SubElement(run, f"{{{w_ns}}}annotationRef")
            
            # Add comment text
            run = etree.SubElement(para, f"{{{w_ns}}}r")
            text_elem = etree.SubElement(run, f"{{{w_ns}}}t")
            text_elem.text = metadata["text"]
        
        return etree.tostring(
            comments, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_comments_extended(self) -> bytes:
        """Create word/commentsExtended.xml."""
        w15_ns = self.NAMESPACES["w15"]
        
        comments_ex = etree.Element(
            f"{{{w15_ns}}}commentsEx",
            nsmap=self.WORD_NAMESPACES,
        )
        comments_ex.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")
        
        # Add each comment extension
        for metadata in self._comment_metadata:
            comment_ex = etree.SubElement(
                comments_ex,
                f"{{{w15_ns}}}commentEx",
                {
                    f"{{{w15_ns}}}paraId": metadata["para_id"],
                    f"{{{w15_ns}}}done": "1" if metadata["resolved"] else "0",
                },
            )
            
            # Add parent reference for replies
            if metadata["parent_para_id"]:
                comment_ex.set(f"{{{w15_ns}}}paraIdParent", metadata["parent_para_id"])
        
        return etree.tostring(
            comments_ex, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_comments_ids(self) -> bytes:
        """Create word/commentsIds.xml."""
        w16cid_ns = self.NAMESPACES["w16cid"]
        
        comments_ids = etree.Element(
            f"{{{w16cid_ns}}}commentsIds",
            nsmap=self.WORD_NAMESPACES,
        )
        comments_ids.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")
        
        # Add each comment ID
        for metadata in self._comment_metadata:
            etree.SubElement(
                comments_ids,
                f"{{{w16cid_ns}}}commentId",
                {
                    f"{{{w16cid_ns}}}paraId": metadata["para_id"],
                    f"{{{w16cid_ns}}}durableId": metadata["durable_id"],
                },
            )
        
        return etree.tostring(
            comments_ids, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_settings(self) -> bytes:
        """Create word/settings.xml."""
        root = etree.fromstring(SETTINGS_XML.encode("utf-8"))
        w_ns = self.NAMESPACES["w"]
        has_even = any(
            section.headers.even is not None or section.footers.even is not None
            for section in self._section_layout
        )
        if has_even and root.find(f"{{{w_ns}}}evenAndOddHeaders") is None:
            root.insert(0, etree.Element(f"{{{w_ns}}}evenAndOddHeaders"))
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=False)

    def _create_web_settings(self) -> bytes:
        """Create word/webSettings.xml."""
        return WEB_SETTINGS_XML.encode("utf-8")

    def _create_footnotes(self) -> bytes:
        """Create word/footnotes.xml."""
        w_ns = self.NAMESPACES["w"]
        w14_ns = self.NAMESPACES["w14"]
        footnotes = etree.Element(
            f"{{{w_ns}}}footnotes",
            nsmap={"w": w_ns, "w14": w14_ns},
        )
        self._add_note_separator(footnotes, "footnote", "separator", "-1")
        self._add_note_separator(footnotes, "footnote", "continuationSeparator", "0")
        return etree.tostring(
            footnotes, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _create_endnotes(self) -> bytes:
        """Create word/endnotes.xml."""
        w_ns = self.NAMESPACES["w"]
        w14_ns = self.NAMESPACES["w14"]
        endnotes = etree.Element(
            f"{{{w_ns}}}endnotes",
            nsmap={"w": w_ns, "w14": w14_ns},
        )
        self._add_note_separator(endnotes, "endnote", "separator", "-1")
        self._add_note_separator(endnotes, "endnote", "continuationSeparator", "0")
        return etree.tostring(
            endnotes, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _add_note_separator(
        self, parent: XMLElement, tag: str, sep_tag: str, note_id: str
    ) -> None:
        """Add a note separator entry for footnotes or endnotes."""
        w_ns = self.NAMESPACES["w"]
        w14_ns = self.NAMESPACES["w14"]
        note = etree.SubElement(
            parent,
            f"{{{w_ns}}}{tag}",
            {f"{{{w_ns}}}type": sep_tag, f"{{{w_ns}}}id": note_id},
        )
        para = etree.SubElement(note, f"{{{w_ns}}}p")
        para.set(f"{{{w14_ns}}}paraId", self._generate_hex_id(8))
        para.set(f"{{{w14_ns}}}textId", "77777777")
        p_pr = etree.SubElement(para, f"{{{w_ns}}}pPr")
        spacing = etree.SubElement(p_pr, f"{{{w_ns}}}spacing")
        spacing.set(f"{{{w_ns}}}after", "0")
        spacing.set(f"{{{w_ns}}}line", "240")
        spacing.set(f"{{{w_ns}}}lineRule", "auto")
        run = etree.SubElement(para, f"{{{w_ns}}}r")
        etree.SubElement(run, f"{{{w_ns}}}{sep_tag}")

    def _create_font_table(self) -> bytes:
        """Create word/fontTable.xml."""
        return FONT_TABLE_XML.encode("utf-8")

    def _create_theme(self) -> bytes:
        """Create word/theme/theme1.xml."""
        return base64.b64decode(THEME_XML_B64)

    def _create_core_properties(self) -> bytes:
        """Create docProps/core.xml."""
        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        core_xml = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" "
            "xmlns:dcterms=\"http://purl.org/dc/terms/\" "
            "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
            "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
            f"<dc:title>{self.spec.title}</dc:title>"
            "<dc:subject></dc:subject>"
            f"<dc:creator>{self.spec.author}</dc:creator>"
            "<cp:keywords></cp:keywords>"
            "<dc:description></dc:description>"
            "<cp:lastModifiedBy></cp:lastModifiedBy>"
            "<cp:revision>1</cp:revision>"
            f"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created>"
            f"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified>"
            "</cp:coreProperties>"
        )
        return core_xml.encode("utf-8")

    def _create_app_properties(self) -> bytes:
        """Create docProps/app.xml."""
        return APP_PROPERTIES_XML.encode("utf-8")
    
    def _create_numbering(self, has_list_numbering: bool = True, has_heading_numbering: bool = False) -> bytes:
        """Create word/numbering.xml with multilevel numbering definitions."""
        w_ns = self.NAMESPACES["w"]
        w15_ns = self.NAMESPACES["w15"]
        w16cid_ns = self.NAMESPACES["w16cid"]

        # Create numbering element with all namespaces
        numbering = etree.Element(
            f"{{{w_ns}}}numbering",
            nsmap=self.WORD_NAMESPACES,
        )
        numbering.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")

        if has_list_numbering:
            self._add_list_abstract_num(numbering, w_ns, w15_ns)

        if has_heading_numbering:
            self._add_heading_abstract_num(numbering, w_ns, w15_ns)

        # Create concrete numbering instances
        if has_list_numbering:
            num = etree.SubElement(
                numbering,
                f"{{{w_ns}}}num",
                {f"{{{w_ns}}}numId": "1"},
            )
            num.set(f"{{{w16cid_ns}}}durableId", "283199500")
            etree.SubElement(num, f"{{{w_ns}}}abstractNumId", {f"{{{w_ns}}}val": "0"})

        if has_heading_numbering:
            abstract_id = "1" if has_list_numbering else "0"
            num_id = "2" if has_list_numbering else "1"
            num = etree.SubElement(
                numbering,
                f"{{{w_ns}}}num",
                {f"{{{w_ns}}}numId": num_id},
            )
            num.set(f"{{{w16cid_ns}}}durableId", "283199501")
            etree.SubElement(num, f"{{{w_ns}}}abstractNumId", {f"{{{w_ns}}}val": abstract_id})

        return etree.tostring(
            numbering, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _add_list_abstract_num(self, numbering: XMLElement, w_ns: str, w15_ns: str) -> None:
        """Add the legal-list abstract numbering definition (abstractNumId=0)."""
        abstractNum = etree.SubElement(
            numbering,
            f"{{{w_ns}}}abstractNum",
            {f"{{{w_ns}}}abstractNumId": "0"},
        )
        abstractNum.set(f"{{{w15_ns}}}restartNumberingAfterBreak", "0")

        etree.SubElement(abstractNum, f"{{{w_ns}}}nsid", {f"{{{w_ns}}}val": "355246F9"})
        etree.SubElement(abstractNum, f"{{{w_ns}}}multiLevelType", {f"{{{w_ns}}}val": "multilevel"})
        etree.SubElement(abstractNum, f"{{{w_ns}}}tmpl", {f"{{{w_ns}}}val": "2000001F"})

        level_formats = [
            "%1.", "%1.%2.", "%1.%2.%3.", "%1.%2.%3.%4.",
            "%1.%2.%3.%4.%5.", "%1.%2.%3.%4.%5.%6.",
            "%1.%2.%3.%4.%5.%6.%7.", "%1.%2.%3.%4.%5.%6.%7.%8.",
            "%1.%2.%3.%4.%5.%6.%7.%8.%9.",
        ]
        indents = [
            (360, 360), (792, 432), (1224, 504), (1728, 648),
            (2232, 792), (2736, 936), (3240, 1080), (3744, 1224),
            (4320, 1440),
        ]

        for i, (lvl_text, (left, hanging)) in enumerate(zip(level_formats, indents)):
            lvl = etree.SubElement(
                abstractNum, f"{{{w_ns}}}lvl",
                {f"{{{w_ns}}}ilvl": str(i)},
            )
            etree.SubElement(lvl, f"{{{w_ns}}}start", {f"{{{w_ns}}}val": "1"})
            etree.SubElement(lvl, f"{{{w_ns}}}numFmt", {f"{{{w_ns}}}val": "decimal"})
            etree.SubElement(lvl, f"{{{w_ns}}}lvlText", {f"{{{w_ns}}}val": lvl_text})
            etree.SubElement(lvl, f"{{{w_ns}}}lvlJc", {f"{{{w_ns}}}val": "left"})
            pPr = etree.SubElement(lvl, f"{{{w_ns}}}pPr")
            etree.SubElement(
                pPr, f"{{{w_ns}}}ind",
                {f"{{{w_ns}}}left": str(left), f"{{{w_ns}}}hanging": str(hanging)},
            )

    def _add_heading_abstract_num(self, numbering: XMLElement, w_ns: str, w15_ns: str) -> None:
        """Add the heading-style abstract numbering definition.

        Each level has a pStyle back-reference to HeadingN so that the
        numbering linkage lives in the style definitions rather than in
        each paragraph's numPr.
        """
        # Use abstractNumId=1 if list numbering already claimed 0, else 0
        existing = numbering.findall(f"{{{w_ns}}}abstractNum")
        abstract_id = str(len(existing))

        abstractNum = etree.SubElement(
            numbering,
            f"{{{w_ns}}}abstractNum",
            {f"{{{w_ns}}}abstractNumId": abstract_id},
        )
        abstractNum.set(f"{{{w15_ns}}}restartNumberingAfterBreak", "0")

        etree.SubElement(abstractNum, f"{{{w_ns}}}nsid", {f"{{{w_ns}}}val": "4A2E17B8"})
        etree.SubElement(abstractNum, f"{{{w_ns}}}multiLevelType", {f"{{{w_ns}}}val": "multilevel"})
        etree.SubElement(abstractNum, f"{{{w_ns}}}tmpl", {f"{{{w_ns}}}val": "0409001D"})

        # Level definitions: (numFmt, lvlText, pStyle, jc)
        heading_levels = [
            ("decimal", "%1", "Heading1", "left"),
            ("decimal", "%1.%2", "Heading2", "left"),
            ("lowerLetter", "(%3)", "Heading3", "left"),
            ("lowerRoman", "(%4)", "Heading4", "left"),
        ]
        indents = [(432, 432), (576, 576), (720, 720), (864, 864)]

        for i, ((num_fmt, lvl_text, p_style, jc), (left, hanging)) in enumerate(
            zip(heading_levels, indents)
        ):
            lvl = etree.SubElement(
                abstractNum, f"{{{w_ns}}}lvl",
                {f"{{{w_ns}}}ilvl": str(i)},
            )
            etree.SubElement(lvl, f"{{{w_ns}}}start", {f"{{{w_ns}}}val": "1"})
            etree.SubElement(lvl, f"{{{w_ns}}}numFmt", {f"{{{w_ns}}}val": num_fmt})
            etree.SubElement(lvl, f"{{{w_ns}}}pStyle", {f"{{{w_ns}}}val": p_style})
            etree.SubElement(lvl, f"{{{w_ns}}}lvlText", {f"{{{w_ns}}}val": lvl_text})
            etree.SubElement(lvl, f"{{{w_ns}}}lvlJc", {f"{{{w_ns}}}val": jc})
            pPr = etree.SubElement(lvl, f"{{{w_ns}}}pPr")
            etree.SubElement(
                pPr, f"{{{w_ns}}}ind",
                {f"{{{w_ns}}}left": str(left), f"{{{w_ns}}}hanging": str(hanging)},
            )
    
    def _create_styles(self, has_heading_numbering: bool = False) -> bytes:
        """Create word/styles.xml with minimal defaults, ListParagraph, and optional headings."""
        w_ns = self.NAMESPACES["w"]

        styles = etree.Element(
            f"{{{w_ns}}}styles",
            nsmap=self.WORD_NAMESPACES,
        )
        styles.set(f"{{{self.NAMESPACES['mc']}}}Ignorable", "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14")

        doc_defaults = etree.SubElement(styles, f"{{{w_ns}}}docDefaults")
        r_pr_default = etree.SubElement(doc_defaults, f"{{{w_ns}}}rPrDefault")
        r_pr = etree.SubElement(r_pr_default, f"{{{w_ns}}}rPr")
        r_fonts = etree.SubElement(r_pr, f"{{{w_ns}}}rFonts")
        r_fonts.set(f"{{{w_ns}}}asciiTheme", "minorHAnsi")
        r_fonts.set(f"{{{w_ns}}}eastAsiaTheme", "minorEastAsia")
        r_fonts.set(f"{{{w_ns}}}hAnsiTheme", "minorHAnsi")
        r_fonts.set(f"{{{w_ns}}}cstheme", "minorBidi")
        etree.SubElement(r_pr, f"{{{w_ns}}}kern", {f"{{{w_ns}}}val": "2"})
        etree.SubElement(r_pr, f"{{{w_ns}}}sz", {f"{{{w_ns}}}val": "24"})
        etree.SubElement(r_pr, f"{{{w_ns}}}szCs", {f"{{{w_ns}}}val": "24"})
        lang = etree.SubElement(r_pr, f"{{{w_ns}}}lang")
        lang.set(f"{{{w_ns}}}val", "en-CH")
        lang.set(f"{{{w_ns}}}eastAsia", "en-CH")
        lang.set(f"{{{w_ns}}}bidi", "ar-SA")

        p_pr_default = etree.SubElement(doc_defaults, f"{{{w_ns}}}pPrDefault")
        p_pr = etree.SubElement(p_pr_default, f"{{{w_ns}}}pPr")
        spacing = etree.SubElement(p_pr, f"{{{w_ns}}}spacing")
        spacing.set(f"{{{w_ns}}}after", "160")
        spacing.set(f"{{{w_ns}}}line", "278")
        spacing.set(f"{{{w_ns}}}lineRule", "auto")

        normal = etree.SubElement(
            styles,
            f"{{{w_ns}}}style",
            {
                f"{{{w_ns}}}type": "paragraph",
                f"{{{w_ns}}}default": "1",
                f"{{{w_ns}}}styleId": "Normal",
            },
        )
        etree.SubElement(normal, f"{{{w_ns}}}name", {f"{{{w_ns}}}val": "Normal"})
        etree.SubElement(normal, f"{{{w_ns}}}qFormat")

        default_para = etree.SubElement(
            styles,
            f"{{{w_ns}}}style",
            {
                f"{{{w_ns}}}type": "character",
                f"{{{w_ns}}}default": "1",
                f"{{{w_ns}}}styleId": "DefaultParagraphFont",
            },
        )
        etree.SubElement(
            default_para, f"{{{w_ns}}}name", {f"{{{w_ns}}}val": "Default Paragraph Font"}
        )
        etree.SubElement(default_para, f"{{{w_ns}}}uiPriority", {f"{{{w_ns}}}val": "1"})
        etree.SubElement(default_para, f"{{{w_ns}}}semiHidden")
        etree.SubElement(default_para, f"{{{w_ns}}}unhideWhenUsed")

        table_normal = etree.SubElement(
            styles,
            f"{{{w_ns}}}style",
            {
                f"{{{w_ns}}}type": "table",
                f"{{{w_ns}}}default": "1",
                f"{{{w_ns}}}styleId": "TableNormal",
            },
        )
        etree.SubElement(table_normal, f"{{{w_ns}}}name", {f"{{{w_ns}}}val": "Normal Table"})
        etree.SubElement(table_normal, f"{{{w_ns}}}uiPriority", {f"{{{w_ns}}}val": "99"})
        etree.SubElement(table_normal, f"{{{w_ns}}}semiHidden")
        etree.SubElement(table_normal, f"{{{w_ns}}}unhideWhenUsed")
        tbl_pr = etree.SubElement(table_normal, f"{{{w_ns}}}tblPr")
        etree.SubElement(tbl_pr, f"{{{w_ns}}}tblInd", {f"{{{w_ns}}}w": "0", f"{{{w_ns}}}type": "dxa"})
        tbl_cell_mar = etree.SubElement(tbl_pr, f"{{{w_ns}}}tblCellMar")
        etree.SubElement(tbl_cell_mar, f"{{{w_ns}}}top", {f"{{{w_ns}}}w": "0", f"{{{w_ns}}}type": "dxa"})
        etree.SubElement(tbl_cell_mar, f"{{{w_ns}}}left", {f"{{{w_ns}}}w": "108", f"{{{w_ns}}}type": "dxa"})
        etree.SubElement(tbl_cell_mar, f"{{{w_ns}}}bottom", {f"{{{w_ns}}}w": "0", f"{{{w_ns}}}type": "dxa"})
        etree.SubElement(tbl_cell_mar, f"{{{w_ns}}}right", {f"{{{w_ns}}}w": "108", f"{{{w_ns}}}type": "dxa"})

        no_list = etree.SubElement(
            styles,
            f"{{{w_ns}}}style",
            {
                f"{{{w_ns}}}type": "numbering",
                f"{{{w_ns}}}default": "1",
                f"{{{w_ns}}}styleId": "NoList",
            },
        )
        etree.SubElement(no_list, f"{{{w_ns}}}name", {f"{{{w_ns}}}val": "No List"})
        etree.SubElement(no_list, f"{{{w_ns}}}uiPriority", {f"{{{w_ns}}}val": "99"})
        etree.SubElement(no_list, f"{{{w_ns}}}semiHidden")
        etree.SubElement(no_list, f"{{{w_ns}}}unhideWhenUsed")

        style = etree.SubElement(
            styles,
            f"{{{w_ns}}}style",
            {
                f"{{{w_ns}}}type": "paragraph",
                f"{{{w_ns}}}styleId": "ListParagraph",
            },
        )
        etree.SubElement(style, f"{{{w_ns}}}name", {f"{{{w_ns}}}val": "List Paragraph"})
        etree.SubElement(style, f"{{{w_ns}}}basedOn", {f"{{{w_ns}}}val": "Normal"})
        etree.SubElement(style, f"{{{w_ns}}}uiPriority", {f"{{{w_ns}}}val": "34"})
        etree.SubElement(style, f"{{{w_ns}}}semiHidden")
        etree.SubElement(style, f"{{{w_ns}}}unhideWhenUsed")
        etree.SubElement(style, f"{{{w_ns}}}qFormat")
        p_pr = etree.SubElement(style, f"{{{w_ns}}}pPr")
        etree.SubElement(p_pr, f"{{{w_ns}}}contextualSpacing")

        if has_heading_numbering:
            self._add_heading_styles(styles, w_ns)

        return etree.tostring(
            styles, xml_declaration=True, encoding="UTF-8", pretty_print=True
        )

    def _add_heading_styles(self, styles: XMLElement, w_ns: str) -> None:
        """Add Heading1-Heading4 styles with embedded numPr for styled numbering.

        The numbering linkage lives inside the style: each heading style
        contains a <w:numPr> with the appropriate ilvl, and the matching
        abstractNum levels contain <w:pStyle> back-references.  Paragraphs
        only need ``<w:pStyle w:val="HeadingN"/>``, no explicit numPr.
        """
        # Determine the numId for heading numbering
        # If list numbering is also present, heading numbering uses numId=2
        has_list = any(p.numbering for p in self.spec.paragraphs)
        heading_num_id = "2" if has_list else "1"

        # (styleId, display name, ilvl, sz, bold, spacing_before, spacing_after)
        heading_defs = [
            ("Heading1", "heading 1", 0, "32", True, 240, 0),
            ("Heading2", "heading 2", 1, "28", True, 200, 0),
            ("Heading3", "heading 3", 2, "24", False, 160, 0),
            ("Heading4", "heading 4", 3, "22", False, 120, 0),
        ]

        for style_id, name, ilvl, sz, bold, sp_before, sp_after in heading_defs:
            style = etree.SubElement(
                styles,
                f"{{{w_ns}}}style",
                {f"{{{w_ns}}}type": "paragraph", f"{{{w_ns}}}styleId": style_id},
            )
            etree.SubElement(style, f"{{{w_ns}}}name", {f"{{{w_ns}}}val": name})
            etree.SubElement(style, f"{{{w_ns}}}basedOn", {f"{{{w_ns}}}val": "Normal"})
            etree.SubElement(style, f"{{{w_ns}}}next", {f"{{{w_ns}}}val": "Normal"})
            etree.SubElement(style, f"{{{w_ns}}}qFormat")

            # Paragraph properties with numbering reference
            pPr = etree.SubElement(style, f"{{{w_ns}}}pPr")
            etree.SubElement(pPr, f"{{{w_ns}}}keepNext")
            etree.SubElement(pPr, f"{{{w_ns}}}keepLines")

            numPr = etree.SubElement(pPr, f"{{{w_ns}}}numPr")
            if ilvl > 0:
                etree.SubElement(numPr, f"{{{w_ns}}}ilvl", {f"{{{w_ns}}}val": str(ilvl)})
            etree.SubElement(numPr, f"{{{w_ns}}}numId", {f"{{{w_ns}}}val": heading_num_id})

            spacing = etree.SubElement(pPr, f"{{{w_ns}}}spacing")
            spacing.set(f"{{{w_ns}}}before", str(sp_before))
            spacing.set(f"{{{w_ns}}}after", str(sp_after))

            etree.SubElement(
                pPr, f"{{{w_ns}}}outlineLvl", {f"{{{w_ns}}}val": str(ilvl)}
            )

            # Run properties (font size, bold)
            rPr = etree.SubElement(style, f"{{{w_ns}}}rPr")
            rFonts = etree.SubElement(rPr, f"{{{w_ns}}}rFonts")
            rFonts.set(f"{{{w_ns}}}asciiTheme", "majorHAnsi")
            rFonts.set(f"{{{w_ns}}}eastAsiaTheme", "majorEastAsia")
            rFonts.set(f"{{{w_ns}}}hAnsiTheme", "majorHAnsi")
            rFonts.set(f"{{{w_ns}}}cstheme", "majorBidi")
            if bold:
                etree.SubElement(rPr, f"{{{w_ns}}}b")
                etree.SubElement(rPr, f"{{{w_ns}}}bCs")
            etree.SubElement(rPr, f"{{{w_ns}}}sz", {f"{{{w_ns}}}val": sz})
            etree.SubElement(rPr, f"{{{w_ns}}}szCs", {f"{{{w_ns}}}val": sz})
