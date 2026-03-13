"""
RDL Analyzer 獨立測試腳本
- 不需要 Django、不需要開 GUI
- 直接在此檔案頂端設定 XML 內容與頁面參數
- 用 logging 逐步確認每個值有正確讀到
執行方式：
    python test_rdl_analyzer.py
"""

import logging
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Optional, List, Tuple

# ─────────────────────────────────────────────────────────
# ★ 在這裡設定你的測試 INPUT ★
# ─────────────────────────────────────────────────────────

# 頁面設定（單位：cm）
TEST_PAGE_W   = 21.0   # 頁寬
TEST_PAGE_H   = 29.7   # 頁高
TEST_MARGIN_L = 1.27   # 左邊距
TEST_MARGIN_R = 1.27   # 右邊距
TEST_MARGIN_T = 1.27   # 上邊距
TEST_MARGIN_B = 1.27   # 下邊距

# 測試用 XML（模擬一份簡單的 RDL，Body 寬度超出以觸發 R1）
TEST_XML = """<?xml version="1.0" encoding="utf-8"?>
<Report MustUnderstand="df" xmlns="[http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition](http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition)" xmlns:rd="[http://schemas.microsoft.com/SQLServer/reporting/reportdesigner](http://schemas.microsoft.com/SQLServer/reporting/reportdesigner)" xmlns:df="[http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily](http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily)">
<df:DefaultFontFamily>Segoe UI</df:DefaultFontFamily>
<AutoRefresh>0</AutoRefresh>
<ReportSections>
<ReportSection>
<Body>
<ReportItems>
<Tablix Name="Tablix1">
<TablixBody>
<TablixColumns>
<TablixColumn>
<Width>2.5cm</Width>
</TablixColumn>
<TablixColumn>
<Width>2.5cm</Width>
</TablixColumn>
<TablixColumn>
<Width>2.5cm</Width>
</TablixColumn>
</TablixColumns>
<TablixRows>
<TablixRow>
<Height>0.6cm</Height>
<TablixCells>
<TablixCell>
<CellContents>
<Textbox Name="Textbox1">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox1</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox3">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox3</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox5">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox5</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
</TablixCells>
</TablixRow>
<TablixRow>
<Height>0.6cm</Height>
<TablixCells>
<TablixCell>
<CellContents>
<Textbox Name="Textbox2">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox2</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox4">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox4</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox6">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox6</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
</TablixCells>
</TablixRow>
</TablixRows>
</TablixBody>
<TablixColumnHierarchy>
<TablixMembers>
<TablixMember />
<TablixMember />
<TablixMember />
</TablixMembers>
</TablixColumnHierarchy>
<TablixRowHierarchy>
<TablixMembers>
<TablixMember>
<KeepWithGroup>After</KeepWithGroup>
</TablixMember>
<TablixMember>
<Group Name="詳細資料" />
</TablixMember>
</TablixMembers>
</TablixRowHierarchy>
<Top>0.65828cm</Top>
<Height>1.2cm</Height>
<Width>7.5cm</Width>
<Style>
<Border>
<Style>None</Style>
</Border>
</Style>
</Tablix>
<Tablix Name="Tablix2">
<TablixBody>
<TablixColumns>
<TablixColumn>
<Width>2.35995cm</Width>
</TablixColumn>
<TablixColumn>
<Width>2.35995cm</Width>
</TablixColumn>
<TablixColumn>
<Width>2.35995cm</Width>
</TablixColumn>
</TablixColumns>
<TablixRows>
<TablixRow>
<Height>0.6cm</Height>
<TablixCells>
<TablixCell>
<CellContents>
<Textbox Name="Textbox7">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox7</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox9">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox9</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox11">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox11</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
</TablixCells>
</TablixRow>
<TablixRow>
<Height>0.6cm</Height>
<TablixCells>
<TablixCell>
<CellContents>
<Textbox Name="Textbox8">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox8</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox10">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox10</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox12">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox12</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
</TablixCells>
</TablixRow>
</TablixRows>
</TablixBody>
<TablixColumnHierarchy>
<TablixMembers>
<TablixMember />
<TablixMember />
<TablixMember />
</TablixMembers>
</TablixColumnHierarchy>
<TablixRowHierarchy>
<TablixMembers>
<TablixMember>
<KeepWithGroup>After</KeepWithGroup>
</TablixMember>
<TablixMember>
<Group Name="詳細資料1" />
</TablixMember>
</TablixMembers>
</TablixRowHierarchy>
<Top>4.8387cm</Top>
<Left>0.42016cm</Left>
<Height>1.2cm</Height>
<Width>7.07984cm</Width>
<ZIndex>1</ZIndex>
<Style>
<Border>
<Style>None</Style>
</Border>
</Style>
</Tablix>
<Tablix Name="Tablix3">
<TablixBody>
<TablixColumns>
<TablixColumn>
<Width>3.4525cm</Width>
</TablixColumn>
<TablixColumn>
<Width>1.5475cm</Width>
</TablixColumn>
<TablixColumn>
<Width>2.5cm</Width>
</TablixColumn>
</TablixColumns>
<TablixRows>
<TablixRow>
<Height>0.6cm</Height>
<TablixCells>
<TablixCell>
<CellContents>
<Textbox Name="Textbox13">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox7</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox14">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox9</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox15">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox11</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
</TablixCells>
</TablixRow>
<TablixRow>
<Height>0.6cm</Height>
<TablixCells>
<TablixCell>
<CellContents>
<Textbox Name="Textbox16">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox8</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox17">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox10</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
<TablixCell>
<CellContents>
<Textbox Name="Textbox18">
<CanGrow>true</CanGrow>
<KeepTogether>true</KeepTogether>
<Paragraphs>
<Paragraph>
<TextRuns>
<TextRun>
<Value />
<Style />
</TextRun>
</TextRuns>
<Style />
</Paragraph>
</Paragraphs>
<rd:DefaultName>Textbox12</rd:DefaultName>
<Style>
<Border>
<Color>LightGrey</Color>
<Style>Solid</Style>
</Border>
<PaddingLeft>2pt</PaddingLeft>
<PaddingRight>2pt</PaddingRight>
<PaddingTop>2pt</PaddingTop>
<PaddingBottom>2pt</PaddingBottom>
</Style>
</Textbox>
</CellContents>
</TablixCell>
</TablixCells>
</TablixRow>
</TablixRows>
</TablixBody>
<TablixColumnHierarchy>
<TablixMembers>
<TablixMember />
<TablixMember />
<TablixMember />
</TablixMembers>
</TablixColumnHierarchy>
<TablixRowHierarchy>
<TablixMembers>
<TablixMember>
<KeepWithGroup>After</KeepWithGroup>
</TablixMember>
<TablixMember>
<Group Name="詳細資料2" />
</TablixMember>
</TablixMembers>
</TablixRowHierarchy>
<Top>2.74698cm</Top>
<Left>0cm</Left>
<Height>1.2cm</Height>
<Width>7.5cm</Width>
<ZIndex>2</ZIndex>
<Style>
<Border>
<Style>None</Style>
</Border>
</Style>
</Tablix>
</ReportItems>
<Height>3.84375in</Height>
<Style />
</Body>
<Width>6.5in</Width>
<Page>
<PageHeight>29.7cm</PageHeight>
<PageWidth>21cm</PageWidth>
<LeftMargin>2cm</LeftMargin>
<RightMargin>2cm</RightMargin>
<TopMargin>2cm</TopMargin>
<BottomMargin>2cm</BottomMargin>
<ColumnSpacing>0.13cm</ColumnSpacing>
<Style />
</Page>
</ReportSection>
</ReportSections>
<ReportParametersLayout>
<GridLayoutDefinition>
<NumberOfColumns>4</NumberOfColumns>
<NumberOfRows>2</NumberOfRows>
</GridLayoutDefinition>
</ReportParametersLayout>
<rd:ReportUnitType>Cm</rd:ReportUnitType>
<rd:ReportID>68832d2a-82d2-4ae8-86d9-f96d127387f5</rd:ReportID>
</Report>
"""

# ─────────────────────────────────────────────────────────
# LOGGING 設定（同時輸出到 console 和 .log 檔）
# ─────────────────────────────────────────────────────────




# 清除殘留 handler
root_logger = logging.getLogger()
for h in root_logger.handlers[:]:
    root_logger.removeHandler(h)
    h.close()

# 重新設定（只用 StreamHandler + FileHandler，不用 Socket）
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("test_rdl_result.log", encoding="utf-8", mode="w"),
    ],
)
log = logging.getLogger("rdl_test")

# ─────────────────────────────────────────────────────────
# 從原本的程式碼複製過來的核心類別（不動 GUI 部分）
# ─────────────────────────────────────────────────────────

SEVERITY_ORDER = {"critical": 0, "warning": 1, "info": 2}

ITEM_TAGS = [
    "Textbox", "Rectangle", "Tablix", "Table", "Matrix",
    "List", "Image", "Chart", "Gauge", "Map", "Subreport", "CustomReportItem",
]

def parse_size(val: Optional[str]) -> Optional[float]:
    if not val:
        return None
    s = val.strip().lower()
    try:
        if s.endswith("in"):  return float(s[:-2]) * 2.54
        if s.endswith("cm"):  return float(s[:-2])
        if s.endswith("mm"):  return float(s[:-2]) / 10
        if s.endswith("pt"):  return float(s[:-2]) * 0.0352778
        if s.endswith("px"):  return float(s[:-2]) * 0.0264583
        return float(s)
    except ValueError:
        return None


@dataclass
class Issue:
    rule_id: int
    severity: str
    category: str
    title: str
    detail: str
    fix_hint: str
    fixable: bool = False


class RDLDocument:
    def __init__(self, xml_string: str):
        self.raw = xml_string
        self.tree = ET.ElementTree(ET.fromstring(xml_string))
        self.root = self.tree.getroot()
        m = re.match(r"\{([^}]+)\}", self.root.tag)
        self.ns = m.group(1) if m else ""

    def _tag(self, name: str) -> str:
        return f"{{{self.ns}}}{name}" if self.ns else name

    def find(self, el, *tags):
        cur = el
        for tag in tags:
            cur = cur.find(self._tag(tag))
            if cur is None:
                return None
        return cur

    def findall(self, el, tag: str):
        return el.findall(".//" + self._tag(tag))

    def findall_direct(self, parent, tag: str):
        return [c for c in parent if c.tag == self._tag(tag)]

    def text(self, el, *tags) -> Optional[str]:
        node = self.find(el, *tags)
        return node.text.strip() if node is not None and node.text else None

    def sz(self, el, *tags) -> Optional[float]:
        return parse_size(self.text(el, *tags))


# ─────────────────────────────────────────────────────────
# 測試主流程
# ─────────────────────────────────────────────────────────

def run_test():
    log.info("=" * 55)
    log.info("RDL Analyzer 測試開始")
    log.info("=" * 55)

    # ── Step 1: 確認頁面參數有讀到 ──────────────────────────
    log.info("[ Step 1 ] 頁面設定參數")
    log.info(f"  頁寬           : {TEST_PAGE_W} cm")
    log.info(f"  頁高           : {TEST_PAGE_H} cm")
    log.info(f"  左/右邊距      : {TEST_MARGIN_L} / {TEST_MARGIN_R} cm")
    log.info(f"  上/下邊距      : {TEST_MARGIN_T} / {TEST_MARGIN_B} cm")
    max_w = TEST_PAGE_W - TEST_MARGIN_L - TEST_MARGIN_R
    max_h = TEST_PAGE_H - TEST_MARGIN_T - TEST_MARGIN_B
    log.info(f"  允許最大寬度   : {max_w:.3f} cm")
    log.info(f"  允許最大高度   : {max_h:.3f} cm")

    # ── Step 2: 解析 XML ────────────────────────────────────
    log.info("-" * 55)
    log.info("[ Step 2 ] 解析 XML")
    try:
        rdl = RDLDocument(TEST_XML)
        log.info(f"  XML 解析成功")
        log.info(f"  根節點 tag    : {rdl.root.tag}")
        log.info(f"  命名空間 (ns) : {rdl.ns if rdl.ns else '（無）'}")
    except Exception as e:
        log.error(f"  XML 解析失敗：{e}")
        return

    # ── Step 3: 讀取 Body ───────────────────────────────────
    log.info("-" * 55)
    log.info("[ Step 3 ] 讀取 Body")
    body_list = rdl.findall(rdl.root, "Body")
    body = body_list[0] if body_list else None
    if body is None:
        log.error("  找不到 Body 節點！")
        return

    bw = rdl.sz(body, "Width")
    bh = rdl.sz(body, "Height")
    log.info(f"  Body.Width    : {bw} cm  →  {'⚠ 超出' if bw and bw > max_w else '✓ 正常'} (允許 {max_w:.3f}cm)")
    log.info(f"  Body.Height   : {bh} cm")

    # ── Step 4: 讀取所有 ReportItems ────────────────────────
    log.info("-" * 55)
    log.info("[ Step 4 ] 掃描 ReportItems")
    body_ri = rdl.find(body, "ReportItems")
    if body_ri is None:
        log.warning("  Body 內沒有 ReportItems")
    else:
        for tag in ITEM_TAGS:
            for el in rdl.findall_direct(body_ri, tag):
                name  = el.get("Name", "(未命名)")
                left  = rdl.sz(el, "Left")  or 0
                top   = rdl.sz(el, "Top")   or 0
                width = rdl.sz(el, "Width") or 0
                height= rdl.sz(el, "Height")or 0
                r_edge = left + width
                status = "⚠ 超出右邊界" if r_edge > max_w + 0.005 else "✓"
                log.info(
                    f"  [{tag}] {name:<15} "
                    f"Left={left:.2f} Top={top:.2f} "
                    f"W={width:.2f} H={height:.2f}  "
                    f"右邊緣={r_edge:.3f}cm  {status}"
                )

    # ── Step 5: 讀取 Tablix 欄寬 ────────────────────────────
    # ── Step 5: 讀取 Tablix 欄寬 ────────────────────────────
    log.info("-" * 55)
    log.info("[ Step 5 ] Tablix 欄寬檢查")
    tablixes = rdl.findall(rdl.root, "Tablix")
    tablix_boundaries = []

    for tablix in tablixes:
        name  = tablix.get("Name", "(未命名)")
        tw    = rdl.sz(tablix, "Width") or 0
        cols  = rdl.findall(tablix, "TablixColumn")
        col_w = sum(rdl.sz(c, "Width") or 0 for c in cols)
        diff  = abs(col_w - tw)
        status = f"⚠ 差 {diff:.3f}cm" if diff > 0.005 else "✓ 一致"
        log.info(f"  Tablix [{name}]  Width={tw:.3f}cm  欄寬總和={col_w:.3f}cm  {status}")
        for i, c in enumerate(cols):
            log.debug(f"    第{i+1}欄寬 = {rdl.sz(c, 'Width') or 0:.3f}cm")

        # 累計邊界（含 Left）
        acc    = rdl.sz(tablix, "Left") or 0
        bounds = []
        for c in cols:
            acc += rdl.sz(c, "Width") or 0
            bounds.append(round(acc, 5))
        tablix_boundaries.append((name, bounds))

    # ── Step 5b: Left 起點對齊 ──────────────────────────────
    log.info("-" * 55)
    log.info("[ Step 5b ] Tablix 左起點對齊檢查")
    left_vals = [(t.get("Name", "(未命名)"), rdl.sz(t, "Left") or 0) for t in tablixes]
    for name, l in left_vals:
        log.info(f"  Tablix [{name}]  Left={l:.5f}cm")
    min_l = min(v for _, v in left_vals)
    max_l = max(v for _, v in left_vals)
    if max_l - min_l > 0.05:
        log.warning(f"  ⚠ 左起點不一致！最大差值 {max_l - min_l:.5f}cm")
    else:
        log.info(f"  ✓ 所有 Tablix 左起點一致")

    # ── Step 5c: 欄位垂直線對齊 ────────────────────────────
    log.info("-" * 55)
    log.info("[ Step 5c ] Tablix 欄位垂直線對齊檢查")
    for i in range(len(tablix_boundaries)):
        for j in range(i + 1, len(tablix_boundaries)):
            na, ba = tablix_boundaries[i]
            nb, bb = tablix_boundaries[j]
            if len(ba) != len(bb):
                log.warning(f"  ⚠ [{na}] 與 [{nb}] 欄數不同（{len(ba)} vs {len(bb)}），無法比對")
                continue
            for col_idx, (ca, cb) in enumerate(zip(ba, bb)):
                if abs(ca - cb) > 0.05:
                    log.warning(f"  ⚠ [{na}] 與 [{nb}] 第{col_idx+1}條垂直線不對齊：{ca:.5f}cm vs {cb:.5f}cm，差={abs(ca-cb):.5f}cm")
                else:
                    log.info(f"  ✓ [{na}] 與 [{nb}] 第{col_idx+1}條垂直線對齊：{ca:.5f}cm")

    # ── Step 6: 讀取 Page 節點驗證頁面尺寸 ─────────────────
    log.info("-" * 55)
    log.info("[ Step 6 ] Page 節點頁面尺寸驗證")
    rp_list = rdl.findall(rdl.root, "Page")
    rp = rp_list[0] if rp_list else None
    if rp is None:
        rp_list = rdl.findall(rdl.root, "ReportPage")
        rp = rp_list[0] if rp_list else None
        log.warning("  找不到 Page 節點")
    if rp is not None:
        rpw = rdl.sz(rp, "PageWidth")
        rph = rdl.sz(rp, "PageHeight")
        log.info(f"  RDL PageWidth  : {rpw} cm  →  {'⚠ 與設定不符' if rpw and abs(rpw - TEST_PAGE_W) > 0.05 else '✓ 符合'}")
        log.info(f"  RDL PageHeight : {rph} cm  →  {'⚠ 與設定不符' if rph and abs(rph - TEST_PAGE_H) > 0.05 else '✓ 符合'}")
   

    # ── Step 7: 完成 ────────────────────────────────────────
    log.info("=" * 55)
    log.info("測試完成，請查看上方 LOG 確認各數值是否正確")
    log.info("詳細記錄也已存至 test_rdl_result.log")
    log.info("=" * 55)


if __name__ == "__main__":
    run_test()
