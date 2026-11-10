// CmdShape.cpp — 핸들(노란 조절점) 달린 사용자 정의 도형 삽입 (plan/29).
//
// Excel COM/C API에는 사용자 정의 형상(a:custGeom)을 만드는 함수가 없으므로
// 가이드의 방식 A2를 쓴다: 도형 XML 하나만 든 최소 .xlsx를 임시 폴더에 쓰고
// (ZIP은 무압축 'stored'로 직접 기록 — 외부 라이브러리 없음), 숨김으로 열어
// Shape.Copy → 활성 시트 Paste → 닫고 삭제한다. 도형은 활성 셀의 좌상단에
// 도형별 기본 크기로 놓인다(셀 맞춤 아님).
//
// 조절값 규칙(가이드): 높이의 1/100000. 수식의 길이 상수는 EMU.
// 주의: 붙여넣기 때문에 사용자 클립보드 내용이 바뀐다.

#include "Commands.h"
#include "CmdCommon.h"
#include "../../core/ShapeGrid.h"

#include <cstdint>
#include <filesystem>
#include <fstream>
#include <functional>
#include <string>
#include <utility>
#include <vector>

namespace egtools::commands
{
    namespace
    {
        using S = std::string;

        // ── 최소 ZIP(stored) 기록기 ────────────────────────────────────────
        uint32_t crc32(const S& d)
        {
            static uint32_t table[256];
            static bool init = false;
            if (!init)
            {
                for (uint32_t i = 0; i < 256; ++i)
                {
                    uint32_t c = i;
                    for (int k = 0; k < 8; ++k) c = (c & 1) ? (0xEDB88320u ^ (c >> 1)) : (c >> 1);
                    table[i] = c;
                }
                init = true;
            }
            uint32_t c = 0xFFFFFFFFu;
            for (unsigned char ch : d) c = table[(c ^ ch) & 0xFF] ^ (c >> 8);
            return c ^ 0xFFFFFFFFu;
        }
        void le16(S& o, uint32_t v) { o += (char)(v & 0xFF); o += (char)((v >> 8) & 0xFF); }
        void le32(S& o, uint32_t v) { le16(o, v & 0xFFFF); le16(o, v >> 16); }

        S buildZip(const std::vector<std::pair<S, S>>& parts)
        {
            constexpr uint32_t kTime = 0, kDate = ((2026 - 1980) << 9) | (1 << 5) | 1;
            S out, central;
            for (const auto& [name, data] : parts)
            {
                const uint32_t off = (uint32_t)out.size(), crc = crc32(data);
                const uint32_t sz = (uint32_t)data.size(), nl = (uint32_t)name.size();
                le32(out, 0x04034b50); le16(out, 20); le16(out, 0); le16(out, 0);
                le16(out, kTime); le16(out, kDate); le32(out, crc); le32(out, sz); le32(out, sz);
                le16(out, nl); le16(out, 0); out += name; out += data;

                le32(central, 0x02014b50); le16(central, 20); le16(central, 20); le16(central, 0);
                le16(central, 0); le16(central, kTime); le16(central, kDate);
                le32(central, crc); le32(central, sz); le32(central, sz);
                le16(central, nl); le16(central, 0); le16(central, 0); le16(central, 0);
                le16(central, 0); le32(central, 0); le32(central, off); central += name;
            }
            const uint32_t cdOff = (uint32_t)out.size();
            out += central;
            le32(out, 0x06054b50); le16(out, 0); le16(out, 0);
            le16(out, (uint32_t)parts.size()); le16(out, (uint32_t)parts.size());
            le32(out, (uint32_t)central.size()); le32(out, cdOff); le16(out, 0);
            return out;
        }

        // ── DrawingML 조립 도우미 ──────────────────────────────────────────
        S gd(const char* n, const S& f)
        { return S("<a:gd name=\"") + n + "\" fmla=\"" + f + "\"/>"; }
        S pt(const char* x, const char* y)
        { return S("<a:pt x=\"") + x + "\" y=\"" + y + "\"/>"; }
        S mv(const char* x, const char* y) { return "<a:moveTo>" + pt(x, y) + "</a:moveTo>"; }
        S ln(const char* x, const char* y) { return "<a:lnTo>" + pt(x, y) + "</a:lnTo>"; }
        S arc(const char* wR, const char* hR, const char* st, const char* sw)
        {
            return S("<a:arcTo wR=\"") + wR + "\" hR=\"" + hR + "\" stAng=\"" + st +
                   "\" swAng=\"" + sw + "\"/>";
        }
        // 타원 한 바퀴(시계 방향 4분호). 시작점 = 왼쪽 끝 (cx-wR, cy).
        S fullEllipse(const char* wR, const char* hR)
        {
            return arc(wR, hR, "cd2", "cd4") + arc(wR, hR, "3cd4", "cd4") +
                   arc(wR, hR, "0", "cd4") + arc(wR, hR, "cd4", "cd4");
        }
        // 연결선 접점. ang = 연결선이 도형에서 나가는 방향(3cd4 위, cd2 왼쪽, cd4 아래, 0 오른쪽).
        // 접점 순서가 곧 번호이므로 바꾸면 이미 연결된 선이 다른 접점으로 붙는다.
        S cn(const char* ang, const char* x, const char* y)
        {
            return S("<a:cxn ang=\"") + ang + "\"><a:pos x=\"" + x + "\" y=\"" + y + "\"/></a:cxn>";
        }
        S geom(const S& av, const S& gds, const S& ah, const S& path, const S& cxns = S())
        {
            return "<a:custGeom><a:avLst>" + av + "</a:avLst><a:gdLst>" + gds +
                   "</a:gdLst><a:ahLst>" + ah +
                   "</a:ahLst><a:cxnLst>" + cxns + "</a:cxnLst><a:rect l=\"l\" t=\"t\" r=\"r\" b=\"b\"/>"
                   "<a:pathLst><a:path>" + path + "</a:path></a:pathLst></a:custGeom>";
        }

        // 한쪽 반원 사각형: 왼쪽 평평, 오른쪽 곡선(깊이 = h·adj, 최대 h/2 = 정반원).
        // 곡선은 폭이 줄어도 모양을 유지한다: 폭 < 깊이이면 반타원의 왼쪽을 도형
        // 경계(l)에서 잘라 그린다(직선부 0, 왼쪽 변이 곡선 끝점 사이 직선).
        //  xc = r−d(타원 중심 x), dxl = max(0, d−w)(경계가 중심에서 벗어난 거리),
        //  dyv = 경계 위치의 반높이, a = at2(dxl, dyv)(끝점의 시각 각도),
        //  시작 각도 = 360°−a, 스윕 = 2a (dxl=0이면 270°→+180° = 정반원).
        S geomOneSide()
        {
            return geom(
                gd("adj", "val 50000"),
                gd("a0", "pin 0 adj 50000") + gd("d", "*/ h a0 100000") +
                gd("dd", "max d 1") + gd("xc", "+- r 0 d") + gd("xs", "max xc l") +
                gd("dw", "+- d 0 w") + gd("dxl", "max 0 dw") + gd("t1", "*/ hd2 dxl dd") +
                gd("p1", "*/ hd2 hd2 1") + gd("p2", "*/ t1 t1 1") + gd("p3", "+- p1 0 p2") +
                gd("dyv", "sqrt p3") + gd("ys", "+- vc 0 dyv") + gd("ye", "+- vc dyv 0") +
                gd("ang", "at2 dxl dyv") + gd("st", "+- 21600000 0 ang") +
                gd("sw", "*/ ang 2 1") + gd("xt", "+/ l xs 2"),
                "<a:ahXY gdRefX=\"adj\" minX=\"0\" maxX=\"50000\"><a:pos x=\"xs\" y=\"vc\"/></a:ahXY>",
                mv("l", "ys") + ln("xs", "ys") + arc("d", "hd2", "st", "sw") +
                ln("l", "ye") + "<a:close/>",
                // 사각형 기준 4점: 위·아래는 직선부 중앙(폭이 좁아 직선부가 없으면 곡선 끝점)
                cn("3cd4", "xt", "ys") + cn("cd2", "l", "vc") + cn("cd4", "xt", "ye") + cn("0", "r", "vc"));
        }

        // 도넛형: 구멍 반경 = min(h·adj, ss/2 − 0.25pt).
        S geomDonut()
        {
            return geom(
                gd("adj", "val 20000"),
                gd("ssd2", "*/ ss 1 2") + gd("maxR", "+- ssd2 0 3175") +
                gd("maxAdj", "*/ maxR 100000 h") + gd("a", "pin 0 adj maxAdj") +
                gd("rh", "*/ h a 100000") + gd("xh", "+- hc rh 0") + gd("xhl", "+- hc 0 rh") +
                gd("yht", "+- vc 0 rh") + gd("yhb", "+- vc rh 0"),
                "<a:ahXY gdRefX=\"adj\" minX=\"0\" maxX=\"maxAdj\"><a:pos x=\"xh\" y=\"vc\"/></a:ahXY>",
                mv("l", "vc") + fullEllipse("wd2", "hd2") + "<a:close/>" +
                mv("xh", "vc") + arc("rh", "rh", "0", "-5400000") +
                arc("rh", "rh", "3cd4", "-5400000") + arc("rh", "rh", "cd2", "-5400000") +
                arc("rh", "rh", "cd4", "-5400000") + "<a:close/>",
                // 바깥 원 4점 + 구멍 원 4점(상·좌·하·우)
                cn("3cd4", "hc", "t") + cn("cd2", "l", "vc") + cn("cd4", "hc", "b") + cn("0", "r", "vc") +
                cn("cd4", "hc", "yht") + cn("0", "xhl", "vc") + cn("3cd4", "hc", "yhb") + cn("cd2", "xh", "vc"));
        }

        // 탭 달린 전극 패드: 위쪽 중앙 탭(폭 adj1, 길이 adj2) + 모서리 R(adj3) 패드.
        S geomTabPad()
        {
            return geom(
                gd("adj1", "val 20000") + gd("adj2", "val 30000") + gd("adj3", "val 10000"),
                gd("a2", "pin 0 adj2 70000") + gd("tl", "*/ h a2 100000") +
                gd("bh", "+- h 0 tl") + gd("mind", "min w bh") + gd("maxR", "*/ mind 1 2") +
                gd("maxA3", "*/ maxR 100000 h") + gd("a3", "pin 0 adj3 maxA3") +
                gd("rr", "*/ h a3 100000") + gd("wm", "+- w 0 rr") + gd("wm2", "+- wm 0 rr") +
                gd("maxA1", "*/ wm2 100000 h") + gd("a1", "pin 0 adj1 maxA1") +
                gd("tw", "*/ h a1 100000") + gd("twd2", "*/ tw 1 2") +
                gd("x1", "+- hc 0 twd2") + gd("x2", "+- hc twd2 0") +
                gd("yb", "+- t tl 0") + gd("yb2", "+- yb rr 0") + gd("yb3", "+- b 0 rr") +
                gd("xr", "+- r 0 rr") + gd("xl", "+- l rr 0") + gd("ypm", "+/ yb b 2"),
                "<a:ahXY gdRefX=\"adj1\" minX=\"0\" maxX=\"maxA1\"><a:pos x=\"x2\" y=\"t\"/></a:ahXY>"
                "<a:ahXY gdRefY=\"adj2\" minY=\"0\" maxY=\"70000\"><a:pos x=\"hc\" y=\"yb\"/></a:ahXY>"
                "<a:ahXY gdRefY=\"adj3\" minY=\"0\" maxY=\"maxA3\"><a:pos x=\"l\" y=\"yb2\"/></a:ahXY>",
                mv("x1", "t") + ln("x2", "t") + ln("x2", "yb") + ln("xr", "yb") +
                arc("rr", "rr", "3cd4", "cd4") + ln("r", "yb3") +
                arc("rr", "rr", "0", "cd4") + ln("xl", "b") +
                arc("rr", "rr", "cd4", "cd4") + ln("l", "yb2") +
                arc("rr", "rr", "cd2", "cd4") + ln("x1", "yb") + "<a:close/>",
                // 위 = 꼭지 끝, 좌·우 = 패드 사각형 중앙, 아래 = 패드 아래변 중앙
                cn("3cd4", "hc", "t") + cn("cd2", "l", "ypm") + cn("cd4", "hc", "b") + cn("0", "r", "ypm"));
        }

        // 벌어진 홈 사각형: 바깥 사각형(왼쪽 두 모서리 R=adj1) 오른쪽에서 왼쪽으로 홈이
        // 파이고(깊이 = w·adj2, 홈 높이 = 2·h·adj5, 홈 안쪽 끝 모서리 R=adj3),
        // 홈 입구는 바깥 위·아래 오른쪽 꼭지점에서 사분 타원호(가로 = w·adj4)로 벌어진다.
        // 핸들: 1 바깥 왼쪽 R / 2 홈 깊이 / 3 홈 안쪽 끝 R / 4 입구 호 길이(가로)+홈 높이(세로, adj5) — 3·4번 사이는 항상 수평 직선.
        S geomSlot()
        {
            return geom(
                gd("adj1", "val 15000") + gd("adj2", "val 72000") +
                gd("adj3", "val 14000") + gd("adj4", "val 30000") +
                gd("adj5", "val 18000"),
                gd("hm", "+- hd2 0 12700") + gd("maxA5", "*/ hm 100000 h") +
                gd("a5", "pin 2000 adj5 maxA5") + gd("sh", "*/ h a5 100000") + gd("hr", "+- hd2 0 sh") +
                gd("mr1", "min wd2 hd2") + gd("maxA1", "*/ mr1 100000 h") +
                gd("a1", "pin 0 adj1 maxA1") + gd("r1", "*/ h a1 100000") +
                gd("maxA3", "*/ sh 100000 h") + gd("a3", "pin 0 adj3 maxA3") +
                gd("r3", "*/ h a3 100000") + gd("minA2", "*/ r3 100000 w") +
                gd("wm1", "+- w 0 12700") + gd("maxA2", "*/ wm1 100000 w") +
                gd("a2", "pin minA2 adj2 maxA2") + gd("dd", "*/ w a2 100000") +
                gd("xe", "+- r 0 dd") + gd("ddr", "+- dd 0 r3") +
                gd("maxA4", "*/ ddr 100000 w") + gd("a4", "pin 0 adj4 maxA4") +
                gd("fl", "*/ w a4 100000") + gd("x4", "+- r 0 fl") +
                gd("yt", "+- vc 0 sh") + gd("yb", "+- vc sh 0") +
                gd("xs3", "+- xe r3 0") + gd("yt3", "+- yt r3 0") + gd("yb3", "+- yb 0 r3") +
                gd("y1", "+- t r1 0") + gd("xr1", "+- l r1 0") + gd("yb1", "+- b 0 r1") +
                gd("xtm", "+/ xr1 r 2"),
                "<a:ahXY gdRefX=\"adj1\" minX=\"0\" maxX=\"maxA1\"><a:pos x=\"xr1\" y=\"t\"/></a:ahXY>"
                "<a:ahXY gdRefX=\"adj2\" minX=\"minA2\" maxX=\"maxA2\"><a:pos x=\"xe\" y=\"vc\"/></a:ahXY>"
                "<a:ahXY gdRefX=\"adj3\" minX=\"0\" maxX=\"maxA3\"><a:pos x=\"xs3\" y=\"yt\"/></a:ahXY>"
                "<a:ahXY gdRefX=\"adj4\" minX=\"0\" maxX=\"maxA4\" gdRefY=\"adj5\" minY=\"2000\" maxY=\"maxA5\"><a:pos x=\"x4\" y=\"yt\"/></a:ahXY>",
                mv("l", "y1") + arc("r1", "r1", "cd2", "cd4") + ln("r", "t") +
                arc("fl", "hr", "0", "cd4") + ln("xs3", "yt") +
                arc("r3", "r3", "3cd4", "-5400000") + ln("xe", "yb3") +
                arc("r3", "r3", "cd2", "-5400000") + ln("x4", "yb") +
                arc("fl", "hr", "3cd4", "cd4") + ln("xr1", "b") +
                arc("r1", "r1", "cd4", "cd4") + "<a:close/>",
                // 바깥 사각형 4점(오른쪽은 홈이 열린 보이지 않는 변) + 홈 안쪽 끝 중앙 1점
                cn("3cd4", "xtm", "t") + cn("cd2", "l", "vc") + cn("cd4", "xtm", "b") + cn("0", "r", "vc") +
                cn("0", "xe", "vc"));
        }

        // 벤 다이어그램의 원 하나. pos: 0=왼쪽 위(A) 1=오른쪽 위(B) 2=가운데 아래(C).
        // 세로 지름 dv = h(2원) 또는 h/1.52(3원, 중심이 정삼각형이 되는 기본 겹침),
        // 가로 지름 dw = dv·adj (adj 100000 = 정원).
        S geomVennCircle(int pos, bool three)
        {
            const char* x0 = pos == 0 ? "l" : pos == 1 ? "xB" : "xC";
            const char* y0 = pos == 2 ? "yC" : "t";
            S gds;
            gds = gd("dv", three ? "*/ h 100 152" : "*/ h 1 1") +
                  gd("maxAdj", "*/ w 100000 dv") + gd("a", "pin 0 adj maxAdj") +
                  gd("dw", "*/ dv a 100000") + gd("dwd2", "*/ dw 1 2") + gd("dvd2", "*/ dv 1 2") +
                  gd("xB", "+- r 0 dw") + gd("xC", "+- hc 0 dwd2") + gd("yC", "+- b 0 dv") +
                  gd("x0", S("+- ") + x0 + " 0 0") + gd("xR", "+- x0 dw 0") +
                  gd("yc", S("+- ") + y0 + " dvd2 0") + gd("xm", "+- x0 dwd2 0") +
                  gd("y0g", S("+- ") + y0 + " 0 0") + gd("ybv", "+- y0g dv 0");
            return geom(
                gd("adj", "val 100000"), gds,
                S("<a:ahXY gdRefX=\"adj\" minX=\"0\" maxX=\"maxAdj\"><a:pos x=\"") +
                    (pos == 1 ? "x0" : "xR") + "\" y=\"yc\"/></a:ahXY>",
                mv("x0", "yc") + fullEllipse("dwd2", "dvd2") + "<a:close/>",
                cn("3cd4", "xm", "y0g") + cn("cd2", "x0", "yc") + cn("cd4", "xm", "ybv") + cn("0", "xR", "yc"));
        }

        S solid(const char* rgb, int alpha = 0)
        {
            S c = S("<a:solidFill><a:srgbClr val=\"") + rgb + "\">";
            if (alpha) c += "<a:alpha val=\"" + std::to_string(alpha) + "\"/>";
            return c + "</a:srgbClr></a:solidFill>";
        }
        S outline(const char* rgb)
        { return S("<a:ln w=\"12700\">") + solid(rgb) + "</a:ln>"; }

        S xfrm(long long cx, long long cy)
        {
            return "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + std::to_string(cx) +
                   "\" cy=\"" + std::to_string(cy) + "\"/></a:xfrm>";
        }

        S spShape(int id, const char* name, long long cx, long long cy, const S& geomXml,
                  const S& fill, const S& line)
        {
            return "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"" +
                   std::to_string(id) + "\" name=\"" + name + "\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
                   "<xdr:spPr>" + xfrm(cx, cy) + geomXml + fill + line + "</xdr:spPr></xdr:sp>";
        }

        S spShapeTx(int id, const char* name, long long cx, long long cy, const S& geomXml,
                    const S& fill, const S& line, const S& tx)
        {
            S s = spShape(id, name, cx, cy, geomXml, fill, line);
            s.insert(s.rfind("</xdr:sp>"), tx);
            return s;
        }

        S wrapDrawing(long long cx, long long cy, const S& body)
        {
            return S("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                     "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""
                     " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                     "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"") +
                   std::to_string(cx) + "\" cy=\"" + std::to_string(cy) + "\"/>" + body +
                   "<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>";
        }

        // 제목/본문 2단 사각형(그룹): 도형 2개.
        //  HS_TblFrame: 테두리 있는 도형. 바깥 사각형 + 접합선을 그리고, 자기 텍스트 틀(위~접합선)을
        //   제목 글로 쓴다. 핸들1(오른쪽) = 접합선 y(adj, 절대 EMU) — 크기를 바꿔도 제목 높이 유지.
        //  연결점 8개: 테두리 도형(Frame) = 사변 중앙 4 + 제목 칸 좌·우 중앙 2, 본문 도형(Body) = 본문 칸 좌·우 중앙 2.
        //  HS_TblBody: 본문 글 영역(핸들2 윗변~아래)만 덮는, 테두리 없고 투명 채우기(클릭 가능)인 본문 텍스트 도형. 핸들2(가운데) =
        //   본문 글 틀의 윗변(adj, 절대 EMU). 자기 크기 안에서만 움직이며 그룹 크기를 따라 커진다.
        S geomTableFrame()
        {
            return "<a:custGeom><a:avLst>" + gd("adj", "val 432000") + "</a:avLst><a:gdLst>" +
                   gd("maxA", "+- h 0 108000") + gd("yj", "pin 108000 adj maxA") + gd("ymh", "+/ t yj 2") +
                   "</a:gdLst><a:ahLst>"
                   "<a:ahXY gdRefY=\"adj\" minY=\"108000\" maxY=\"maxA\"><a:pos x=\"r\" y=\"yj\"/></a:ahXY>"
                   "</a:ahLst><a:cxnLst>" + cn("3cd4", "hc", "t") + cn("cd4", "hc", "b") +
                   cn("cd2", "l", "vc") + cn("0", "r", "vc") +
                   cn("cd2", "l", "ymh") + cn("0", "r", "ymh") +
                   "</a:cxnLst><a:rect l=\"l\" t=\"t\" r=\"r\" b=\"yj\"/><a:pathLst>"
                   "<a:path>" + mv("l", "t") + ln("r", "t") + ln("r", "b") + ln("l", "b") + "<a:close/></a:path>"
                   "<a:path fill=\"none\">" + mv("l", "yj") + ln("r", "yj") + "</a:path>"
                   "</a:pathLst></a:custGeom>";
        }

        S geomTableBody()
        {
            return "<a:custGeom><a:avLst>" + gd("adj", "val 468000") + "</a:avLst><a:gdLst>" +
                   gd("maxA", "+- h 0 108000") + gd("yt", "pin 108000 adj maxA") + gd("ymb", "+/ yt b 2") +
                   "</a:gdLst><a:ahLst>"
                   "<a:ahXY gdRefY=\"adj\" minY=\"108000\" maxY=\"maxA\"><a:pos x=\"hc\" y=\"yt\"/></a:ahXY>"
                   "</a:ahLst><a:cxnLst>" + cn("cd2", "l", "ymb") + cn("0", "r", "ymb") +
                   "</a:cxnLst><a:rect l=\"l\" t=\"yt\" r=\"r\" b=\"b\"/><a:pathLst>"
                   "<a:path>" + mv("l", "yt") + ln("r", "yt") + ln("r", "b") + ln("l", "b") + "<a:close/></a:path>"
                   "</a:pathLst></a:custGeom>";
        }

        S txBody(const char* anchor, const S& paras)
        {
            return S("<xdr:txBody><a:bodyPr wrap=\"square\" lIns=\"72000\" tIns=\"36000\" rIns=\"72000\" bIns=\"36000\" anchor=\"") +
                   anchor + "\"/><a:lstStyle/>" + paras + "</xdr:txBody>";
        }

        S tableGroup(long long cx, long long cy)
        {
            S g = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"10\" name=\"HS_Shape\"/>"
                  "<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>" + xfrm(cx, cy);
            g.insert(g.rfind("</a:xfrm>"),
                     "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"" + std::to_string(cx) +
                     "\" cy=\"" + std::to_string(cy) + "\"/>");
            g += "</xdr:grpSpPr>";
            const S line = "<a:ln w=\"25400\">" + solid("4472C4") + "</a:ln>";
            const S black = "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>";
            // 샘플 글 없음: 빈 문단만 두고, 입력 시 쓸 글꼴 크기(제목 18pt 가운데 / 본문 11pt 일반 문단)만 지정.
            const S titlePara = "<a:p><a:pPr algn=\"ctr\"/><a:endParaRPr lang=\"en-US\" sz=\"1800\">" + black +
                                "</a:endParaRPr></a:p>";
            const S bodyPara = "<a:p><a:endParaRPr lang=\"en-US\" sz=\"1100\">" + black +
                               "</a:endParaRPr></a:p>";
            g += spShapeTx(11, "HS_TblFrame", cx, cy, geomTableFrame(), solid("FFFFFF"), line,
                           txBody("ctr", titlePara));
            g += spShapeTx(12, "HS_TblBody", cx, cy, geomTableBody(), 
                           // 투명도 100% 채우기: noFill이면 글이 없는 곳은 클릭이 아래 도형으로 통과한다.
                           "<a:solidFill><a:srgbClr val=\"FFFFFF\"><a:alpha val=\"0\"/></a:srgbClr></a:solidFill>",
                           "<a:ln><a:noFill/></a:ln>", txBody("t", bodyPara));
            return g + "</xdr:grpSp>";
        }

        S vennGroup(long long cx, long long cy, bool three)
        {
            S g = "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"10\" name=\"HS_Shape\"/>"
                  "<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>" + xfrm(cx, cy);
            // xfrm 에 chOff/chExt 를 덧붙인다(자식 좌표계 = 그룹 좌표계).
            g.insert(g.rfind("</a:xfrm>"),
                     "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"" + std::to_string(cx) +
                     "\" cy=\"" + std::to_string(cy) + "\"/>");
            g += "</xdr:grpSpPr>";
            g += spShape(11, "HS_VennA", cx, cy, geomVennCircle(0, three),
                         solid("4472C4", 55000), outline("2F528F"));
            g += spShape(12, "HS_VennB", cx, cy, geomVennCircle(1, three),
                         solid("ED7D31", 55000), outline("C55A11"));
            if (three)
                g += spShape(13, "HS_VennC", cx, cy, geomVennCircle(2, three),
                             solid("70AD47", 55000), outline("548235"));
            return g + "</xdr:grpSp>";
        }

        constexpr const char* kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

        S buildXlsx(const S& drawingXml)
        {
            const S d = kXmlDecl;
            std::vector<std::pair<S, S>> parts = {
                { "[Content_Types].xml", d +
                  "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                  "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                  "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                  "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                  "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                  "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"
                  "</Types>" },
                { "_rels/.rels", d +
                  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                  "</Relationships>" },
                { "xl/workbook.xml", d +
                  "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                  " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                  "<sheets><sheet name=\"Shapes\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>" },
                { "xl/_rels/workbook.xml.rels", d +
                  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                  "</Relationships>" },
                { "xl/worksheets/sheet1.xml", d +
                  "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                  " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                  "<sheetData/><drawing r:id=\"rId1\"/></worksheet>" },
                { "xl/worksheets/_rels/sheet1.xml.rels", d +
                  "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                  "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/>"
                  "</Relationships>" },
                { "xl/drawings/drawing1.xml", drawingXml },
            };
            return buildZip(parts);
        }

        // ── 도형 정의(기본 크기 mm) ────────────────────────────────────────
        enum class Kind { OneSide, Donut, TabPad, Slot, Table, Venn2, Venn3 };

        struct ShapeDef
        {
            Kind           kind;
            double         wMm, hMm;
            const wchar_t* nameKey;      // i18n 키: 도형 이름(삽입 도형 Name 접두)
            long           placement;    // 1 MoveAndSize, 2 Move, 3 FreeFloating
        };
        constexpr double kVennD = 24.0;   // 벤 원 기본 지름(mm), 겹침 = 0.6D
        constexpr ShapeDef kOneSide{ Kind::OneSide, 30.0, 15.0, L"ribbon.egpp.shpOneSide.label", 1 };
        constexpr ShapeDef kDonut  { Kind::Donut,   20.0, 20.0, L"ribbon.egpp.shpDonut.label",   1 };
        constexpr ShapeDef kTabPad { Kind::TabPad,  24.0, 30.0, L"ribbon.egpp.shpTabPad.label",  1 };
        constexpr ShapeDef kSlot   { Kind::Slot,    36.0, 24.0, L"ribbon.egpp.shpSlot.label",    1 };
        constexpr ShapeDef kTable  { Kind::Table,   60.0, 44.0, L"ribbon.egpp.shpTable.label",   1 };
        constexpr ShapeDef kVenn2  { Kind::Venn2,   kVennD * 1.6, kVennD,        L"ribbon.egpp.shpVennTwo.label", 2 };
        constexpr ShapeDef kVenn3  { Kind::Venn3,   kVennD * 1.6, kVennD * 1.52, L"ribbon.egpp.shpVennThree.label", 2 };

        S drawingFor(const ShapeDef& d)
        {
            const long long cx = (long long)(d.wMm * 36000.0 + 0.5);
            const long long cy = (long long)(d.hMm * 36000.0 + 0.5);
            switch (d.kind)
            {
            case Kind::OneSide:
                return wrapDrawing(cx, cy, spShape(2, "HS_Shape", cx, cy, geomOneSide(),
                                   solid("2E75B6"), outline("1F4E79")));
            case Kind::Donut:
                return wrapDrawing(cx, cy, spShape(2, "HS_Shape", cx, cy, geomDonut(),
                                   solid("2E75B6"), outline("1F4E79")));
            case Kind::TabPad:
                return wrapDrawing(cx, cy, spShape(2, "HS_Shape", cx, cy, geomTabPad(),
                                   solid("2E75B6"), outline("1F4E79")));
            case Kind::Slot:
                return wrapDrawing(cx, cy, spShape(2, "HS_Shape", cx, cy, geomSlot(),
                                   solid("2E75B6"), outline("1F4E79")));
            case Kind::Table: return wrapDrawing(cx, cy, tableGroup(cx, cy));
            case Kind::Venn2: return wrapDrawing(cx, cy, vennGroup(cx, cy, false));
            case Kind::Venn3: return wrapDrawing(cx, cy, vennGroup(cx, cy, true));
            }
            return S();
        }

        std::filesystem::path tempXlsxPath()
        {
            wchar_t dir[MAX_PATH + 2] = {};
            const DWORD n = GetTempPathW(MAX_PATH, dir);
            std::wstring name = L"EGTools_shape_" + std::to_wstring(GetCurrentProcessId()) +
                                L"_" + std::to_wstring(GetTickCount64()) + L".xlsx";
            return std::filesystem::path(std::wstring(dir, n)) / name;
        }

        VARIANT varBool(bool b)
        {
            VARIANT v; VariantInit(&v);
            v.vt = VT_BOOL; v.boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;
            return v;
        }

        // 임시 통합문서를 열어 도형을 복사 → 활성 시트의 활성 셀 좌상단에 붙여넣는다.
        void insertShape(const ShapeDef& def)
        {
            using egtools::i18n::t;
            std::filesystem::path tmp;
            try
            {
                auto& xapp = xloil::thisApp();
                xloil::PauseExcel pause(xapp);

                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };

                IDispatch* ws = getObject(ad, L"ActiveSheet");
                if (!ws) { msgWarn(t(L"cmd.shape.noSheet")); return; }
                Releaser rWs{ ws };
                IDispatch* cell = getObject(ad, L"ActiveCell");
                if (!cell) { msgWarn(t(L"cmd.shape.noSheet")); return; }
                Releaser rCell{ cell };
                IDispatch* targetWin = getObject(ad, L"ActiveWindow");
                Releaser rWin{ targetWin };

                tmp = tempXlsxPath();
                {
                    const S zip = buildXlsx(drawingFor(def));
                    std::ofstream f(tmp, std::ios::binary);
                    f.write(zip.data(), (std::streamsize)zip.size());
                    if (!f) { msgWarn(t(L"cmd.shape.fail")); return; }
                }

                IDispatch* wbs = getObject(ad, L"Workbooks");
                if (!wbs) { msgWarn(t(L"cmd.shape.fail")); return; }
                Releaser rWbs{ wbs };

                IDispatch* src = nullptr;
                bool pasted = false;
                IDispatch* sh = nullptr;
                auto cleanup = [&]
                {
                    putBool(ad, L"CutCopyMode", false);
                    if (src)
                    {
                        VARIANT sv = varBool(false);
                        callMethod(src, L"Close", &sv, 1);
                        src->Release(); src = nullptr;
                    }
                    std::error_code ec;
                    std::filesystem::remove(tmp, ec);
                };
                struct Guard { std::function<void()> f; ~Guard() { if (f) f(); } } guard{ cleanup };

                {
                    VARIANT a[3] = { varBStr(tmp.wstring()), varLong(0), varBool(true) };
                    src = getObject(wbs, L"Open", a, 3);
                    VariantClear(&a[0]);
                }
                if (!src) { msgWarn(t(L"cmd.shape.fail")); return; }
                if (IDispatch* wins = getObject(src, L"Windows"))
                {
                    Releaser r1{ wins };
                    if (IDispatch* w1 = getObjectIdx(wins, L"Item", 1))
                    {
                        Releaser r2{ w1 };
                        putBool(w1, L"Visible", false);
                    }
                }
                if (targetWin) callMethod(targetWin, L"Activate");   // Open이 활성 창을 바꿈

                // 원본 도형 복사 → 붙여넣기. 클립보드를 다른 프로그램(클립보드 관리자
                // 등)이 잠시 쥐고 있으면 Copy/Paste가 간헐적으로 실패하므로 재시도한다.
                IDispatch* sws = nullptr;
                if (IDispatch* sheets = getObject(src, L"Worksheets"))
                {
                    Releaser r{ sheets };
                    sws = getObjectIdx(sheets, L"Item", 1);
                }
                Releaser rSws{ sws };
                IDispatch* sshapes = sws ? getObject(sws, L"Shapes") : nullptr;
                Releaser rSsh{ sshapes };
                IDispatch* sshp = sshapes ? getObjectIdx(sshapes, L"Item", 1) : nullptr;
                Releaser rSshp{ sshp };
                if (!sshp) { msgWarn(t(L"cmd.shape.fail")); return; }

                IDispatch* shapes = getObject(ws, L"Shapes");
                if (!shapes) { msgWarn(t(L"cmd.shape.fail")); return; }
                Releaser rShapes{ shapes };
                const long before = getLong(shapes, L"Count", 0);

                long n = before;
                for (int attempt = 0; attempt < 5 && n <= before; ++attempt)
                {
                    if (attempt) Sleep(200);
                    if (targetWin) callMethod(targetWin, L"Activate");
                    if (!callMethod(sshp, L"Copy")) continue;
                    pasted = callMethod(ws, L"Paste");
                    n = getLong(shapes, L"Count", 0);
                }
                if (n <= before) { msgWarn(t(L"cmd.shape.protected")); return; }

                sh = getObjectIdx(shapes, L"Item", n);
                if (!sh) { msgWarn(t(L"cmd.shape.fail")); return; }
                Releaser rSh{ sh };

                // 활성 셀 좌상단(도형 좌표계 보정: ShapeGrid)
                double l = 0, tp = 0, w = 0, h = 0;
                latecom::ShapeGrid grid;
                if (grid.rect(sh, cell, l, tp, w, h))
                {
                    putDouble(sh, L"Left", l);
                    putDouble(sh, L"Top", tp);
                }
                putLong(sh, L"Placement", def.placement);
                putBStr(sh, L"Name",
                        t(def.nameKey) + L" " + std::to_wstring(getLong(sh, L"ID", n)));

                guard.f = nullptr;     // 아래에서 명시 정리 후 선택
                cleanup();
                callMethod(sh, L"Select");
            }
            catch (...)
            {
                std::error_code ec;
                if (!tmp.empty()) std::filesystem::remove(tmp, ec);
                msgWarn(t(L"cmd.shape.fail"));
            }
        }
    }

    void insertShapeOneSide() { insertShape(kOneSide); }
    void insertShapeDonut()   { insertShape(kDonut); }
    void insertShapeTabPad()  { insertShape(kTabPad); }
    void insertShapeSlot()    { insertShape(kSlot); }
    void insertShapeTable()   { insertShape(kTable); }
    void insertShapeVenn2()   { insertShape(kVenn2); }
    void insertShapeVenn3()   { insertShape(kVenn3); }
}
