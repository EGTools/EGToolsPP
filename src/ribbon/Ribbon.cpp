// Ribbon.cpp — EGTools++ Ribbon (COM add-in) via xlOil.
//
// xlOil's makeComAddin()/connect() registers a COM add-in and serves the given
// customUI XML through IRibbonExtensibility — this works from a static XLL and
// runs its callbacks on Excel's main thread.
//
// 구조(plan/23 M1):
//  - 그룹/버튼은 선언 테이블(kGroups)로 정의하고 XML은 이 테이블에서 생성한다.
//    마일스톤 진행에 따라 그룹을 추가하면 XML·콜백·아이콘이 함께 확장된다.
//  - 라벨/툴팁은 XML에 넣지 않고 GetLabel/GetSupertip 콜백에서 i18n 카탈로그
//    키("ribbon.<컨트롤ID>.label|.tip")로 조회한다(6개국어, resources/i18n).
//  - 아이콘은 RCDATA PNG(32×32) → WIC 디코드 → HICON → IPictureDisp를
//    GetImage 콜백으로 반환한다. loadImage 콜백은 xlOil 디스패치가 첫 인수를
//    IRibbonControl*로 캐스팅하므로 사용 불가(BSTR 인수 → 크래시). getImage는
//    표준 컨트롤 시그니처라 안전하다. PICTYPE_ICON은 알파 보존이 확실한 경로.
//  - tag="LT2021"/"GT2016" + GetVisible로 Excel 버전별 표시 게이트(VB 규약,
//    365는 최신 취급).

#include "Ribbon.h"
#include "Convert.h"
#include "commands/Commands.h"
#include "../core/I18n.h"
#include "../core/Version.h"
#include <resource.h>

#include <xlOil/ExcelUI.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/State.h>

#include <windows.h>
#include <commctrl.h>
#include <shellapi.h>
#include <winhttp.h>
#include <wincodec.h>
#include <olectl.h>

#include <atomic>
#include <cwctype>
#include <map>
#include <memory>
#include <string>
#include <vector>

#pragma comment(lib, "winhttp.lib")

// EG_VERSION_STR("major.minor.build")는 빌드마다 생성되는 BuildNum.h가 제공.
#include <BuildNum.h>

// Widen the narrow EG_VERSION_STR macro to a wide literal.
#define EG_WIDEN2(x) L##x
#define EG_WIDEN(x)  EG_WIDEN2(x)

using namespace xloil;

// Static XLL does not auto-connect COM (see FxLet.cpp).
namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::ribbon
{
    namespace
    {
        std::shared_ptr<IComAddin> g_addin;
        std::atomic<bool> g_shutdown{ false };
        int g_attempts = 0;                     // main-thread only
        constexpr int kMaxAttempts = 600;       // ~10 min at 1s between retries

        // ── declarative ribbon definition ───────────────────────────────────
        // 컨트롤 ID는 i18n 키("ribbon.<id>.label"/".tip")와 아이콘·콜백 매핑의
        // 단일 식별자다. 그룹 추가 시 이 테이블과 kIcons, callbackFor()의 명령
        // 항목만 확장하면 된다.
        struct BtnDef
        {
            const wchar_t* id;        // control id
            const wchar_t* onAction;  // command callback name
            const wchar_t* image;     // custom icon name (kIcons) — nullptr이면 imageMso
            const wchar_t* imageMso;  // built-in Office icon (image == nullptr일 때)
            const wchar_t* tag;       // visibility gate "LT<year>"/"GT<year>" or nullptr
        };
        // NB: 목록은 반드시 "이름 있는 constexpr 배열 + 포인터/개수"로 담는다.
        // std::initializer_list 멤버는 MSVC에서 백킹 배열이 정적 수명을 얻지 못해
        // 런타임 순회 시 AV(0xC0000005)가 났다(2026-08-03 실측 — catch(...)로도
        // 잡히지 않아 Excel 즉사).
        struct SplitDef
        {
            const wchar_t* id;        // splitButton id
            BtnDef         defBtn;    // 큰 기본 버튼(메뉴 항목 중 하나와 같은 동작,
                                      // id는 "<항목id>Btn" — i18n 키는 접미사 제거로 공유)
            const wchar_t* menuId;    // 드롭다운 메뉴 id (라벨은 ribbon.<menuId>.label)
            const BtnDef*  items;
            size_t         count;
        };
        struct ControlDef        // btn/split 중 정확히 하나만 non-null
        {
            const BtnDef*   btn;
            const SplitDef* split;
        };
        struct GroupDef
        {
            const wchar_t*    id;
            const ControlDef* controls;
            size_t            count;
        };
        struct CtxMenuDef        // Office 기본 컨텍스트 메뉴 확장
        {
            const wchar_t* idMso;
            const BtnDef*  items;
            size_t         count;
        };

        // ── M3~M4: 보이는 셀 ──
        constexpr BtnDef kVisCopyMenuItems[] = {
            { L"egpp.visCopy", L"onVisCopy", L"visible-copy", nullptr, nullptr },
            { L"egpp.copyAll", L"onCopyAll", nullptr, L"Copy", nullptr },
        };
        constexpr SplitDef kSbVisCopy = {
            L"egpp.sbVisCopy",
            { L"egpp.visCopyBtn", L"onVisCopy", L"visible-copy", nullptr, nullptr },
            L"egpp.mnViCopy", kVisCopyMenuItems, _countof(kVisCopyMenuItems),
        };
        constexpr BtnDef kVisPasteMenuItems[] = {
            { L"egpp.pasteValues",   L"onPasteValues",   L"paste-values",   nullptr, nullptr },
            { L"egpp.pasteFormulas", L"onPasteFormulas", L"paste-formulas", nullptr, nullptr },
            { L"egpp.pasteAll",      L"onPasteAll",      L"paste-all",      nullptr, nullptr },
        };
        constexpr SplitDef kSbVisPaste = {
            L"egpp.sbVisPaste",
            { L"egpp.pasteMain", L"onPasteValues", L"paste-values", nullptr, nullptr },
            L"egpp.mnViPaste", kVisPasteMenuItems, _countof(kVisPasteMenuItems),
        };
        constexpr ControlDef kVisibleControls[] = {
            { nullptr, &kSbVisCopy }, { nullptr, &kSbVisPaste },
        };

        // ── M2~M4: 병합/나누기 (VB 메뉴 순서) ──
        constexpr BtnDef kMergeMenuItems[] = {
            { L"egpp.mergeClone",     L"onMergeClone",     L"merge-clone",      nullptr, nullptr },
            { L"egpp.mergeFake",      L"onMergeFake",      L"merge-fake",       nullptr, nullptr },
            { L"egpp.mergeFakeClone", L"onMergeFakeClone", L"merge-fake-clone", nullptr, nullptr },
            { L"egpp.mergeJoin",      L"onMergeJoin",      L"merge-join",       nullptr, nullptr },
            { L"egpp.mergeJoinAll",   L"onMergeJoinAll",   L"merge-join-all",   nullptr, nullptr },
            { L"egpp.mergeRow",       L"onMergeRow",       L"merge-row",        nullptr, nullptr },
            { L"egpp.mergeRowAll",    L"onMergeRowAll",    L"merge-row-all",    nullptr, nullptr },
            { L"egpp.mergeCol",       L"onMergeCol",       L"merge-col",        nullptr, nullptr },
            { L"egpp.mergeColAll",    L"onMergeColAll",    L"merge-col-all",    nullptr, nullptr },
        };
        constexpr SplitDef kSbMerge = {
            L"egpp.sbMerge",
            { L"egpp.mergeMain", L"onMergeJoin", L"merge-join", nullptr, nullptr },
            L"egpp.mnMerge", kMergeMenuItems, _countof(kMergeMenuItems),
        };
        constexpr BtnDef kSplitMenuItems[] = {
            { L"egpp.splitFill",   L"onSplitFill",   L"split-fill",    nullptr, nullptr },
            { L"egpp.splitRow",    L"onSplitRow",    L"split-row",     nullptr, nullptr },
            { L"egpp.splitRowAll", L"onSplitRowAll", L"split-row-all", nullptr, nullptr },
            { L"egpp.splitCol",    L"onSplitCol",    L"split-col",     nullptr, nullptr },
            { L"egpp.splitColAll", L"onSplitColAll", L"split-col-all", nullptr, nullptr },
        };
        constexpr SplitDef kSbSplit = {
            L"egpp.sbSplit",
            { L"egpp.splitMain", L"onSplitRow", L"split-row", nullptr, nullptr },
            L"egpp.mnSplit", kSplitMenuItems, _countof(kSplitMenuItems),
        };
        constexpr ControlDef kMergeControls[] = {
            { nullptr, &kSbMerge }, { nullptr, &kSbSplit },
        };

        // ── M2~M3: 그림/사진 (M4: 추출 저장 추가) ──
        constexpr BtnDef kPicMenuItems[] = {
            { L"egpp.picInsert", L"onPicInsert", L"picture-insert", nullptr, nullptr },
            { L"egpp.picFolder", L"onPicFolder", L"picture-folder", nullptr, nullptr },
            { L"egpp.picForm",   L"onPicForm",   L"picture-form",   nullptr, nullptr },
        };
        constexpr SplitDef kSbPicture = {
            L"egpp.sbPicture",
            { L"egpp.picInsertBtn", L"onPicInsert", L"picture-insert", nullptr, nullptr },
            L"egpp.mnPicture", kPicMenuItems, _countof(kPicMenuItems),
        };
        constexpr BtnDef kFitMenuItems[] = {
            { L"egpp.fitCell",    L"onFitCell",    L"fit-cell",        nullptr, nullptr },
            { L"egpp.fitAll",     L"onFitAll",     L"fit-all",         nullptr, nullptr },
            { L"egpp.exportPics", L"onExportPics", L"export-pictures", nullptr, nullptr },
        };
        constexpr SplitDef kSbFit = {
            L"egpp.sbFit",
            { L"egpp.fitCellBtn", L"onFitCell", L"fit-cell", nullptr, nullptr },
            L"egpp.mnFit", kFitMenuItems, _countof(kFitMenuItems),
        };
        constexpr ControlDef kPictureControls[] = {
            { nullptr, &kSbPicture }, { nullptr, &kSbFit },
        };

        // ── M3: 달력/일정표 ──
        constexpr BtnDef kCalMenuItems[] = {
            { L"egpp.calMonth", L"onCalMonth", L"calendar-month",  nullptr, nullptr },
            { L"egpp.calYearA", L"onCalYearA", L"calendar-year-a", nullptr, nullptr },
            { L"egpp.calYearB", L"onCalYearB", L"calendar-year-b", nullptr, nullptr },
        };
        constexpr SplitDef kSbCal = {
            L"egpp.sbCal",
            { L"egpp.calMonthBtn", L"onCalMonth", L"calendar-month", nullptr, nullptr },
            L"egpp.mnCal", kCalMenuItems, _countof(kCalMenuItems),
        };
        constexpr BtnDef kSchMenuItems[] = {
            { L"egpp.schWeek", L"onSchWeek", L"schedule-week",  nullptr, nullptr },
            { L"egpp.schDayA", L"onSchDayA", L"schedule-day-a", nullptr, nullptr },
            { L"egpp.schDayB", L"onSchDayB", L"schedule-day-b", nullptr, nullptr },
        };
        constexpr SplitDef kSbSch = {
            L"egpp.sbSch",
            { L"egpp.schDayBBtn", L"onSchDayB", L"schedule-day-b", nullptr, nullptr },
            L"egpp.mnSch", kSchMenuItems, _countof(kSchMenuItems),
        };
        constexpr ControlDef kCalendarControls[] = {
            { nullptr, &kSbCal }, { nullptr, &kSbSch },
        };

        // ── M3~M5: 양식 ──
        constexpr BtnDef kFormMenuItems[] = {
            { L"egpp.labels",    L"onLabels",    L"labels",     nullptr, nullptr },
            { L"egpp.mailMerge", L"onMailMerge", L"mail-merge", nullptr, nullptr },
        };
        constexpr SplitDef kSbForm = {
            L"egpp.sbForm",
            { L"egpp.labelsBtn", L"onLabels", L"labels", nullptr, nullptr },
            L"egpp.mnForm", kFormMenuItems, _countof(kFormMenuItems),
        };
        constexpr ControlDef kFormControls[] = { { nullptr, &kSbForm } };

        // ── M2~M3: 도구 (M4: 모양 뽑기 추가) ──
        constexpr BtnDef kToolsMenuItems[] = {
            { L"egpp.deleteStyles", L"onDeleteStyles", L"delete-styles", nullptr, nullptr },
            { L"egpp.deleteNames",  L"onDeleteNames",  L"delete-names",  nullptr, nullptr },
            { L"egpp.clearError",   L"onClearError",   L"clear-error",   nullptr, nullptr },
            { L"egpp.clearEmpty",   L"onClearEmpty",   L"clear-empty",   nullptr, nullptr },
            { L"egpp.memoArrange",  L"onMemoArrange",  L"memo-arrange",  nullptr, nullptr },
            { L"egpp.borderColor",  L"onBorderColor",  L"border-color",  nullptr, nullptr },
            { L"egpp.drawShape",    L"onDrawShape",    L"draw-shape",    nullptr, nullptr },
            { L"egpp.udfFreeze",    L"onUdfFreeze",    L"udf-freeze",    nullptr, nullptr },
            { L"egpp.checkBox",     L"onCheckBox",     L"checkbox",      nullptr, nullptr },
            { L"egpp.recalc",       L"onRecalc",       L"recalc",        nullptr, nullptr },
        };
        constexpr SplitDef kSbTools = {
            L"egpp.sbTools",
            { L"egpp.recalcBtn", L"onRecalc", L"recalc", nullptr, nullptr },
            L"egpp.mnTools", kToolsMenuItems, _countof(kToolsMenuItems),
        };
        constexpr ControlDef kToolsControls[] = { { nullptr, &kSbTools } };

        // ── EGTools: 스플릿 버튼 단일 그룹 — 대표 기능은 EGTools 정보, API 키
        //    관리는 도구 메뉴에서 이 그룹으로 이동(사용자 지시, 2026-08-04) ──
        constexpr BtnDef kMainMenuItems[] = {
            { L"egpp.toCompat",  L"onToCompat",  L"compat-apply",   nullptr, nullptr },
            { L"egpp.toNative",  L"onToNative",  L"restore-native", nullptr, nullptr },
            { L"egpp.smtpSetup", L"onSmtpSetup", L"smtp-setup",     nullptr, nullptr },
            { L"egpp.apiKeys",   L"onApiKeys",   L"api-keys",       nullptr, nullptr },
            { L"egpp.about",     L"onAbout",     L"about",          nullptr, nullptr },
        };
        constexpr SplitDef kSbMain = {
            L"egpp.sbMain",
            { L"egpp.aboutBtn", L"onAbout", L"about", nullptr, nullptr },
            L"egpp.mnMain", kMainMenuItems, _countof(kMainMenuItems),
        };
        constexpr ControlDef kMainControls[] = { { nullptr, &kSbMain } };

        constexpr GroupDef kGroups[] = {
            { L"egpp.grpVisible",  kVisibleControls,  _countof(kVisibleControls) },
            { L"egpp.grpMerge",    kMergeControls,    _countof(kMergeControls) },
            { L"egpp.grpPicture",  kPictureControls,  _countof(kPictureControls) },
            { L"egpp.grpCalendar", kCalendarControls, _countof(kCalendarControls) },
            { L"egpp.grpForm",     kFormControls,     _countof(kFormControls) },
            { L"egpp.grpTools",    kToolsControls,    _countof(kToolsControls) },
            { L"egpp.grpMain",     kMainControls,     _countof(kMainControls) },
        };

        // ── 컨텍스트 메뉴: VB와 동일 4곳(날짜입력 + 링크열기) ──
        // 컨트롤 id는 메뉴마다 달라야 하므로 숫자 접미사; i18n 키는 접미사 제거로 공유.
        constexpr BtnDef kCtxCell[] = {
            { L"egpp.ctxDatePicker1", L"onDatePicker", L"date-picker", nullptr, nullptr },
            { L"egpp.ctxOpenUrl1",    L"onOpenUrl",    nullptr, L"HyperlinkEdit", nullptr },
        };
        constexpr BtnDef kCtxList[] = {
            { L"egpp.ctxDatePicker2", L"onDatePicker", L"date-picker", nullptr, nullptr },
            { L"egpp.ctxOpenUrl2",    L"onOpenUrl",    nullptr, L"HyperlinkEdit", nullptr },
        };
        constexpr BtnDef kCtxCellLay[] = {
            { L"egpp.ctxDatePicker3", L"onDatePicker", L"date-picker", nullptr, nullptr },
            { L"egpp.ctxOpenUrl3",    L"onOpenUrl",    nullptr, L"HyperlinkEdit", nullptr },
        };
        constexpr BtnDef kCtxListLay[] = {
            { L"egpp.ctxDatePicker4", L"onDatePicker", L"date-picker", nullptr, nullptr },
            { L"egpp.ctxOpenUrl4",    L"onOpenUrl",    nullptr, L"HyperlinkEdit", nullptr },
        };
        constexpr CtxMenuDef kCtxMenus[] = {
            { L"ContextMenuCell",            kCtxCell,    _countof(kCtxCell) },
            { L"ContextMenuListRange",       kCtxList,    _countof(kCtxList) },
            { L"ContextMenuCellLayout",      kCtxCellLay, _countof(kCtxCellLay) },
            { L"ContextMenuListRangeLayout", kCtxListLay, _countof(kCtxListLay) },
        };

        // 삽입 탭 CheckBox 버튼 (VB tabInsert grCheckBox — 365 미만에서만 표시)
        constexpr BtnDef kInsertTabButtons[] = {
            { L"egpp.insCheckBox", L"onCheckBox", L"checkbox", nullptr, L"LT365" },
        };

        // 커스텀 아이콘 이름 → RCDATA PNG 리소스 ID.
        constexpr std::pair<const wchar_t*, int> kIcons[] = {
            { L"compat-apply",   IDR_ICON_COMPAT_APPLY },
            { L"restore-native", IDR_ICON_RESTORE_NATIVE },
            { L"about",          IDR_ICON_ABOUT },
            { L"split-fill",     IDR_ICON_SPLIT_FILL },
            { L"picture-insert", IDR_ICON_PICTURE_INSERT },
            { L"fit-cell",       IDR_ICON_FIT_CELL },
            { L"fit-all",        IDR_ICON_FIT_ALL },
            { L"recalc",         IDR_ICON_RECALC },
            { L"delete-names",   IDR_ICON_DELETE_NAMES },
            { L"memo-arrange",   IDR_ICON_MEMO_ARRANGE },
            { L"border-color",   IDR_ICON_BORDER_COLOR },
            { L"checkbox",       IDR_ICON_CHECKBOX },
            { L"api-keys",       IDR_ICON_API_KEYS },
            { L"merge-join",       IDR_ICON_MERGE_JOIN },
            { L"merge-join-all",   IDR_ICON_MERGE_JOIN_ALL },
            { L"merge-clone",      IDR_ICON_MERGE_CLONE },
            { L"merge-fake",       IDR_ICON_MERGE_FAKE },
            { L"merge-fake-clone", IDR_ICON_MERGE_FAKE_CLONE },
            { L"merge-row",        IDR_ICON_MERGE_ROW },
            { L"merge-row-all",    IDR_ICON_MERGE_ROW_ALL },
            { L"merge-col",        IDR_ICON_MERGE_COL },
            { L"merge-col-all",    IDR_ICON_MERGE_COL_ALL },
            { L"split-row",        IDR_ICON_SPLIT_ROW },
            { L"split-row-all",    IDR_ICON_SPLIT_ROW_ALL },
            { L"split-col",        IDR_ICON_SPLIT_COL },
            { L"split-col-all",    IDR_ICON_SPLIT_COL_ALL },
            { L"visible-copy",     IDR_ICON_VISIBLE_COPY },
            { L"picture-folder",   IDR_ICON_PICTURE_FOLDER },
            { L"picture-form",     IDR_ICON_PICTURE_FORM },
            { L"calendar-month",   IDR_ICON_CAL_MONTH },
            { L"calendar-year-a",  IDR_ICON_CAL_YEAR_A },
            { L"calendar-year-b",  IDR_ICON_CAL_YEAR_B },
            { L"schedule-week",    IDR_ICON_SCH_WEEK },
            { L"schedule-day-a",   IDR_ICON_SCH_DAY_A },
            { L"schedule-day-b",   IDR_ICON_SCH_DAY_B },
            { L"labels",           IDR_ICON_LABELS },
            { L"delete-styles",    IDR_ICON_DELETE_STYLES },
            { L"clear-error",      IDR_ICON_CLEAR_ERROR },
            { L"clear-empty",      IDR_ICON_CLEAR_EMPTY },
            { L"udf-freeze",       IDR_ICON_UDF_FREEZE },
            { L"paste-values",     IDR_ICON_PASTE_VALUES },
            { L"paste-formulas",   IDR_ICON_PASTE_FORMULAS },
            { L"paste-all",        IDR_ICON_PASTE_ALL },
            { L"export-pictures",  IDR_ICON_EXPORT_PICTURES },
            { L"draw-shape",       IDR_ICON_DRAW_SHAPE },
            { L"date-picker",      IDR_ICON_DATE_PICKER },
            { L"mail-merge",       IDR_ICON_MAIL_MERGE },
            { L"smtp-setup",       IDR_ICON_SMTP_SETUP },
        };

        // 버튼 XML. topLevel=true면 size="large"(그룹 직속/스플릿 기본 버튼),
        // false면 메뉴 항목(크기는 menu의 itemSize가 결정).
        void appendButton(std::wstring& xml, const BtnDef& b, bool topLevel,
                          const wchar_t* extraAttrs = nullptr)
        {
            xml += L"<button id=\"";
            xml += b.id;
            xml += L"\"";
            if (topLevel) xml += L" size=\"large\"";
            xml += L" getLabel=\"GetLabel\" getScreentip=\"GetLabel\""
                   L" getSupertip=\"GetSupertip\"";
            if (b.tag)
            {
                xml += L" tag=\"";
                xml += b.tag;
                xml += L"\" getVisible=\"GetVisible\"";
            }
            if (b.image)
                xml += L" getImage=\"GetImage\"";
            else if (b.imageMso)
            {
                xml += L" imageMso=\"";
                xml += b.imageMso;
                xml += L"\"";
            }
            if (extraAttrs) xml += extraAttrs;
            xml += L" onAction=\"";
            xml += b.onAction;
            xml += L"\"/>";
        }

        std::wstring ribbonXml()
        {
            std::wstring xml =
                L"<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
                L"<ribbon><tabs>";

            // 삽입 탭 CheckBox 버튼 (VB tabInsert grCheckBox, tag=LT365 —
            // 네이티브 확인란이 있는 365에서는 숨김)
            xml += L"<tab idMso=\"TabInsert\">"
                   L"<group id=\"egpp.grpInsertCB\" getLabel=\"GetLabel\""
                   L" tag=\"LT365\" getVisible=\"GetVisible\""
                   L" insertBeforeMso=\"GroupInsertChartsExcel\">";
            for (const auto& b : kInsertTabButtons)
                appendButton(xml, b, true);
            xml += L"</group></tab>";

            xml += L"<tab id=\"egpp.tab\" insertBeforeMso=\"TabHome\" getLabel=\"GetLabel\">";
            for (const auto& g : kGroups)
            {
                xml += L"<group id=\"";
                xml += g.id;
                xml += L"\" getLabel=\"GetLabel\">";
                for (size_t c = 0; c < g.count; ++c)
                {
                    const ControlDef& ctl = g.controls[c];
                    if (ctl.btn)
                        appendButton(xml, *ctl.btn, true);
                    else if (ctl.split)
                    {
                        const SplitDef& s = *ctl.split;
                        xml += L"<splitButton id=\"";
                        xml += s.id;
                        xml += L"\" size=\"large\">";
                        appendButton(xml, s.defBtn, false);   // splitButton 내부는 size 미지정
                        xml += L"<menu id=\"";
                        xml += s.menuId;
                        xml += L"\" getLabel=\"GetLabel\" itemSize=\"large\">";
                        for (size_t i = 0; i < s.count; ++i)
                            appendButton(xml, s.items[i], false);
                        xml += L"</menu></splitButton>";
                    }
                }
                xml += L"</group>";
            }
            xml += L"</tab></tabs></ribbon>";

            // 컨텍스트 메뉴(2010+ 전용 스키마 — EGTools++ 지원 범위와 일치)
            xml += L"<contextMenus>";
            int sep = 0;
            for (const auto& m : kCtxMenus)
            {
                xml += L"<contextMenu idMso=\"";
                xml += m.idMso;
                xml += L"\">";
                for (size_t i = 0; i < m.count; ++i)
                    appendButton(xml, m.items[i], false,
                                 L" insertBeforeMso=\"Cut\"");
                xml += L"<menuSeparator id=\"egpp.ctxSep";
                xml += std::to_wstring(++sep);
                xml += L"\" insertBeforeMso=\"Cut\"/>";
                xml += L"</contextMenu>";
            }
            xml += L"</contextMenus></customUI>";
            return xml;
        }

        // ── i18n helpers ────────────────────────────────────────────────────
        // 컨트롤 id → i18n 키 정규화: 스플릿 기본 버튼의 "Btn" 접미사와 컨텍스트
        // 메뉴 항목의 숫자 접미사(egpp.ctxOpenUrl1..4)를 제거해 문자열을 공유한다.
        std::wstring normalizedId(const wchar_t* id)
        {
            std::wstring s = id ? id : L"";
            while (!s.empty() && iswdigit(s.back())) s.pop_back();
            if (s.size() > 3 && s.compare(s.size() - 3, 3, L"Btn") == 0)
                s.erase(s.size() - 3);
            return s;
        }

        std::wstring ctrlText(const RibbonControl& ctrl, const wchar_t* suffix)
        {
            std::wstring key = L"ribbon." + normalizedId(ctrl.Id) + suffix;
            return egtools::i18n::t(key);
        }

        void setResult(VARIANT* result, const std::wstring& s)
        {
            if (!result) return;
            result->vt = VT_BSTR;
            result->bstrVal = SysAllocString(s.c_str());
        }

        // ── visibility gate ("LT2021"/"GT2016"; 365는 최신 취급) ────────────
        int normYear(int y) { return (y == 365 || y == 3650) ? 9999 : y; }

        bool visibleForTag(const wchar_t* tag)
        {
            if (!tag || wcslen(tag) < 3) return true;
            const int host = normYear(egtools::core::excelYear());
            if (host <= 0) return true;              // 판정 불가 → 표시
            const int target = normYear(_wtoi(tag + 2));
            if (target <= 0) return true;
            if (wcsncmp(tag, L"LT", 2) == 0) return host < target;
            if (wcsncmp(tag, L"GT", 2) == 0) return host > target;
            return true;
        }

        // ── ribbon icons: RCDATA PNG → WIC → HICON → IPictureDisp ──────────
        std::map<std::wstring, IPictureDisp*> g_pics;   // main-thread only cache

        HMODULE thisModule()
        {
            HMODULE mod = nullptr;
            GetModuleHandleExW(
                GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                reinterpret_cast<LPCWSTR>(&thisModule), &mod);
            return mod;
        }

        IPictureDisp* createPictureFromPng(int resId)
        {
            HMODULE mod = thisModule();
            HRSRC hRes = FindResourceW(mod, MAKEINTRESOURCEW(resId), (LPCWSTR)RT_RCDATA);
            if (!hRes) return nullptr;
            HGLOBAL hMem = LoadResource(mod, hRes);
            if (!hMem) return nullptr;
            const DWORD size = SizeofResource(mod, hRes);
            const BYTE* data = static_cast<const BYTE*>(LockResource(hMem));
            if (!data || !size) return nullptr;

            IPictureDisp* pic = nullptr;
            IWICImagingFactory* fac = nullptr;
            IWICStream* stream = nullptr;
            IWICBitmapDecoder* dec = nullptr;
            IWICBitmapFrameDecode* frame = nullptr;
            IWICBitmapSource* src = nullptr;
            HBITMAP hbmColor = nullptr;
            do
            {
                if (FAILED(CoCreateInstance(CLSID_WICImagingFactory, nullptr,
                        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&fac)))) break;
                if (FAILED(fac->CreateStream(&stream))) break;
                if (FAILED(stream->InitializeFromMemory(
                        const_cast<BYTE*>(data), size))) break;
                if (FAILED(fac->CreateDecoderFromStream(stream, nullptr,
                        WICDecodeMetadataCacheOnDemand, &dec))) break;
                if (FAILED(dec->GetFrame(0, &frame))) break;
                if (FAILED(WICConvertBitmapSource(GUID_WICPixelFormat32bppPBGRA,
                        frame, &src))) break;

                UINT w = 0, h = 0;
                if (FAILED(src->GetSize(&w, &h)) || !w || !h) break;

                BITMAPINFO bi{};
                bi.bmiHeader.biSize = sizeof(bi.bmiHeader);
                bi.bmiHeader.biWidth = (LONG)w;
                bi.bmiHeader.biHeight = -(LONG)h;    // top-down
                bi.bmiHeader.biPlanes = 1;
                bi.bmiHeader.biBitCount = 32;
                bi.bmiHeader.biCompression = BI_RGB;
                void* bits = nullptr;
                hbmColor = CreateDIBSection(nullptr, &bi, DIB_RGB_COLORS,
                                            &bits, nullptr, 0);
                if (!hbmColor || !bits) break;
                if (FAILED(src->CopyPixels(nullptr, w * 4, w * h * 4,
                        static_cast<BYTE*>(bits)))) break;

                // HICON 경유(PICTYPE_ICON): 32bpp 알파가 리본에서 확실히 보존됨.
                HBITMAP hbmMask = CreateBitmap(w, h, 1, 1, nullptr);
                if (!hbmMask) break;
                ICONINFO ii{};
                ii.fIcon = TRUE;
                ii.hbmColor = hbmColor;
                ii.hbmMask = hbmMask;
                HICON icon = CreateIconIndirect(&ii);
                DeleteObject(hbmMask);
                if (!icon) break;

                PICTDESC pd{};
                pd.cbSizeofstruct = sizeof(pd);
                pd.picType = PICTYPE_ICON;
                pd.icon.hicon = icon;
                if (FAILED(OleCreatePictureIndirect(&pd, IID_IPictureDisp,
                        TRUE /*fOwn*/, reinterpret_cast<void**>(&pic))))
                {
                    DestroyIcon(icon);
                    pic = nullptr;
                }
            } while (false);

            if (hbmColor) DeleteObject(hbmColor);   // CreateIconIndirect가 복사함
            if (src) src->Release();
            if (frame) frame->Release();
            if (dec) dec->Release();
            if (stream) stream->Release();
            if (fac) fac->Release();
            return pic;
        }

        // control id → BtnDef 인덱스(모든 테이블 평면화, 최초 사용 시 1회 구축).
        const BtnDef* findButton(const wchar_t* ctrlId)
        {
            static std::map<std::wstring, const BtnDef*> index;   // main-thread only
            if (index.empty())
            {
                auto add = [&](const BtnDef& b) { index[b.id] = &b; };
                for (const auto& g : kGroups)
                    for (size_t c = 0; c < g.count; ++c)
                    {
                        if (g.controls[c].btn) add(*g.controls[c].btn);
                        else if (g.controls[c].split)
                        {
                            add(g.controls[c].split->defBtn);
                            for (size_t i = 0; i < g.controls[c].split->count; ++i)
                                add(g.controls[c].split->items[i]);
                        }
                    }
                for (const auto& m : kCtxMenus)
                    for (size_t i = 0; i < m.count; ++i)
                        add(m.items[i]);
                for (const auto& b : kInsertTabButtons)
                    add(b);
            }
            const auto it = index.find(ctrlId ? ctrlId : L"");
            return it == index.end() ? nullptr : it->second;
        }

        IPictureDisp* iconFor(const wchar_t* ctrlId)
        {
            const BtnDef* btn = findButton(ctrlId);
            const wchar_t* image = btn ? btn->image : nullptr;
            if (!image) return nullptr;

            auto it = g_pics.find(image);
            if (it != g_pics.end()) return it->second;

            int resId = 0;
            for (const auto& [name, id] : kIcons)
                if (wcscmp(name, image) == 0) { resId = id; break; }
            if (!resId) return nullptr;

            IPictureDisp* pic = createPictureFromPng(resId);
            if (pic) g_pics[image] = pic;            // cache holds one ref
            return pic;
        }

        void releaseIcons()
        {
            for (auto& [name, pic] : g_pics)
                if (pic) pic->Release();
            g_pics.clear();
        }

        constexpr const wchar_t* kRepoUrl = L"https://github.com/EGTools/EGToolsPP";
        constexpr const wchar_t* kReleasesUrl = L"https://github.com/EGTools/EGToolsPP/releases/latest";

        // GitHub 최신 릴리즈 태그("v0.2.40") 조회. 짧은 타임아웃, 실패 시 false.
        bool latestReleaseTag(std::wstring& tag)
        {
            HINTERNET hSession = WinHttpOpen(L"EGTools++",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (!hSession) return false;
            WinHttpSetTimeouts(hSession, 2000, 2000, 3000, 4000);

            std::string body;
            HINTERNET hConnect = WinHttpConnect(hSession, L"api.github.com",
                                                INTERNET_DEFAULT_HTTPS_PORT, 0);
            if (hConnect)
            {
                HINTERNET hReq = WinHttpOpenRequest(hConnect, L"GET",
                    L"/repos/EGTools/EGToolsPP/releases/latest", nullptr,
                    WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
                if (hReq)
                {
                    if (WinHttpSendRequest(hReq, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
                        WinHttpReceiveResponse(hReq, nullptr))
                    {
                        DWORD avail = 0;
                        do
                        {
                            avail = 0;
                            if (!WinHttpQueryDataAvailable(hReq, &avail) || avail == 0) break;
                            std::string chunk((size_t)avail, '\0');
                            DWORD read = 0;
                            if (!WinHttpReadData(hReq, chunk.data(), avail, &read)) break;
                            body.append(chunk.data(), read);
                        } while (avail > 0);
                    }
                    WinHttpCloseHandle(hReq);
                }
                WinHttpCloseHandle(hConnect);
            }
            WinHttpCloseHandle(hSession);

            const size_t k = body.find("\"tag_name\"");
            if (k == std::string::npos) return false;
            const size_t q1 = body.find('"', body.find(':', k));
            if (q1 == std::string::npos) return false;
            const size_t q2 = body.find('"', q1 + 1);
            if (q2 == std::string::npos) return false;
            const std::string t = body.substr(q1 + 1, q2 - q1 - 1);
            tag.assign(t.begin(), t.end());   // 태그는 ASCII
            return !tag.empty();
        }

        // 태그("v0.2.40" / "0.2.40") → (major, minor, build). 부족한 자리는 0.
        void parseTag(const std::wstring& tag, int& M, int& m, int& b)
        {
            M = m = b = 0;
            const wchar_t* p = tag.c_str();
            if (*p == L'v' || *p == L'V') ++p;
            swscanf_s(p, L"%d.%d.%d", &M, &m, &b);
        }

        // 새 릴리즈가 있으면 태그를 돌려준다.
        bool newVersionAvailable(std::wstring& tag)
        {
            if (!latestReleaseTag(tag)) return false;
            int M, m, b;
            parseTag(tag, M, m, b);
            const int cur[3] = { EG_VER_MAJOR, EG_VER_MINOR, EG_BUILD_NUM };
            const int rel[3] = { M, m, b };
            for (int i = 0; i < 3; ++i)
            {
                if (rel[i] > cur[i]) return true;
                if (rel[i] < cur[i]) return false;
            }
            return false;
        }

        // TaskDialog 하이퍼링크 클릭 → 기본 브라우저로 이동.
        HRESULT CALLBACK aboutDlgProc(HWND, UINT msg, WPARAM, LPARAM lParam, LONG_PTR)
        {
            if (msg == TDN_HYPERLINK_CLICKED && lParam)
                ShellExecuteW(nullptr, L"open", (LPCWSTR)lParam, nullptr, nullptr, SW_SHOWNORMAL);
            return S_OK;
        }

        // About: 모달 TaskDialog — 버전/비트수, 클릭 가능한 배포처 링크, 새 릴리즈 알림.
        // (comctl32 v6의 TaskDialogIndirect를 동적 로드 — 없으면 MessageBox 폴백.)
        // 문구는 i18n 카탈로그 "ribbon.about.*" 키.
        void showAbout()
        {
            using egtools::i18n::t;
            const std::wstring aboutTitle = t(L"ribbon.egpp.about.label");
            const wchar_t* bits = (sizeof(void*) == 8) ? L"64" : L"32";
            const HWND owner = (HWND)xloil::Environment::excelProcess().hWnd;

            std::wstring title = L"EGTools++ " EG_WIDEN(EG_VERSION_STR) L" (";
            title += bits;
            title += L"-bit)";

            std::wstring newVerTag;
            const bool hasNew = newVersionAvailable(newVerTag);

            using TDI = HRESULT(WINAPI*)(const TASKDIALOGCONFIG*, int*, int*, BOOL*);
            HMODULE hCom = GetModuleHandleW(L"comctl32.dll");
            if (!hCom) hCom = LoadLibraryW(L"comctl32.dll");
            const TDI taskDialog = hCom ? (TDI)GetProcAddress(hCom, "TaskDialogIndirect")
                                        : nullptr;
            if (taskDialog)
            {
                std::wstring content = t(L"ribbon.about.subtitle");
                content += L"\n\n";
                content += t(L"ribbon.about.dist");
                content += L": <a href=\"";
                content += kRepoUrl;
                content += L"\">github.com/EGTools/EGToolsPP</a>\n";
                content += t(L"ribbon.about.license");
                content += L": Apache License 2.0\n";
                content += t(L"ribbon.about.base");
                content += L": xlOil 0.22.1 (Apache-2.0)";
                if (hasNew)
                {
                    wchar_t line[256];
                    swprintf_s(line, t(L"ribbon.about.newver").c_str(), newVerTag.c_str());
                    content += L"\n\n<a href=\"";
                    content += kReleasesUrl;
                    content += L"\">";
                    content += line;
                    content += L"</a>";
                }

                TASKDIALOGCONFIG cfg{};
                cfg.cbSize = sizeof(cfg);
                cfg.hwndParent = owner;
                cfg.dwFlags = TDF_ENABLE_HYPERLINKS | TDF_ALLOW_DIALOG_CANCELLATION |
                              TDF_SIZE_TO_CONTENT;
                cfg.dwCommonButtons = TDCBF_OK_BUTTON;
                cfg.pszWindowTitle = aboutTitle.c_str();
                cfg.pszMainIcon = TD_INFORMATION_ICON;
                cfg.pszMainInstruction = title.c_str();
                cfg.pszContent = content.c_str();
                cfg.pfCallback = aboutDlgProc;
                taskDialog(&cfg, nullptr, nullptr, nullptr);
                return;
            }

            // 폴백: 링크 없는 일반 메시지 상자.
            std::wstring text = title + L"\n";
            text += t(L"ribbon.about.subtitle");
            text += L"\n\n";
            text += t(L"ribbon.about.dist");
            text += L": ";
            text += kRepoUrl;
            text += L"\n";
            text += t(L"ribbon.about.license");
            text += L": Apache License 2.0\n";
            text += t(L"ribbon.about.base");
            text += L": xlOil 0.22.1 (Apache-2.0)";
            if (hasNew)
            {
                wchar_t line[256];
                swprintf_s(line, t(L"ribbon.about.newver").c_str(), newVerTag.c_str());
                text += L"\n\n";
                text += line;
            }
            MessageBoxW(owner, text.c_str(), aboutTitle.c_str(), MB_OK | MB_ICONINFORMATION);
        }

        // Maps customUI callback names to C++ callbacks. Office invokes them as
        // name(control[, args…]); command callbacks ignore the args, resource
        // callbacks (GetLabel 등) write their answer into pvarResult.
        IComAddin::RibbonCallback callbackFor(const wchar_t* name)
        {
            const std::wstring n = name ? name : L"";

            // ── resource callbacks ──
            if (n == L"GetLabel")
                return [](const RibbonControl& c, VARIANT* r, int, VARIANT**)
                       { setResult(r, ctrlText(c, L".label")); };
            if (n == L"GetSupertip")
                return [](const RibbonControl& c, VARIANT* r, int, VARIANT**)
                       { setResult(r, ctrlText(c, L".tip")); };
            if (n == L"GetVisible")
                return [](const RibbonControl& c, VARIANT* r, int, VARIANT**)
                       {
                           if (!r) return;
                           r->vt = VT_BOOL;
                           r->boolVal = visibleForTag(c.Tag) ? VARIANT_TRUE
                                                             : VARIANT_FALSE;
                       };
            if (n == L"GetImage")
                return [](const RibbonControl& c, VARIANT* r, int, VARIANT**)
                       {
                           if (!r) return;
                           IPictureDisp* pic = iconFor(c.Id);
                           if (!pic) return;
                           pic->AddRef();          // 호출자(Office) 몫의 참조
                           r->vt = VT_DISPATCH;
                           r->pdispVal = pic;
                       };

            // ── commands ──
            using namespace egtools::commands;
            struct Cmd { const wchar_t* name; void (*fn)(); };
            static constexpr Cmd kCommands[] = {
                { L"onToCompat",     &convertToCompat },
                { L"onToNative",     &convertToNative },
                { L"onAbout",        &showAbout },
                { L"onVisCopy",      &copyVisibleCells },
                { L"onCopyAll",      &copyAllCells },
                { L"onPasteValues",  &pasteVisibleValues },
                { L"onPasteFormulas", &pasteVisibleFormulas },
                { L"onPasteAll",     &pasteVisibleAll },
                { L"onMergeClone",   &mergeClone },
                { L"onMergeFake",    &fakeMerge },
                { L"onMergeFakeClone", &fakeMergeClone },
                { L"onMergeJoin",    &mergeJoinText },
                { L"onMergeJoinAll", &mergeJoinAll },
                { L"onMergeRow",     &mergeRowText },
                { L"onMergeRowAll",  &mergeRowAll },
                { L"onMergeCol",     &mergeColumnText },
                { L"onMergeColAll",  &mergeColumnAll },
                { L"onSplitFill",    &splitFill },
                { L"onSplitRow",     &splitRowText },
                { L"onSplitRowAll",  &splitRowAll },
                { L"onSplitCol",     &splitColumnText },
                { L"onSplitColAll",  &splitColumnAll },
                { L"onExportPics",   &exportAllPictures },
                { L"onDrawShape",    &drawFreeForm },
                { L"onDatePicker",   &pickDateIntoCell },
                { L"onPicInsert",    &insertPicture },
                { L"onPicFolder",    &insertPicturesFromFolder },
                { L"onPicForm",      &insertPicturesIntoForm },
                { L"onFitCell",      &pictureFitToCell },
                { L"onFitAll",       &pictureFitAll },
                { L"onCalMonth",     &makeMonthlyCalendar },
                { L"onCalYearA",     &makeAnnualCalendarA },
                { L"onCalYearB",     &makeAnnualCalendarB },
                { L"onSchWeek",      &makeWeeklyCalendar },
                { L"onSchDayA",      &makeDailyCalendarA },
                { L"onSchDayB",      &makeDailyCalendarB },
                { L"onLabels",       &makeLabels },
                { L"onMailMerge",    &mailMerge },
                { L"onDeleteStyles", &deleteStyles },
                { L"onDeleteNames",  &deleteNames },
                { L"onClearError",   &clearErrors },
                { L"onClearEmpty",   &removeEmptyStrings },
                { L"onMemoArrange",  &rearrangeMemo },
                { L"onBorderColor",  &changeLineColors },
                { L"onUdfFreeze",    &udfFreeze },
                { L"onCheckBox",     &applyCheckBox },
                { L"onRecalc",       &recalcAll },
                { L"onApiKeys",      &manageApiKeys },
                { L"onSmtpSetup",    &smtpSettings },
                { L"onOpenUrl",      &openUrlInCell },
            };
            for (const auto& c : kCommands)
                if (n == c.name)
                {
                    auto fn = c.fn;
                    return [fn](const RibbonControl&, VARIANT*, int, VARIANT**) { fn(); };
                }
            return IComAddin::RibbonCallback();
        }
    }

    void install()
    {
        // When Excel auto-loads the .xll at startup, xlAutoOpen runs before the
        // COM/ribbon machinery is ready — a direct connect fails silently and
        // the tab never appears (plan/14 §6). Defer to the WINDOW queue (the
        // only queue that runs in this static XLL) and retry until Excel
        // accepts the COM add-in connection: return false → xlOil reschedules
        // after waitBetweenRetries.
        g_shutdown = false;
        g_attempts = 0;
        xloil::detail::runExcelThreadImpl(
            []() -> bool
            {
                if (g_shutdown || g_addin)
                    return true;            // done (or cancelled) — stop retrying
                ++g_attempts;
                try
                {
                    if (COM::connectCom())
                    {
                        auto addin = makeComAddin(L"EGTools++", L"EGTools++ 호환 변환 도구");
                        const std::wstring xml = ribbonXml();
                        addin->connect(xml.c_str(), callbackFor);
                        g_addin = std::move(addin);
                        return true;
                    }
                }
                catch (...) {}
                // connectCom() needs a workbook window (it binds via the EXCEL7
                // child of XLMAIN), which does NOT exist while Excel sits on the
                // Start screen — the whole reason a direct install at xlAutoOpen
                // never worked. Keep retrying so the tab appears as soon as the
                // user opens/creates a workbook; give up after ~10 min.
                if (g_attempts >= kMaxAttempts)
                    return true;            // give up quietly
                return false;               // false → xlOil requeues us
            },
            ExcelRunQueue::WINDOW | ExcelRunQueue::ENQUEUE,
            /*waitBeforeCall*/ 200, /*waitBetweenRetries*/ 1000);
    }

    void uninstall()
    {
        g_shutdown = true;   // cancel any pending deferred install
        try
        {
            if (g_addin)
            {
                g_addin->disconnect();
                g_addin->close();
            }
        }
        catch (...) {}
        g_addin.reset();
        releaseIcons();
    }
}
