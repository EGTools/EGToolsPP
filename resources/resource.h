// resource.h — resource identifiers for embedded i18n catalogs and ribbon icons.
#pragma once

#define IDR_I18N_EN     101
#define IDR_I18N_KO     102
#define IDR_I18N_ZH_CN  103
#define IDR_I18N_ZH_TW  104
#define IDR_I18N_JA     105
#define IDR_I18N_ES     106

// Ribbon icons — RCDATA PNG 64x64 (rendered from resources/icons/svg by
// tools/render_icons.ps1; both SVG source and PNG output are committed).
// 64px: Office가 고DPI 리본에서 축소 방향으로 스케일링해 선명(32px 임베드는
// 125~150% 배율에서 확대되며 계단 현상 — 2026-08-04 사용자 피드백).
#define IDR_ICON_COMPAT_APPLY    201
#define IDR_ICON_RESTORE_NATIVE  202
#define IDR_ICON_ABOUT           203
#define IDR_ICON_SPLIT_FILL      204
#define IDR_ICON_PICTURE_INSERT  205
#define IDR_ICON_FIT_CELL        206
#define IDR_ICON_FIT_ALL         207
#define IDR_ICON_RECALC          208
#define IDR_ICON_DELETE_NAMES    209
#define IDR_ICON_MEMO_ARRANGE    210
#define IDR_ICON_BORDER_COLOR    211
#define IDR_ICON_CHECKBOX        212
#define IDR_ICON_API_KEYS        213
#define IDR_ICON_MERGE_JOIN      214
#define IDR_ICON_MERGE_JOIN_ALL  215
#define IDR_ICON_MERGE_CLONE     216
#define IDR_ICON_MERGE_FAKE      217
#define IDR_ICON_MERGE_FAKE_CLONE 218
#define IDR_ICON_MERGE_ROW       219
#define IDR_ICON_MERGE_ROW_ALL   220
#define IDR_ICON_MERGE_COL       221
#define IDR_ICON_MERGE_COL_ALL   222
#define IDR_ICON_SPLIT_ROW       223
#define IDR_ICON_SPLIT_ROW_ALL   224
#define IDR_ICON_SPLIT_COL       225
#define IDR_ICON_SPLIT_COL_ALL   226
#define IDR_ICON_VISIBLE_COPY    227
#define IDR_ICON_PICTURE_FOLDER  228
#define IDR_ICON_PICTURE_FORM    229
#define IDR_ICON_CAL_MONTH       230
#define IDR_ICON_CAL_YEAR_A      231
#define IDR_ICON_CAL_YEAR_B      232
#define IDR_ICON_SCH_WEEK        233
#define IDR_ICON_SCH_DAY_A       234
#define IDR_ICON_SCH_DAY_B       235
#define IDR_ICON_LABELS          236
#define IDR_ICON_DELETE_STYLES   237
#define IDR_ICON_CLEAR_ERROR     238
#define IDR_ICON_CLEAR_EMPTY     239
#define IDR_ICON_UDF_FREEZE      240
#define IDR_ICON_PASTE_VALUES    241
#define IDR_ICON_PASTE_FORMULAS  242
#define IDR_ICON_PASTE_ALL       243
#define IDR_ICON_EXPORT_PICTURES 244
#define IDR_ICON_DRAW_SHAPE      245
#define IDR_ICON_DATE_PICKER     246
#define IDR_ICON_MAIL_MERGE      247
#define IDR_ICON_SMTP_SETUP      248
