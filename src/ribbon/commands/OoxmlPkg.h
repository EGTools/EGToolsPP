// OoxmlPkg.h — 저장된 통합문서 패키지(.xlsx = zip) 읽기 공용 헬퍼.
// 그림 추출(CmdPicture)과 셀 내 목록 변환(CmdListCells)이 공유한다.
// 외부 zip/XML 라이브러리 없이 Shell.Application 해제 + 문자열 탐색으로
// 처리한다(의존성 추가 금지 원칙).
#pragma once
#include <filesystem>
#include <string>

namespace egtools::commands::ooxml
{
    // [from,to) 구간에서 open…close 사이 텍스트(첫 일치).
    std::string tagBetween(const std::string& s, size_t from, size_t to,
                           const std::string& open, const std::string& close);

    // 요소 태그 구간 [from,to)에서 attr="값"(첫 일치). 없으면 "".
    std::string attrValue(const std::string& s, size_t from, size_t to,
                          const std::string& attr);

    std::string readFileBytes(const std::filesystem::path& p);

    // UTF-8 → UTF-16.
    std::wstring widen(const std::string& s);

    // XML 문자 참조 복원(&amp; &lt; &gt; &quot; &apos; &#NN; &#xHH;).
    std::string unescapeXml(const std::string& s);

    // Shell COM으로 ZIP 해제(완료 대기 폴링, 최대 30초).
    bool shellUnzip(const std::wstring& zipPath, const std::wstring& destDir);

    // workbook.xml.rels 등에서 Id 또는 Type(부분 일치)으로 Target을 찾는다.
    // rid가 비면 Type만으로, typeContains가 nullptr이면 Id만으로 매칭.
    std::string relTarget(const std::filesystem::path& relsFile,
                          const std::string& rid, const char* typeContains);
}
