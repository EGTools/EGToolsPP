// Smtp.h — 메일머지 v2 메일 발송 공용 인터페이스 (plan/23 §3-3, CmdSmtp.cpp).
//
// SMTP 설정은 암호를 제외하고 HKCU\Software\EGTools\Smtp에 저장한다(사용자
// 지시, 2026-08-03). 암호는 저장하지 않으며 발송 시마다 입력받아 메모리에서만
// 사용한다(VB의 CDOsmtp\UserPass 임시 저장 방식 폐기).
#pragma once
#include <string>

namespace egtools::commands::smtp
{
    struct Settings
    {
        std::wstring name;      // 발송인 이름
        std::wstring email;     // 발송인 메일주소(= SMTP 로그인 계정)
        std::wstring server;    // 보내는 메일서버
        std::wstring port;      // 서버 포트(기본 465)
        bool         ssl = true;
    };

    Settings load();                       // 레지스트리에서 읽기(없으면 기본값)
    bool complete(const Settings& s);      // 발송에 필요한 값이 다 있는가
    bool ensure(Settings& s);              // 미비 시 설정 다이얼로그 → 재확인

    // CDO.Message(late-bound COM)로 1건 발송. 성공 시 빈 문자열, 실패 시 오류
    // 설명(EXCEPINFO description 포함)을 반환한다. cc/bcc/attachment는 빈 값
    // 허용. 수신자 구분자는 CDO 규약(',' 또는 ';')을 그대로 따른다.
    std::wstring send(const Settings& s, const std::wstring& password,
                      const std::wstring& to, const std::wstring& cc,
                      const std::wstring& bcc, const std::wstring& subject,
                      const std::wstring& body, const std::wstring& attachment);
}
