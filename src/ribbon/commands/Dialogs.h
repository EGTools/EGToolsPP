// Dialogs.h — 리본 명령 공용 Win32 다이얼로그 (plan/23 §3-1).
// 전부 Excel 메인 스레드에서 모달로 띄운다(부모 = Excel hWnd).
#pragma once
#include <windows.h>
#include <string>

namespace egtools::dialogs
{
    // 1줄 텍스트 입력. 확인=true(값은 value), 취소=false.
    bool inputText(HWND owner, const std::wstring& title,
                   const std::wstring& label, std::wstring& value,
                   bool password = false);

    // 숫자 입력(inputText 래퍼). 숫자가 아니면 false(VB InputBox 규약과 동일:
    // "숫자 이외의 값을 입력하면 종료"). 반환값은 [minV, maxV]로 클램프.
    bool inputNumber(HWND owner, const std::wstring& title,
                     const std::wstring& label, double& value,
                     double minV, double maxV, double defV);

    // 색 선택(ChooseColor 공용 대화상자). 확인=true(color=COLORREF).
    bool pickColor(HWND owner, COLORREF& color);

    // 이미지 파일 열기(GetOpenFileName). 확인=true(path=전체 경로).
    bool pickImageFile(HWND owner, std::wstring& path);

    // 폴더 선택(IFileOpenDialog + FOS_PICKFOLDERS). 확인=true(path, 후행 \ 없음).
    bool pickFolder(HWND owner, std::wstring& path);
}
