// Commands.h — 리본 명령 진입점 (plan/23 M2~). Ribbon.cpp의 onAction 콜백이
// 호출한다. 전부 Excel 메인 스레드에서 실행되며 내부에서 예외를 처리한다
// (밖으로 던지지 않음).
#pragma once

namespace egtools::commands
{
    // ── 보이는 셀 (CmdCopy.cpp) ──
    void copyVisibleCells();   // 보이는 셀만 복사
    void copyAllCells();       // 전체 복사(원본 기억)
    void pasteVisibleValues();   // 보이는 셀에 값 붙여넣기
    void pasteVisibleFormulas(); // 보이는 셀에 수식 붙여넣기
    void pasteVisibleAll();      // 보이는 셀에 모두 붙여넣기

    // ── 병합/나누기 (CmdMerge.cpp) ──
    void splitFill();          // 나누고 채우기
    void mergeClone();         // 연속값 병합(계층형)
    void mergeJoinText();      // 서식없이 내용병합
    void mergeJoinAll();       // 서식대로 내용병합
    void mergeRowText();       // 서식없이 행끼리 병합
    void mergeRowAll();        // 서식대로 행끼리 병합
    void mergeColumnText();    // 서식없이 열끼리 병합
    void mergeColumnAll();     // 서식대로 열끼리 병합
    void fakeMerge();          // 내용감춘 병합(값 유지)
    void fakeMergeClone();     // 연속값 감춘 병합
    void splitRowText();       // 서식없이 행 나누기
    void splitRowAll();        // 서식대로 행 나누기
    void splitColumnText();    // 서식없이 열 나누기
    void splitColumnAll();     // 서식대로 열 나누기

    // ── 그림/사진 (CmdPicture.cpp) ──
    void insertPicture();      // 이미지 삽입(선택 영역 맞춤)
    void insertPicturesFromFolder(); // 셀 값=파일명 매칭 폴더 삽입
    void insertPicturesIntoForm();   // {{}} 양식 복제 삽입
    void pictureFitToCell();   // 선택 이미지 셀 맞춤
    void pictureFitAll();      // 시트 전체 이미지 셀 맞춤
    void exportAllPictures();  // 시트 그림 전체를 폴더로 추출(원본 화질)

    // ── 도구 (CmdUtil.cpp) ──
    void applyCheckBox();      // CheckBox 셀서식 적용
    void recalcAll();          // 전체 재계산
    void openUrlInCell();      // 셀 내용의 URL 열기 (컨텍스트 메뉴 포함)
    void rearrangeMemo();      // 메모 정돈
    void deleteNames();        // 오류/숨김 이름 삭제
    void changeLineColors();   // 테두리 색 일괄 변경
    void deleteStyles();       // 미사용/전체 사용자 스타일 삭제
    void clearErrors();        // 오류 수식 IFERROR 감싸기
    void removeEmptyStrings(); // 길이 0 문자열 셀 → 빈 셀
    void udfFreeze();          // EGTools UDF 수식 → 값 고정(배포용)
    void drawFreeForm();       // 색칠 셀 외곽 자유형 도형 생성

    // ── 날짜 선택 (DatePicker.cpp) ──
    void pickDateIntoCell();   // 달력에서 날짜를 골라 현재 셀에 입력(컨텍스트 메뉴)

    // ── 달력/일정표 (CmdCalendar.cpp) ──
    void makeMonthlyCalendar();
    void makeAnnualCalendarA();
    void makeAnnualCalendarB();
    void makeWeeklyCalendar();
    void makeDailyCalendarA();
    void makeDailyCalendarB();

    // ── 양식 (CmdForm.cpp) ──
    void makeLabels();      // 라벨지 만들기
    void mailMerge();       // 메일머지 — 시트 복제·치환·파일저장(xlsx/PDF)·인쇄·
                            //   메일 발송(v2, 옵션 비트 4 — Smtp.h/CmdSmtp.cpp)

    // ── API 키 관리 (CmdApiKeys.cpp) ──
    void manageApiKeys();      // 공공 API 키 등록/삭제 다이얼로그

    // ── SMTP 설정 (CmdSmtp.cpp) ──
    void smtpSettings();       // 메일 발송 SMTP 설정 다이얼로그(암호 제외 저장)
}
