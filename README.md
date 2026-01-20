# 📊 Excel to PowerPoint Automation Tool

Excel 데이터를 PowerPoint 슬라이드에 자동으로 복사-붙여넣기하는 VBA 매크로입니다.

## ✨ 주요 기능

- 📝 **텍스트 자동 업데이트**: Excel 셀 내용을 PPT 텍스트박스에 자동 복사
- 📊 **차트 데이터 업데이트**: 파이 차트, 막대 차트, 선 차트 등 모든 차트 지원
- 📋 **표 데이터 업데이트**: Excel 범위를 PPT 표에 자동 복사
- 🔄 **다중 슬라이드 지원**: 여러 슬라이드를 한 번에 업데이트
- 🎯 **확장 가능한 구조**: 새 슬라이드 추가 시 쉽게 확장 가능
- 🛡️ **에러 처리**: 파일 없음, 객체 없음 등 예외 상황 대응

## 🚀 빠른 시작

### 1. 파일 준비
- Excel 파일에 `Copy_to_Slide` 시트 생성
- PowerPoint 파일 준비
- PPT 객체에 이름 지정 (예: `Slide_1_Text_01`)

### 2. VBA 코드 설치
1. Excel 파일에서 `Alt + F11`로 VBA 편집기 열기
2. `파일` → `가져오기` 선택
3. `Excel_to_PPT_Automation.bas` 파일 선택
4. 또는 새 모듈을 만들고 코드 복사-붙여넣기

### 3. 설정 수정
```vba
' 코드에서 PPT 파일 경로 수정
pptPath = "여기에_실제_파일_경로_입력.pptx"
```

### 4. 실행
- `Alt + F8`로 매크로 대화상자 열기
- `Update_All_Slides` 선택 후 실행

## 📋 사용 예시

### Excel 데이터 구조 (Copy_to_Slide 시트)
```
A5: 2025 12월 누적 매출 Trend

A9:B12 (파이 차트)
제품 A    45.1%
제품 B    24.0%
제품 C    30.9%

A16:C20 (막대 차트)
         2026   2025
제품 A    4.3    2.4
제품 B    2.5    4.4
...
```

### PowerPoint 객체 명명 규칙
- 텍스트박스: `Slide_1_Text_01`, `Slide_1_Text_02`, ...
- 차트: `Slide_1_Pie_01`, `Slide_1_Bar_01`, `Slide_2_Line_01`, ...
- 표: `Slide_1_Text_02` (HasTable 속성 사용)

## 🗺️ 매핑 구조

### Slide 1
| PPT Object | Excel Range | Type | 
|------------|-------------|------|
| Slide_1_Text_01 | A5 | 텍스트박스 |
| Slide_1_Pie_01 | A9:B12 | 파이 차트 | 
| Slide_1_Bar_01 | A16:C20 | 막대 차트 | 
| Slide_1_Text_02 | A24:G25 | 표 (2행 6열) | 

### Slide 2
| PPT Object | Excel Range | Type | 
|------------|-------------|------|
| Slide_2_Text_01 | A33 | 텍스트박스 | 
| Slide_2_Line_01 | A37:D41 | 선 차트 | 
| Slide_2_Bar_01 | A44:C48 | 막대 차트 | 

## 🔧 새 슬라이드 추가하기

1. **함수 복사**
```vba
Sub Update_Slide3(pptPres As Object, ws As Worksheet)
    Dim pptSlide As Object
    Set pptSlide = pptPres.Slides(3)  ' 슬라이드 번호 변경
    
    ' 객체 업데이트 코드 추가
    ...
End Sub
```

2. **메인 함수에 추가**
```vba
Sub Update_All_Slides()
    ...
    Call Update_Slide1(pptPres, ws)
    Call Update_Slide2(pptPres, ws)
    Call Update_Slide3(pptPres, ws)  ' 추가!
    ...
End Sub
```

## 📌 주의사항

### PowerPoint 객체 이름 설정
1. PPT에서 객체 선택
2. `홈` → `선택` → `선택창` 열기
3. 객체 이름 변경 (예: `Slide_1_Text_01`)

### Excel 시트 이름
- 반드시 `Copy_to_Slide`로 지정
- 또는 코드에서 시트 이름 수정

### 파일 경로
- 절대 경로 사용 (예: `C:\Users\...\file.pptx`)
- 한글 경로 가능하지만 영문 경로 권장

## 🛠️ 트러블슈팅

### "Copy_to_Slide 시트를 찾을 수 없습니다"
→ Excel에 `Copy_to_Slide` 시트가 있는지 확인

### "PPT 파일을 찾을 수 없습니다"
→ 파일 경로가 정확한지 확인 

### "객체를 찾을 수 없음"
→ PPT 객체 이름이 정확히 일치하는지 확인
→ VBA 즉시 창(`Ctrl+G`)에서 디버그 메시지 확인

### 차트가 업데이트되지 않음
→ 차트 데이터 범위가 올바른지 확인
→ Excel 데이터 형식 확인 (텍스트 vs 숫자)

## 📚 코드 구조

```
Update_All_Slides()          # 메인 함수
├── PowerPoint 열기
├── Update_Slide1()          # Slide 1 처리
│   ├── Text 업데이트
│   ├── Chart 업데이트
│   └── Table 업데이트
├── Update_Slide2()          # Slide 2 처리
├── Update_Slide3()          # Slide 3 처리 (템플릿)
├── 저장
└── 리소스 정리
```


## 📄 라이선스

MIT License - 자유롭게 사용, 수정, 배포할 수 있습니다.


## 🙏 감사의 말

이 프로젝트는 반복적인 보고서 작성 업무를 자동화하기 위해 시작되었습니다.


Excel과 PowerPoint를 매일 사용하는 모든 직장인에게 도움이 되길 바랍니다!

---

⭐ 이 프로젝트가 도움이 되었다면 Star를 눌러주세요!
