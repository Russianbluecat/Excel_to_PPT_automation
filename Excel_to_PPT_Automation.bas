Attribute VB_Name = "Module11"
Option Explicit

'==============================
' 전체 슬라이드 업데이트 메인 함수
' 목적: Excel 데이터를 PowerPoint 여러 슬라이드에 자동 복사
'==============================
Sub Update_All_Slides()
    On Error GoTo ErrorHandler
    
    ' ========== 변수 선언 ==========
    Dim pptApp As Object        ' PowerPoint 애플리케이션 객체
    Dim pptPres As Object       ' PowerPoint 프레젠테이션 객체
    Dim ws As Worksheet         ' Excel 워크시트 객체
    Dim pptPath As String       ' PowerPoint 파일 경로
    
    ' ========== Excel 시트 설정 ==========
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Copy_to_Slide")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "Copy_to_Slide 시트를 찾을 수 없습니다."
        Exit Sub
    End If
    
    ' ========== PPT 파일 경로 설정 ==========
    pptPath = "C:\Users\YourName\Documents\sample.pptx"
    
    If Dir(pptPath) = "" Then
        MsgBox "PPT 파일을 찾을 수 없습니다: " & vbCrLf & pptPath
        Exit Sub
    End If
    
    ' ========== PowerPoint 실행 및 파일 열기 ==========
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Open(pptPath)
    
    Debug.Print "========== 전체 슬라이드 업데이트 시작 =========="
    
    ' ========== 각 슬라이드 업데이트 (함수 분리) ==========
    Call Update_Slide1(pptPres, ws)  ' Slide 1 처리
    Call Update_Slide2(pptPres, ws)  ' Slide 2 처리
    ' 추가 슬라이드가 있으면 여기에 추가
    ' Call Update_Slide3(pptPres, ws)
    ' Call Update_Slide4(pptPres, ws)
    
    Debug.Print "========== 전체 슬라이드 업데이트 완료 =========="
    
    ' ========== 파일 저장 ==========
    pptPres.Save
    MsgBox "모든 슬라이드 업데이트 완료!"
    
CleanUp:
    ' ========== 리소스 정리 ==========
    On Error Resume Next
    Application.CutCopyMode = False
    If Not pptPres Is Nothing Then pptPres.Close
    If Not pptApp Is Nothing Then pptApp.Quit
    Set pptPres = Nothing
    Set pptApp = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "오류 발생: " & Err.Description & vbCrLf & "번호: " & Err.Number
    Resume CleanUp
End Sub

'==============================
' Slide 1 업데이트
' 매핑: Text_01(A5), Pie_01(A9:B12), Bar_01(A16:C20), Text_02(A24:G25)
'==============================
Sub Update_Slide1(pptPres As Object, ws As Worksheet)
    On Error Resume Next
    
    Dim pptSlide As Object
    Dim shp As Object
    Dim r As Long, c As Long
    Dim srcRange As Range
    
    Set pptSlide = pptPres.Slides(1)  ' 첫 번째 슬라이드 선택
    Debug.Print "--- Slide 1 업데이트 시작 ---"
    
    ' --- Slide_1_Text_01: A5 복사 ---
    Set shp = pptSlide.Shapes("Slide_1_Text_01")
    If Not shp Is Nothing Then
        ws.Range("A5").Copy
        shp.TextFrame.TextRange.Paste
        Debug.Print "Slide_1_Text_01 완료"
    End If
    Application.CutCopyMode = False
    
    ' --- Slide_1_Pie_01: A9:B12 복사 ---
    Set shp = pptSlide.Shapes("Slide_1_Pie_01")
    If Not shp Is Nothing Then
        With shp.Chart.ChartData
            .Activate
            With .Workbook.Worksheets(1)
                .Cells.Clear
                ws.Range("A9:B12").Copy
                .Range("A1").PasteSpecial xlPasteValues
            End With
            .Workbook.Close
        End With
        shp.Chart.Refresh
        Debug.Print "Slide_1_Pie_01 완료"
    End If
    Application.CutCopyMode = False
    
    ' --- Slide_1_Bar_01: A16:C20 복사 ---
    Set shp = pptSlide.Shapes("Slide_1_Bar_01")
    If Not shp Is Nothing Then
        With shp.Chart.ChartData
            .Activate
            With .Workbook.Worksheets(1)
                .Cells.Clear
                ws.Range("A16:C20").Copy
                .Range("A1").PasteSpecial xlPasteValues
            End With
            .Workbook.Close
        End With
        shp.Chart.Refresh
        Debug.Print "Slide_1_Bar_01 완료"
    End If
    Application.CutCopyMode = False
    
    ' --- Slide_1_Text_02: A24:G25 복사 (2행 6열 표) ---
    Set shp = pptSlide.Shapes("Slide_1_Text_02")
    If Not shp Is Nothing And shp.HasTable Then
        Set srcRange = ws.Range("A24:G25")
        For r = 1 To 2
            For c = 1 To 6
                If c <= shp.Table.Columns.Count And r <= shp.Table.Rows.Count Then
                    shp.Table.Cell(r, c).Shape.TextFrame.TextRange.Text = srcRange.Cells(r, c).Value
                End If
            Next c
        Next r
        Debug.Print "Slide_1_Text_02 완료"
    End If
    
    Debug.Print "--- Slide 1 업데이트 완료 ---"
End Sub

'==============================
' Slide 2 업데이트
' 매핑: Text_01(A33), Line_01(A37:D41), Bar_01(A44:C48)
'==============================
Sub Update_Slide2(pptPres As Object, ws As Worksheet)
    On Error Resume Next
    
    Dim pptSlide As Object
    Dim shp As Object
    
    Set pptSlide = pptPres.Slides(2)  ' 두 번째 슬라이드 선택
    Debug.Print "--- Slide 2 업데이트 시작 ---"
    
    ' --- Slide_2_Text_01: A33 복사 ---
    Set shp = pptSlide.Shapes("Slide_2_Text_01")
    If Not shp Is Nothing Then
        ws.Range("A33").Copy
        shp.TextFrame.TextRange.Paste
        Debug.Print "Slide_2_Text_01 완료"
    End If
    Application.CutCopyMode = False
    
    ' --- Slide_2_Line_01: A37:D41 복사 (선 차트) ---
    Set shp = pptSlide.Shapes("Slide_2_Line_01")
    If Not shp Is Nothing Then
        With shp.Chart.ChartData
            .Activate
            With .Workbook.Worksheets(1)
                .Cells.Clear
                ws.Range("A37:D41").Copy
                .Range("A1").PasteSpecial xlPasteValues
            End With
            .Workbook.Close
        End With
        shp.Chart.Refresh
        Debug.Print "Slide_2_Line_01 완료"
    End If
    Application.CutCopyMode = False
    
    ' --- Slide_2_Bar_01: A44:C48 복사 (막대 차트) ---
    Set shp = pptSlide.Shapes("Slide_2_Bar_01")
    If Not shp Is Nothing Then
        With shp.Chart.ChartData
            .Activate
            With .Workbook.Worksheets(1)
                .Cells.Clear
                ws.Range("A44:C48").Copy
                .Range("A1").PasteSpecial xlPasteValues
            End With
            .Workbook.Close
        End With
        shp.Chart.Refresh
        Debug.Print "Slide_2_Bar_01 완료"
    End If
    Application.CutCopyMode = False
    
    Debug.Print "--- Slide 2 업데이트 완료 ---"
End Sub

'==============================
' Slide 3 업데이트 (템플릿)
' 새 슬라이드 추가 시 이 함수를 복사해서 수정하세요
'==============================
Sub Update_Slide3(pptPres As Object, ws As Worksheet)
    On Error Resume Next
    
    Dim pptSlide As Object
    Dim shp As Object
    
    Set pptSlide = pptPres.Slides(3)  ' 세 번째 슬라이드 선택
    Debug.Print "--- Slide 3 업데이트 시작 ---"
    
    ' 여기에 Slide 3 업데이트 코드 추가
    ' 예시:
    ' Set shp = pptSlide.Shapes("Slide_3_Text_01")
    ' If Not shp Is Nothing Then
    '     ws.Range("A50").Copy
    '     shp.TextFrame.TextRange.Paste
    ' End If
    ' Application.CutCopyMode = False
    
    Debug.Print "--- Slide 3 업데이트 완료 ---"
End Sub
