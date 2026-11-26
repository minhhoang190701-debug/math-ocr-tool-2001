Attribute VB_Name = "ChuanHoaDeToan"
'========================================================================================
' MACRO VBA CHUẨN HÓA ĐỀ THI TOÁN HỌC TRONG MICROSOFT WORD (PHIÊN BẢN NÂNG CẤP)
' Tích hợp tinh hoa từ Math OCR Tool Team, BTPro (S_Standar/S_Math) & Macro cá nhân
' Chức năng chính:
'  1. ChuanHoaToanBoFileWord : Chuẩn hóa A4 NĐ30, TNR 12pt, giãn dòng 1.15, bảo toàn OMML Cambria Math,
'                               làm sạch dấu câu/dòng trống, in đậm Câu X: & A.B.C.D., căn ảnh/bảng.
'  2. CanDapAn_2x2            : Căn phương án trắc nghiệm A-D hoặc ý tự luận a)-z) thành 2 cột (Tab 8.5cm).
'  3. CanDapAn_4Cot           : Căn 4 phương án A. B. C. D. trên 1 dòng (Tab 0.5 / 4.5 / 8.5 / 12.5 cm).
'  4. XuongDongDapAn_4Dong    : Tách các phương án A-D hoặc ý a)-z) xuống từng dòng riêng.
'  5. DanhLaiThuTuCau         : Đánh lại số thứ tự Câu 1, Câu 2... tuần tự toàn bài.
'  6. AnLoiGiai / HienLoiGiai : Ẩn hoặc hiện nhanh các khối Lời giải / Hướng dẫn giải để in đề phát HS.
'========================================================================================

Option Explicit

Sub ChuanHoaToanBoFileWord()
    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    On Error GoTo ErrorHandler

    ' Chuyển toàn bộ đánh số tự động (ListNumber) về text thuần để tránh mất nhãn khi copy
    ActiveDocument.Range.ListFormat.ConvertNumbersToText

    ' 1. Căn lề trang in chuẩn Nghị định 30 (A4: Trên 2, Dưới 2, Trái 3, Phải 2)
    Call ThietLapTrangInA4

    ' 2. Làm sạch văn bản (font cũ C©u, ngắt dòng mềm ^l, khoảng trắng, dòng trống, dấu câu)
    Call LamSachKhoangTrangVaDauCau

    ' 3. Định dạng Font và Đoạn văn (Bảo toàn 100% công thức OMML Cambria Math)
    Call DinhDangFontVaDoanVan

    ' 4. In đậm có chọn lọc: Câu hỏi và Đáp án trắc nghiệm (bảo vệ "Chọn A/B/C/D")
    Call InDamNhanCauHoiVaDapAn

    ' 5. Căn giữa và chuẩn hóa kích thước hình ảnh
    Call ChuanHoaHinhAnh

    ' 6. Căn giữa và co giãn bảng biểu
    Call ChuanHoaBangBieu

    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll

    Dim totalTime As Double
    totalTime = Round(Timer - startTime, 2)

    MsgBox "CHUẨN HÓA ĐỀ TOÁN HOÀN TẤT!" & vbCrLf & _
           "- Trang in A4, lề chuẩn NĐ 30 (2-2-3-2 cm)." & vbCrLf & _
           "- Font Times New Roman 12pt, giãn dòng 1.15, bảo toàn Cambria Math (OMML)." & vbCrLf & _
           "- Đã làm sạch dấu câu, dòng trống & in đậm Câu X:, A. B. C. D." & vbCrLf & _
           "- Đã căn giữa hình ảnh, bảng biểu." & vbCrLf & _
           "- Thời gian xử lý: " & totalTime & " giây.", _
           vbInformation, "Chuẩn Hóa Đề Toán"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    MsgBox "Có lỗi xảy ra: " & Err.Description, vbCritical, "Lỗi VBA"
End Sub

'----------------------------------------------------------------------------------------
' 1. THIẾT LẬP TRANG IN KHỔ A4 THEO NGHỊ ĐỊNH 30/2020/NĐ-CP
'----------------------------------------------------------------------------------------
Private Sub ThietLapTrangInA4()
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        With sec.PageSetup
            .PaperSize = wdPaperA4
            .TopMargin = CentimetersToPoints(2#)
            .BottomMargin = CentimetersToPoints(2#)
            .LeftMargin = CentimetersToPoints(3#)
            .RightMargin = CentimetersToPoints(2#)
            .HeaderDistance = CentimetersToPoints(1.2)
            .FooterDistance = CentimetersToPoints(1.2)
        End With
    Next sec
End Sub

'----------------------------------------------------------------------------------------
' 2. LÀM SẠCH VĂN BẢN (FONT CŨ, NGẮT DÒNG MỀM, KHOẢNG TRẮNG, DÒNG TRỐNG, DẤU CÂU)
'----------------------------------------------------------------------------------------
Private Sub LamSachKhoangTrangVaDauCau()
    ' Sửa lỗi font TCVN3/VNI cũ hay gặp ở chữ Câu
    Call ThayTheChuoi("C©u", "Câu")
    Call ThayTheChuoi("Caâu", "Câu")

    ' Chuyển ngắt dòng mềm (^l / Shift+Enter) thành ngắt đoạn chuẩn (^p)
    Call ThayTheChuoi("^l", "^p")

    ' Xóa khoảng trắng thừa liên tiếp
    Call ThayTheChuoi("   ", " ")
    Call ThayTheChuoi("  ", " ")
    Call ThayTheChuoi("^p ", "^p")
    Call ThayTheChuoi(" ^p", "^p")

    ' Sửa lỗi dấu cách trước dấu câu
    Call ThayTheChuoi(" .", ".")
    Call ThayTheChuoi(" ,", ",")
    Call ThayTheChuoi(" :", ":")
    Call ThayTheChuoi(" ;", ";")
    Call ThayTheChuoi(" !", "!")
    Call ThayTheChuoi(" ?", "?")

    ' Xóa khoảng trắng sau mở ngoặc và trước đóng ngoặc
    Call ThayTheChuoi("( ", "(")
    Call ThayTheChuoi(" )", ")")
    Call ThayTheChuoi("[ ", "[")
    Call ThayTheChuoi(" ]", "]")

    ' Xóa dòng trống thừa liên tiếp
    Call ThayTheChuoi("^p^p^p", "^p")
    Call ThayTheChuoi("^p^p", "^p")
End Sub

'----------------------------------------------------------------------------------------
' 3. ĐỊNH DẠNG FONT CHỮ VÀ ĐOẠN VĂN (BẢO TOÀN CÔNG THỨC TOÁN OMML)
'----------------------------------------------------------------------------------------
Private Sub DinhDangFontVaDoanVan()
    Dim p As Paragraph
    Dim rng As Range
    Dim eq As OMath

    For Each p In ActiveDocument.Paragraphs
        If Not p.Range.Information(wdWithinTable) Then
            With p.Format
                .Alignment = wdAlignParagraphJustify
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.15)
                .SpaceBefore = 0
                .SpaceAfter = 3
            End With
        End If
    Next p

    ' Thiết lập Font chung cho văn bản
    Set rng = ActiveDocument.Content
    With rng.Font
        .Name = "Times New Roman"
        .Size = 12
        .ColorIndex = wdBlack
    End With

    ' Khôi phục / Bảo toàn font Cambria Math cho toàn bộ Equation OMML (từ BTPro Lechdong2)
    On Error Resume Next
    For Each eq In ActiveDocument.OMaths
        eq.Range.Font.Name = "Cambria Math"
    Next eq
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------------------------
' 4. IN ĐẬM CÓ CHỌN LỌC (BẢO VỆ "CHỌN A/B/C/D", IN ĐẬM CÂU HỎI & PHƯƠNG ÁN)
'----------------------------------------------------------------------------------------
Private Sub InDamNhanCauHoiVaDapAn()
    Dim rng As Range

    ' Bảo vệ cụm "Chọn A." -> "Chọn A" (tránh bị nhận nhầm là phương án A.)
    Call ThayTheChuoi("Ch" & ChrW(7885) & "n A.", "Ch" & ChrW(7885) & "n A")
    Call ThayTheChuoi("Ch" & ChrW(7885) & "n B.", "Ch" & ChrW(7885) & "n B")
    Call ThayTheChuoi("Ch" & ChrW(7885) & "n C.", "Ch" & ChrW(7885) & "n C")
    Call ThayTheChuoi("Ch" & ChrW(7885) & "n D.", "Ch" & ChrW(7885) & "n D")

    ' 4.1. Tìm và in đậm: "Câu X:", "Câu X.", "Bài X:", "Ví dụ X:"
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Replacement.Font.Italic = False
        .MatchWildcards = True

        .Text = "(<[Cc]âu [0-9]{1,4}[:.!?])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(<[Bb]ài [0-9]{1,4}[:.!?])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(<[Vv]í dụ [0-9]{1,4}[:.!?])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(<[Cc]âu [0-9]{1,4} \([!)]@\):)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With

    ' 4.2. Tìm và in đậm nhãn phương án trắc nghiệm A. B. C. D. và ý a) b) c)...
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Replacement.Font.Italic = False
        .MatchWildcards = True

        .Text = "(<[ABCD]\.)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll

        .Text = "(<[a-z]\))"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

'----------------------------------------------------------------------------------------
' 5. CHUẨN HÓA HÌNH ẢNH (CĂN GIỮA, CO GIÃN THEO LỀ A4)
'----------------------------------------------------------------------------------------
Private Sub ChuanHoaHinhAnh()
    Dim shp As InlineShape
    Dim maxWidth As Single
    maxWidth = CentimetersToPoints(15.5)

    For Each shp In ActiveDocument.InlineShapes
        shp.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        shp.Range.ParagraphFormat.SpaceBefore = 4
        shp.Range.ParagraphFormat.SpaceAfter = 4
        shp.LockAspectRatio = msoTrue

        If shp.Width > maxWidth Then
            shp.Width = maxWidth
        End If
        If shp.Height > CentimetersToPoints(6.5) Then
            shp.Height = CentimetersToPoints(5#)
        End If
    Next shp
End Sub

'----------------------------------------------------------------------------------------
' 6. CHUẨN HÓA BẢNG BIỂU (CĂN GIỮA, CO GIÃN TỰ ĐỘNG)
'----------------------------------------------------------------------------------------
Private Sub ChuanHoaBangBieu()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        With tbl
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior wdAutoFitContent
            .TopPadding = CentimetersToPoints(0.1)
            .BottomPadding = CentimetersToPoints(0.1)
            .LeftPadding = CentimetersToPoints(0.15)
            .RightPadding = CentimetersToPoints(0.15)
        End With
    Next tbl
End Sub

'----------------------------------------------------------------------------------------
' 7. CĂN ĐÁP ÁN 2 CỘT (2x2 CHO A-D HOẶC 2 CỘT CHO TỰ LUẬN NHIỀU Ý a)-z))
'    Dùng trên vùng đang bôi đen (nếu có chọn) hoặc toàn văn bản
'----------------------------------------------------------------------------------------
Sub CanDapAn_2x2()
    Dim rng As Range
    Set rng = GetTargetRange()

    Application.ScreenUpdating = False
    ' Đưa các phương án đang dính về 4 dòng trước cho sạch
    Call SplitChoicesInRange(rng)

    With rng.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphLeft
        .TabStops.ClearAll
        .TabStops.Add Position:=CentimetersToPoints(0.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        .TabStops.Add Position:=CentimetersToPoints(8.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With

    ' Thêm Tab trước A. hoặc a) đầu dòng
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .MatchCase = True
        .Wrap = wdFindStop

        .Text = "(^13)([Aa][.,\)])"
        .Replacement.Text = "\1^t\2"
        .Execute Replace:=wdReplaceAll

        ' Ghép B. / b) lên cùng dòng với A. / a), C. / c) xuống dòng, D. / d) lên cùng dòng C. / c)
        .MatchCase = False
        .Text = "(^13)([b-z][.,\)])"
        .Replacement.Text = "^t\2"
        .Execute Replace:=wdReplaceAll

        .Text = "([b-z][.,\)]*)(^t)([c-z][.,\)])"
        .Replacement.Text = "\1^p^t\3"
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True
End Sub

'----------------------------------------------------------------------------------------
' 8. CĂN 4 ĐÁP ÁN TRÊN 1 DÒNG (4 CỘT TAB STOP: 0.5 / 4.5 / 8.5 / 12.5 cm)
'----------------------------------------------------------------------------------------
Sub CanDapAn_4Cot()
    Dim rng As Range
    Set rng = GetTargetRange()

    Application.ScreenUpdating = False
    Call SplitChoicesInRange(rng)

    With rng.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .Alignment = wdAlignParagraphLeft
        .TabStops.ClearAll
        .TabStops.Add Position:=CentimetersToPoints(0.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        .TabStops.Add Position:=CentimetersToPoints(4.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        .TabStops.Add Position:=CentimetersToPoints(8.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        .TabStops.Add Position:=CentimetersToPoints(12.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .MatchCase = True
        .Wrap = wdFindStop

        .Text = "(^13)(A\.)"
        .Replacement.Text = "\1^t\2"
        .Execute Replace:=wdReplaceAll

        .Text = "(^13)([BCD]\.)"
        .Replacement.Text = "^t\2"
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True
End Sub

'----------------------------------------------------------------------------------------
' 9. TÁCH ĐÁP ÁN XUỐNG 4 DÒNG DỌC RIÊNG BIỆT (BẢO VỆ "CHỌN A/B/C/D")
'----------------------------------------------------------------------------------------
Sub XuongDongDapAn_4Dong()
    Dim rng As Range
    Set rng = GetTargetRange()

    Application.ScreenUpdating = False
    Call SplitChoicesInRange(rng)
    With rng.ParagraphFormat
        .TabStops.ClearAll
        .LeftIndent = CentimetersToPoints(0.5)
    End With
    Application.ScreenUpdating = True
End Sub

'----------------------------------------------------------------------------------------
' 10. ĐÁNH LẠI SỐ THỨ TỰ CÂU HỎI TUẦN TỰ (Câu 1:, Câu 2:...)
'----------------------------------------------------------------------------------------
Sub DanhLaiThuTuCau()
    Dim dem As Long
    Dim rng As Range
    dem = 1

    Application.ScreenUpdating = False
    ActiveDocument.Range.ListFormat.ConvertNumbersToText

    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Text = "(<[Cc]âu )([0-9]{1,4})([:.])"
        .Replacement.Text = "Câu #@@#\3"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Text = "#@@#"
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            rng.Text = CStr(dem)
            rng.Font.Bold = True
            rng.Collapse wdCollapseEnd
            dem = dem + 1
        Loop
    End With

    Application.ScreenUpdating = True
    MsgBox "Đã đánh lại số thứ tự cho " & (dem - 1) & " câu hỏi.", vbInformation, "Đánh STT Câu"
End Sub

'----------------------------------------------------------------------------------------
' 11. ẨN / HIỆN KHỐI LỜI GIẢI (GIỮ NGUYÊN CẤU TRÚC ĐỀ ĐỂ IN PHÁT HỌC SINH)
'----------------------------------------------------------------------------------------
Sub AnLoiGiai()
    Dim p As Paragraph
    Dim isHiding As Boolean
    isHiding = False

    Application.ScreenUpdating = False
    For Each p In ActiveDocument.Paragraphs
        Dim txt As String
        txt = Trim(p.Range.Text)

        ' Nếu gặp Câu mới hoặc PHẦN mới thì dừng ẩn
        If txt Like "Câu [0-9]*" Or txt Like "Bài [0-9]*" Or txt Like "PHẦN *" Or txt Like "Phần *" Then
            isHiding = False
        ElseIf txt Like "Lời giải*" Or txt Like "Hướng dẫn giải*" Or txt Like "LỜI GIẢI*" Or txt Like "HƯỚNG DẪN GIẢI*" Then
            isHiding = True
        End If

        If isHiding Then
            p.Range.Font.Hidden = True
        End If
    Next p
    Application.ScreenUpdating = True
    MsgBox "Đã ẩn toàn bộ khối Lời giải. (Dùng HienLoiGiai để hiện lại)", vbInformation, "Ẩn Lời Giải"
End Sub

Sub HienLoiGiai()
    Application.ScreenUpdating = False
    ActiveDocument.Content.Font.Hidden = False
    Application.ScreenUpdating = True
    MsgBox "Đã hiện lại toàn bộ Lời giải.", vbInformation, "Hiện Lời Giải"
End Sub

'----------------------------------------------------------------------------------------
' HÀM TRỢ GIÚP NỘI BỘ
'----------------------------------------------------------------------------------------
Private Function GetTargetRange() As Range
    If Selection.Type = wdSelectionIP Or Selection.Range.End - Selection.Range.Start < 3 Then
        Set GetTargetRange = ActiveDocument.Content
    Else
        Set GetTargetRange = Selection.Range
    End If
End Function

Private Sub SplitChoicesInRange(ByRef rng As Range)
    ' Bảo vệ Chọn A/B/C/D trước khi tách
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Text = "Ch" & ChrW(7885) & "n A.": .Replacement.Text = "Ch" & ChrW(7885) & "n A": .Execute Replace:=wdReplaceAll
        .Text = "Ch" & ChrW(7885) & "n B.": .Replacement.Text = "Ch" & ChrW(7885) & "n B": .Execute Replace:=wdReplaceAll
        .Text = "Ch" & ChrW(7885) & "n C.": .Replacement.Text = "Ch" & ChrW(7885) & "n C": .Execute Replace:=wdReplaceAll
        .Text = "Ch" & ChrW(7885) & "n D.": .Replacement.Text = "Ch" & ChrW(7885) & "n D": .Execute Replace:=wdReplaceAll
    End With

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .MatchCase = False
        .Wrap = wdFindStop
        .Text = "([^9, ])([b-z][.,\)])"
        .Replacement.Text = "^p\2"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub ThayTheChuoi(ByVal strTim As String, ByVal strThayThe As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Text = strTim
        .Replacement.Text = strThayThe
        .Execute Replace:=wdReplaceAll
    End With
End Sub
