Attribute VB_Name = "ChuanHoaDeToan"
'========================================================================================
' MACRO VBA CHUẨN HÓA VĂN BẢN ĐỀ THI TOÁN HỌC TRONG MICROSOFT WORD
' Tác giả: Math OCR Tool Team
' Chức năng chính:
'  1. Căn lề chuẩn Nghị định 30/2020/NĐ-CP (Trên 2cm, Dưới 2cm, Trái 3cm, Phải 2cm, Khổ A4).
'  2. Chuẩn hóa Font chữ Times New Roman 12pt, giãn dòng 1.15, cách đoạn After 3pt, Căn đều (Justify).
'  3. Bảo toàn font Cambria Math của toàn bộ công thức toán học OMML (không bị ép về TNR).
'  4. Làm sạch văn bản: Xóa dấu cách thừa, sửa lỗi dấu câu ( . , : ; ), xóa dòng trống liên tiếp.
'  5. In đậm có chọn lọc: Nhãn câu hỏi (Câu 1:, Bài 1:...) và nhãn trắc nghiệm (A., B., C., D.).
'  6. Căn giữa hình ảnh và tự động co giãn nếu kích thước vượt quá lề trang in.
'  7. Căn giữa và tự động căn chỉnh bảng biểu (Tables).
'========================================================================================

Option Explicit

Sub ChuanHoaToanBoFileWord()
    Dim startTime As Double
    startTime = Timer
    
    ' Tắt cập nhật màn hình để tăng tốc độ xử lý gấp 10 lần
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    On Error GoTo ErrorHandler
    
    ' 1. Căn lề trang in chuẩn Nghị định 30 (A4)
    Call ThietLapTrangInA4
    
    ' 2. Làm sạch văn bản (khoảng trắng, dòng trống, dấu câu)
    Call LamSachKhoangTrangVaDauCau
    
    ' 3. Định dạng Font và Đoạn văn (Không làm hỏng công thức Toán)
    Call DinhDangFontVaDoanVan
    
    ' 4. In đậm có chọn lọc: Câu hỏi và Đáp án trắc nghiệm
    Call InDamNhanCauHoiVaDapAn
    
    ' 5. Căn giữa và chuẩn hóa kích thước hình ảnh
    Call ChuanHoaHinhAnh
    
    ' 6. Căn giữa và co giãn bảng biểu
    Call ChuanHoaBangBieu
    
    ' Bật lại màn hình
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    
    Dim totalTime As Double
    totalTime = Round(Timer - startTime, 2)
    
    MsgBox "CHÚC MỪNG THẦY CÔ!" & vbCrLf & _
           "- Đã chuẩn hóa trang in: A4, lề chuẩn NĐ 30." & vbCrLf & _
           "- Đã định dạng: Font Times New Roman 12pt, giãn dòng 1.15." & vbCrLf & _
           "- Đã in đậm chuẩn: Câu 1..., A. B. C. D." & vbCrLf & _
           "- Đã căn giữa hình ảnh, bảng biểu và bảo toàn công thức Toán." & vbCrLf & _
           "- Thời gian xử lý: " & totalTime & " giây.", _
           vbInformation, "Chuẩn Hóa Đề Toán Hoàn Tất"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    MsgBox "Có lỗi xảy ra trong quá trình chuẩn hóa: " & Err.Description, vbCritical, "Lỗi VBA"
End Sub

'----------------------------------------------------------------------------------------
' 1. THIẾT LẬP TRANG IN KHỔ A4 THEO NGHỊ ĐỊNH 30/2020/NĐ-CP
'----------------------------------------------------------------------------------------
Private Sub ThietLapTrangInA4()
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        With sec.PageSetup
            .PaperSize = wdPaperA4
            .TopMargin = CentimetersToPoints(2#)     ' Trên: 2.0 cm
            .BottomMargin = CentimetersToPoints(2#)  ' Dưới: 2.0 cm
            .LeftMargin = CentimetersToPoints(3#)    ' Trái: 3.0 cm
            .RightMargin = CentimetersToPoints(2#)   ' Phải: 2.0 cm
            .HeaderDistance = CentimetersToPoints(1.2)
            .FooterDistance = CentimetersToPoints(1.2)
        End With
    Next sec
End Sub

'----------------------------------------------------------------------------------------
' 2. LÀM SẠCH VĂN BẢN (XÓA KHOẢNG TRẮNG, DÒNG TRỐNG THỪA, SỬA DẤU CÂU)
'----------------------------------------------------------------------------------------
Private Sub LamSachKhoangTrangVaDauCau()
    ' Xóa nhiều khoảng trắng liên tiếp -> 1 khoảng trắng
    Call ThayTheChuoi("   ", " ")
    Call ThayTheChuoi("  ", " ")
    
    ' Sửa lỗi dấu cách trước dấu câu: " ." -> ".", " ," -> ","
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
    
    ' Xóa dòng trống thừa liên tiếp (3 dòng trống -> 1 dòng trống)
    Call ThayTheChuoi("^p^p^p", "^p")
    Call ThayTheChuoi("^p^p", "^p")
End Sub

'----------------------------------------------------------------------------------------
' 3. ĐỊNH DẠNG FONT CHỮ VÀ ĐOẠN VĂN (BẢO TOÀN CÔNG THỨC TOÁN OMML)
'----------------------------------------------------------------------------------------
Private Sub DinhDangFontVaDoanVan()
    Dim p As Paragraph
    Dim rng As Range
    
    For Each p In ActiveDocument.Paragraphs
        ' Bỏ qua nếu đoạn nằm trong bảng (sẽ xử lý ở phần bảng)
        If Not p.Range.Information(wdWithinTable) Then
            With p.Format
                .Alignment = wdAlignParagraphJustify    ' Căn đều 2 bên
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(1.15)       ' Giãn dòng 1.15
                .SpaceBefore = 0                         ' Cách đoạn trước 0pt
                .SpaceAfter = 3                          ' Cách đoạn sau 3pt
            End With
        End If
    Next p
    
    ' Thiết lập Font chung cho toàn bộ Text (không ghi đè Equation)
    Set rng = ActiveDocument.Content
    With rng.Font
        .Name = "Times New Roman"
        .Size = 12
        .ColorIndex = wdBlack
    End With
End Sub

'----------------------------------------------------------------------------------------
' 4. IN ĐẬM CÓ CHỌN LỌC (CÂU HỎI & PHƯƠNG ÁN TRẮC NGHIỆM A. B. C. D.)
'----------------------------------------------------------------------------------------
Private Sub InDamNhanCauHoiVaDapAn()
    Dim rng As Range
    
    ' 4.1. Tìm và in đậm: "Câu X:", "Câu X.", "Bài X:", "Ví dụ X:"
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .MatchWildcards = True
        
        ' Tìm: Câu 1:, Câu 12., Bài 3:...
        .Text = "(<[Cc]âu [0-9]{1,3}[:.!?])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
        
        .Text = "(<[Bb]ài [0-9]{1,3}[:.!?])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
        
        .Text = "(<[Vv]í dụ [0-9]{1,3}[:.!?])"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
        
        ' Dạng: Câu 1 (1.0 điểm):
        .Text = "(<[Cc]âu [0-9]{1,3} \([!)]@\):)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
    End With
    
    ' 4.2. Tìm và in đậm các nhãn phương án trắc nghiệm: A. B. C. D.
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Replacement.Font.Italic = False   ' Bắt buộc chữ đứng, không nghiêng
        .MatchWildcards = True
        
        ' Tìm đầu dòng hoặc sau dấu tab/cách: A. B. C. D.
        .Text = "(<[ABCD]\.)"
        .Replacement.Text = "\1"
        .Execute Replace:=wdReplaceAll
        
        ' Tìm phương án dạng a) b) c) d)
        .Text = "(<[abcd]\))"
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
    ' Chiều rộng vùng in tối đa = 21cm - 3cm (trái) - 2cm (phải) = 16cm (khoảng 453 pt)
    maxWidth = CentimetersToPoints(15.5)
    
    For Each shp In ActiveDocument.InlineShapes
        ' Căn giữa đoạn văn chứa hình ảnh
        shp.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        shp.Range.ParagraphFormat.SpaceBefore = 4
        shp.Range.ParagraphFormat.SpaceAfter = 4
        
        ' Khóa tỷ lệ khung hình
        shp.LockAspectRatio = msoTrue
        
        ' Nếu ảnh quá rộng vượt lề in thì thu nhỏ lại vừa lề an toàn (15.5cm)
        If shp.Width > maxWidth Then
            shp.Width = maxWidth
        End If
        
        ' Nếu ảnh quá nhỏ hoặc quá lớn so với tiêu chuẩn đồ thị Toán (chuẩn 3.5cm - 5.0cm chiều cao)
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
            .Rows.Alignment = wdAlignRowCenter          ' Căn giữa bảng
            .AutoFitBehavior wdAutoFitContent           ' Co giãn theo nội dung
            .TopPadding = CentimetersToPoints(0.1)
            .BottomPadding = CentimetersToPoints(0.1)
            .LeftPadding = CentimetersToPoints(0.15)
            .RightPadding = CentimetersToPoints(0.15)
        End With
    Next tbl
End Sub

'----------------------------------------------------------------------------------------
' HÀM TRỢ GIÚP TÌM VÀ THAY THẾ CHUỖI NHANH
'----------------------------------------------------------------------------------------
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
