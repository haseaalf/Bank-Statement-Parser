Attribute VB_Name = "Mdl_BSP_Setup"
'===== Bank Statement Parser (BSP) - builds the workbook interface
'
' Run BSP_Setup_UI once in a workbook that contains the three Mdl_BSP_* modules
' (excel\build_xlsm.ps1 does this automatically). Running it again rebuilds the
' sheets but keeps the chosen folders.
Option Explicit

Private Sub Add_Name(nm As String, ws As Worksheet, addr As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & ws.Name & "'!" & ws.Range(addr).Address
End Sub

Private Function Get_Sheet(nm As String, afterSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        If afterSheet Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        End If
        ws.Name = nm
    End If
    Set Get_Sheet = ws
End Function

Private Sub Add_Button(ws As Worksheet, cellAddr As String, w As Double, caption As String, macro As String)
    Dim b As Object
    With ws.Range(cellAddr)
        Set b = ws.Buttons.Add(.Left + 2, .Top + 2, w, .Height * 1.6)
    End With
    b.Caption = caption
    b.OnAction = macro
    b.Font.Size = 10
End Sub

Private Sub Put_Label(ws As Worksheet, addr As String, txt As String, Optional bold As Boolean = False)
    ws.Range(addr).Value = txt
    ws.Range(addr).Font.Bold = bold
End Sub

Private Sub Input_Cell(ws As Worksheet, addr As String)
    With ws.Range(addr)
        .Interior.Color = RGB(255, 255, 230)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
    End With
End Sub

Private Sub List_Cell(ws As Worksheet, addr As String, items As String, defaultValue As String)
    Input_Cell ws, addr
    With ws.Range(addr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=items
    End With
    If ws.Range(addr).Value = "" Then ws.Range(addr).Value = defaultValue
End Sub

Public Sub BSP_Setup_UI()

    Dim ws As Worksheet, wr As Worksheet, wv As Worksheet, wb As Worksheet, sh As Object
    Dim keepR As String, keepEngine As String, keepOut As String, i As Long, heads As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ws = Get_Sheet(SHT_UI, Nothing)
    keepR = ws.Range("C6").Value
    keepEngine = ws.Range("C7").Value
    keepOut = ws.Range("C11").Value
    For Each sh In ws.Shapes
        sh.Delete
    Next sh
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    Set wr = Get_Sheet(SHT_RESULT, ws)
    Set wv = Get_Sheet(SHT_VERIFY, wr)
    Set wb = Get_Sheet(SHT_BANKS, wv)

    'remove the default empty sheets
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> SHT_UI And sh.Name <> SHT_RESULT And sh.Name <> SHT_VERIFY And sh.Name <> SHT_BANKS Then
            If Application.WorksheetFunction.CountA(sh.Cells) = 0 Then sh.Delete
        End If
    Next sh

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 30
    ws.Columns("G").ColumnWidth = 7
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 30
    ws.Columns("J").ColumnWidth = 24
    ws.Columns("K").ColumnWidth = 20
    ws.Columns("L").ColumnWidth = 50

    With ws.Range("A1")
        .Value = "Bank Statement Parser (BSP)"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    Put_Label ws, "A2", "Mengubah PDF rekening koran berbagai bank menjadi Excel format standar (Tanggal, Deskripsi, Debit, Kredit, Saldo) dan memverifikasi saldonya."

    '--- 1. Engine R
    Put_Label ws, "A4", "1. Pengaturan engine R", True
    Put_Label ws, "B6", "Folder R (berisi Rscript.exe)"
    Put_Label ws, "B7", "Folder engine R (berisi " & BSP_ENGINE_ENTRY & ")"
    ws.Range("C6:F6").Merge
    ws.Range("C7:F7").Merge
    Input_Cell ws, "C6:F6"
    Input_Cell ws, "C7:F7"
    Add_Name "BSP_R_Folder", ws, "C6"
    Add_Name "BSP_Engine_Folder", ws, "C7"
    If keepR = "" Then keepR = BSP_Find_R_Folder()
    If keepEngine = "" Then
        If Len(Dir(ThisWorkbook.Path & "\..\BSP_modular_engine\" & BSP_ENGINE_ENTRY)) > 0 Then
            keepEngine = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\")) & "BSP_modular_engine"
        ElseIf Len(Dir(ThisWorkbook.Path & "\BSP_modular_engine\" & BSP_ENGINE_ENTRY)) > 0 Then
            keepEngine = ThisWorkbook.Path & "\BSP_modular_engine"
        End If
    End If
    ws.Range("C6").Value = keepR
    ws.Range("C7").Value = keepEngine
    Add_Button ws, "G6", 60, "Browse...", "BSP_Btn_Browse_R"
    Add_Button ws, "G7", 60, "Browse...", "BSP_Btn_Browse_Engine"
    Add_Name "BSP_Engine_Info", ws, "C8"
    ws.Range("C8").Font.Color = RGB(120, 120, 120)

    '--- 2. Output
    Put_Label ws, "A10", "2. Output", True
    Put_Label ws, "B11", "Folder tujuan output"
    ws.Range("C11:F11").Merge
    Input_Cell ws, "C11:F11"
    ws.Range("C11").Value = keepOut
    Add_Name "BSP_Output_Folder", ws, "C11"
    Add_Button ws, "G11", 60, "Browse...", "BSP_Btn_Browse_Output"
    Put_Label ws, "B12", "Lokasi penyimpanan"
    List_Cell ws, "C12", OPT_LOC_CHOSEN & "," & OPT_LOC_PDF, OPT_LOC_CHOSEN
    Add_Name "BSP_Output_Location", ws, "C12"
    Put_Label ws, "B13", "Mode output"
    List_Cell ws, "C13", OPT_PER_FILE & "," & OPT_CONSOLIDATED, OPT_PER_FILE
    Add_Name "BSP_Output_Mode", ws, "C13"
    Put_Label ws, "B14", "Jika file output sudah ada"
    List_Cell ws, "C14", OPT_SUFFIX & "," & OPT_OVERWRITE, OPT_SUFFIX
    Add_Name "BSP_On_Exists", ws, "C14"
    Put_Label ws, "B15", "Toleransi selisih saldo"
    Input_Cell ws, "C15"
    If ws.Range("C15").Value = "" Then ws.Range("C15").Value = 0.01
    Add_Name "BSP_Tolerance", ws, "C15"
    Put_Label ws, "B16", "Buka folder output setelah selesai"
    List_Cell ws, "C16", "Ya,Tidak", "Ya"
    Add_Name "BSP_Open_Folder", ws, "C16"
    Put_Label ws, "B17", "Nama file output"
    List_Cell ws, "C17", OPT_NAME_ORIGINAL & "," & OPT_NAME_STANDARD, OPT_NAME_ORIGINAL
    Add_Name "BSP_Name_Style", ws, "C17"
    Put_Label ws, "D17", "Mode gabung selalu memakai nama standar (BANK_NoRek_Periode)"
    ws.Range("D17").Font.Color = RGB(120, 120, 120)
    Put_Label ws, "B18", "Buat rekap mutasi (semua file)"
    List_Cell ws, "C18", "Ya,Tidak", "Tidak"
    Add_Name "BSP_Recap", ws, "C18"
    Put_Label ws, "D18", "1 file <tanggal>_<jam>_rekap_mutasi.xlsx: sheet REKAP + 1 sheet per rekening, dengan cek saldo bulanan"
    ws.Range("D18").Font.Color = RGB(120, 120, 120)
    Add_Name "BSP_Last_Input_Folder", ws, "N1"
    ws.Range("N1").Font.Color = RGB(255, 255, 255)

    '--- 3. Files + actions
    Put_Label ws, "A19", "3. File rekening koran", True
    Put_Label ws, "B20", "Kolom Bank: AUTO = deteksi otomatis, atau pilih bank secara manual. Isi Password bila PDF terkunci."
    ws.Range("B20").Font.Color = RGB(120, 120, 120)
    Add_Button ws, "B21", 110, "Pilih File PDF...", "BSP_Btn_Select_Files"
    Add_Button ws, "C21", 100, "Deteksi Bank", "BSP_Btn_Detect"
    Add_Button ws, "D21", 120, "PROSES", "BSP_Btn_Process"
    Add_Button ws, "F21", 110, "Verifikasi saja", "BSP_Btn_Verify_Only"
    Add_Button ws, "I21", 110, "Hapus daftar", "BSP_Btn_Clear_Files"
    Add_Button ws, "J21", 130, "Perbarui daftar bank", "BSP_Btn_Refresh_Banks"
    ws.Rows(21).RowHeight = 22

    heads = Array("No", "File PDF", "Bank", "Format (opsional)", "Password", "Bank terdeteksi", "Skor", _
                  "No. Rekening", "Nama Rekening", "Periode", "Status", "Pesan")
    For i = 0 To UBound(heads)
        ws.Cells(TBL_HDR_ROW, i + 1).Value = heads(i)
    Next i
    With ws.Range(ws.Cells(TBL_HDR_ROW, 1), ws.Cells(TBL_HDR_ROW, COL_MSG))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(TBL_HDR_ROW).RowHeight = 20

    '--- bank list defaults (refreshed from the engine by "Perbarui daftar bank")
    Add_Name "BSP_Bank_List", wb, "F2"
    Add_Name "BSP_Variant_List", wb, "G2"
    wb.Range("F1").Value = "Pilihan Bank"
    wb.Range("F2").Value = "AUTO"
    wb.Range("G1").Value = "Pilihan Format"

    'window settings can fail when Excel runs hidden (build script) - not essential
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A" & TBL_HDR_ROW + 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
