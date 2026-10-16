Attribute VB_Name = "Mdl_BSP_UI"
'===== Bank Statement Parser (BSP) - user interface
'
' Sheets (created by BSP_Setup_UI):
'   BSP         : settings, options, file list and buttons
'   Hasil       : one row per file / consolidated output (run_result.csv)
'   Verifikasi  : verification findings (verification_detail.csv + consolidation_detail.csv)
'   Daftar Bank : bank formats available in the engine (bank_registry.csv) - drop-down source
Option Explicit

Public Const SHT_UI As String = "BSP"
Public Const SHT_RESULT As String = "Hasil"
Public Const SHT_VERIFY As String = "Verifikasi"
Public Const SHT_BANKS As String = "Daftar Bank"

Public Const TBL_HDR_ROW As Long = 22          'file table header row on the BSP sheet
Public Const COL_NO As Long = 1
Public Const COL_FILE As Long = 2
Public Const COL_BANK As Long = 3
Public Const COL_FORMAT As Long = 4
Public Const COL_PASSWORD As Long = 5
Public Const COL_DET_BANK As Long = 6
Public Const COL_SCORE As Long = 7
Public Const COL_ACC As Long = 8
Public Const COL_NAME As Long = 9
Public Const COL_PERIOD As Long = 10
Public Const COL_STATUS As Long = 11
Public Const COL_MSG As Long = 12

Public Const OPT_PER_FILE As String = "Per file (1 PDF = 1 Excel)"
Public Const OPT_CONSOLIDATED As String = "Gabung per rekening"
Public Const OPT_LOC_CHOSEN As String = "Folder pilihan di atas"
Public Const OPT_LOC_PDF As String = "Sama dengan folder PDF"
Public Const OPT_SUFFIX As String = "Beri akhiran (2) (3) dst"
Public Const OPT_OVERWRITE As String = "Timpa file lama"
Public Const OPT_NAME_ORIGINAL As String = "Sama dengan nama PDF"
Public Const OPT_NAME_STANDARD As String = "Standar (BANK_NoRek_Periode)"

Public Function BSP_UI() As Worksheet
    Set BSP_UI = ThisWorkbook.Worksheets(SHT_UI)
End Function

Private Function Last_File_Row() As Long
    Dim r As Long
    r = BSP_UI.Cells(BSP_UI.Rows.Count, COL_FILE).End(xlUp).Row
    If r < TBL_HDR_ROW Then r = TBL_HDR_ROW
    Last_File_Row = r
End Function

Private Sub Speed(onOff As Boolean)
    Application.ScreenUpdating = Not onOff
    Application.EnableEvents = Not onOff
    Application.DisplayAlerts = Not onOff
End Sub

'===== Browse buttons ========================================================

Private Function Browse_Folder(title As String, start As String) As String
    BSP_Trace "Browse_Folder"
    With Application.FileDialog(4)          'msoFileDialogFolderPicker
        .Title = title
        .AllowMultiSelect = False
        If start <> "" Then
            If Len(Dir(start, vbDirectory)) > 0 Then .InitialFileName = start & "\"
        End If
        If .Show = -1 Then Browse_Folder = .SelectedItems(1)
    End With
End Function

Public Sub BSP_Btn_Browse_R()
    BSP_Trace "BSP_Btn_Browse_R"
    Dim f As String
    f = Browse_Folder("Pilih folder R (berisi Rscript.exe), mis. C:\Program Files\R\R-4.2.2\bin", _
                      BSP_UI.Range("BSP_R_Folder").Value)
    If f <> "" Then BSP_UI.Range("BSP_R_Folder").Value = f
End Sub

Public Sub BSP_Btn_Browse_Engine()
    BSP_Trace "BSP_Btn_Browse_Engine"
    Dim f As String
    f = Browse_Folder("Pilih folder engine R (berisi " & BSP_ENGINE_ENTRY & ")", BSP_UI.Range("BSP_Engine_Folder").Value)
    If f <> "" Then BSP_UI.Range("BSP_Engine_Folder").Value = f
End Sub

Public Sub BSP_Btn_Browse_Output()
    BSP_Trace "BSP_Btn_Browse_Output"
    Dim f As String
    f = Browse_Folder("Pilih folder tujuan output Excel", BSP_UI.Range("BSP_Output_Folder").Value)
    If f <> "" Then
        BSP_UI.Range("BSP_Output_Folder").Value = f
        BSP_UI.Range("BSP_Output_Location").Value = OPT_LOC_CHOSEN
    End If
End Sub

'===== File list =============================================================

Public Sub BSP_Btn_Select_Files()
    BSP_Trace "BSP_Btn_Select_Files"

    Dim files As Variant, i As Long, r As Long, k As Long, exists As Boolean
    Dim startDir As String

    startDir = BSP_UI.Range("BSP_Last_Input_Folder").Value
    If startDir <> "" Then
        If Len(Dir(startDir, vbDirectory)) > 0 Then
            ChDrive Left(startDir, 1)
            ChDir startDir
        End If
    End If

    files = Application.GetOpenFilename(FileFilter:="PDF rekening koran (*.pdf),*.pdf", _
                                        Title:="Pilih satu atau lebih file rekening koran", MultiSelect:=True)
    If Not IsArray(files) Then Exit Sub

    Speed True
    For i = LBound(files) To UBound(files)
        exists = False
        For k = TBL_HDR_ROW + 1 To Last_File_Row()
            If LCase(BSP_UI.Cells(k, COL_FILE).Value) = LCase(files(i)) Then exists = True
        Next k
        If Not exists Then
            r = Last_File_Row() + 1
            BSP_UI.Cells(r, COL_NO).Value = r - TBL_HDR_ROW
            BSP_UI.Cells(r, COL_FILE).Value = files(i)
            BSP_UI.Cells(r, COL_BANK).Value = "AUTO"
        End If
    Next i
    BSP_UI.Range("BSP_Last_Input_Folder").Value = Left(files(LBound(files)), InStrRev(files(LBound(files)), "\") - 1)
    Format_File_Table
    Speed False

End Sub

Public Sub BSP_Btn_Clear_Files()
    BSP_Trace "BSP_Btn_Clear_Files"
    If BSP_Msg("Hapus semua file dari daftar?", vbYesNo + vbQuestion, "BSP") = vbNo Then Exit Sub
    Speed True
    If Last_File_Row() > TBL_HDR_ROW Then
        BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, 1), BSP_UI.Cells(Last_File_Row(), COL_MSG)).Clear
    End If
    Format_File_Table
    Speed False
End Sub

Private Sub Format_File_Table()
    BSP_Trace "Format_File_Table"
    Dim lastR As Long, rng As Range
    lastR = Application.Max(Last_File_Row(), TBL_HDR_ROW + 1)
    Set rng = BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, 1), BSP_UI.Cells(lastR, COL_MSG))
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(200, 200, 200)
    rng.VerticalAlignment = xlTop
    With BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, COL_BANK), BSP_UI.Cells(lastR + 50, COL_BANK)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BSP_Bank_List"
    End With
    With BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, COL_FORMAT), BSP_UI.Cells(lastR + 50, COL_FORMAT)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=BSP_Variant_List"
    End With
    BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, COL_FORMAT), BSP_UI.Cells(lastR + 50, COL_FORMAT)).Interior.Color = RGB(255, 255, 230)
    BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, COL_BANK), BSP_UI.Cells(lastR + 50, COL_BANK)).Interior.Color = RGB(255, 255, 230)
    BSP_UI.Range(BSP_UI.Cells(TBL_HDR_ROW + 1, COL_PASSWORD), BSP_UI.Cells(lastR + 50, COL_PASSWORD)).Interior.Color = RGB(255, 255, 230)
End Sub

'===== Validation of settings ================================================

Private Function Check_Settings(needOutput As Boolean) As Boolean
    BSP_Trace "Check_Settings"

    Dim engine As String, outDir As String

    Check_Settings = False
    If BSP_RScript_Path() = "" Then
        BSP_Msg "Mohon pilih folder R yang berisi Rscript.exe!" & vbNewLine & _
               "(contoh: C:\Program Files\R\R-4.2.2\bin)", vbExclamation, "BSP"
        Exit Function
    End If
    engine = BSP_Clean_Path(BSP_UI.Range("BSP_Engine_Folder").Value)
    If engine = "" Or Len(Dir(engine & "\" & BSP_ENGINE_ENTRY)) = 0 Then
        BSP_Msg "Mohon pilih folder engine R yang berisi " & BSP_ENGINE_ENTRY & "!", vbExclamation, "BSP"
        Exit Function
    End If
    If needOutput And BSP_UI.Range("BSP_Output_Location").Value <> OPT_LOC_PDF Then
        outDir = BSP_Clean_Path(BSP_UI.Range("BSP_Output_Folder").Value)
        If outDir = "" Or Len(Dir(outDir, vbDirectory)) = 0 Then
            BSP_Msg "Mohon pilih folder tujuan output!", vbExclamation, "BSP"
            Exit Function
        End If
    End If
    If Last_File_Row() = TBL_HDR_ROW Then
        BSP_Msg "Belum ada file PDF. Klik 'Pilih File PDF' terlebih dahulu.", vbExclamation, "BSP"
        Exit Function
    End If
    Check_Settings = True

End Function

'===== Job files =============================================================

Private Sub Write_Manifest()
    BSP_Trace "Write_Manifest"
    Dim r As Long, txt As String, bank As String, outDir As String, f As String
    txt = "file_path,bank_id,variant,password,output_dir" & vbLf
    For r = TBL_HDR_ROW + 1 To Last_File_Row()
        f = BSP_UI.Cells(r, COL_FILE).Value
        If f <> "" Then
            bank = BSP_UI.Cells(r, COL_BANK).Value
            If bank = "" Then bank = "AUTO"
            outDir = ""
            If BSP_UI.Range("BSP_Output_Location").Value = OPT_LOC_PDF Then outDir = Left(f, InStrRev(f, "\") - 1)
            txt = txt & BSP_Csv_Field(f) & "," & BSP_Csv_Field(bank) & "," & _
                  BSP_Csv_Field(CStr(BSP_UI.Cells(r, COL_FORMAT).Value)) & "," & _
                  BSP_Csv_Field(CStr(BSP_UI.Cells(r, COL_PASSWORD).Value)) & "," & BSP_Csv_Field(outDir) & vbLf
        End If
    Next r
    BSP_Write_Text BSP_Work_Folder() & "\job_manifest.csv", txt
End Sub

Private Sub Write_Options(verifyOnly As Boolean)
    BSP_Trace "Write_Options"
    Dim txt As String, outDir As String
    outDir = BSP_Clean_Path(BSP_UI.Range("BSP_Output_Folder").Value)
    If outDir = "" Then outDir = BSP_Work_Folder()
    txt = "key,value" & vbLf
    txt = txt & "output_dir," & BSP_Csv_Field(outDir) & vbLf
    txt = txt & "output_mode," & IIf(BSP_UI.Range("BSP_Output_Mode").Value = OPT_CONSOLIDATED, "consolidated", "per_file") & vbLf
    txt = txt & "on_exists," & IIf(BSP_UI.Range("BSP_On_Exists").Value = OPT_OVERWRITE, "overwrite", "suffix") & vbLf
    txt = txt & "tolerance," & Replace(CStr(BSP_UI.Range("BSP_Tolerance").Value), ",", ".") & vbLf   'locale-safe decimal
    txt = txt & "verify_only," & IIf(verifyOnly, "1", "0") & vbLf
    'consolidated outputs always get the standard name (the engine enforces it too)
    txt = txt & "name_style," & IIf(BSP_UI.Range("BSP_Name_Style").Value = OPT_NAME_STANDARD Or _
                                    BSP_UI.Range("BSP_Output_Mode").Value = OPT_CONSOLIDATED, "standard", "original") & vbLf
    BSP_Write_Text BSP_Work_Folder() & "\job_options.csv", txt
End Sub

'===== Buttons: refresh banks, detect, process ===============================

Public Sub BSP_Btn_Refresh_Banks()
    BSP_Trace "BSP_Btn_Refresh_Banks"
    Dim errMsg As String
    If BSP_RScript_Path() = "" Or Len(Dir(BSP_Clean_Path(BSP_UI.Range("BSP_Engine_Folder").Value) & "\" & BSP_ENGINE_ENTRY)) = 0 Then
        BSP_Msg "Mohon pilih folder R dan folder engine terlebih dahulu!", vbExclamation, "BSP"
        Exit Sub
    End If
    If Not BSP_Run_Engine("list_banks", errMsg) Then
        BSP_Msg errMsg, vbCritical, "BSP"
        Exit Sub
    End If
    Load_Bank_Registry
    BSP_Msg "Daftar bank diperbarui: " & (ThisWorkbook.Names("BSP_Bank_List").RefersToRange.Rows.Count - 1) & _
           " bank tersedia.", vbInformation, "BSP"
End Sub

Private Sub Load_Bank_Registry()
    BSP_Trace "Load_Bank_Registry"

    Dim data As Variant, ws As Worksheet, r As Long, n As Long, ids As Object, variants As Object
    Dim cId As Long, cName As Long, cVariant As Long, cDesc As Long, k As Variant, row As Long

    data = BSP_Read_Csv(BSP_Work_Folder() & "\bank_registry.csv")
    If IsEmpty(data) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHT_BANKS)
    ws.Cells.Clear
    cId = BSP_Col(data, "bank_id"): cName = BSP_Col(data, "bank_name")
    cVariant = BSP_Col(data, "variant"): cDesc = BSP_Col(data, "description")

    ws.Range("A1:D1").Value = Array("Kode Bank", "Format", "Nama Bank", "Keterangan")
    ws.Range("F1:G1").Value = Array("Pilihan Bank", "Pilihan Format")
    Set ids = CreateObject("Scripting.Dictionary")
    Set variants = CreateObject("Scripting.Dictionary")
    n = 1
    For r = 2 To UBound(data, 1)
        If data(r, cId) <> "" Then
            n = n + 1
            ws.Cells(n, 1).Value = data(r, cId)
            ws.Cells(n, 2).Value = data(r, cVariant)
            ws.Cells(n, 3).Value = data(r, cName)
            ws.Cells(n, 4).Value = data(r, cDesc)
            ids(data(r, cId)) = 1
            variants(data(r, cVariant)) = 1
        End If
    Next r

    ws.Range("F2").Value = "AUTO"
    row = 3
    For Each k In ids.Keys
        ws.Cells(row, 6).Value = k
        row = row + 1
    Next k
    ThisWorkbook.Names("BSP_Bank_List").RefersTo = "='" & SHT_BANKS & "'!$F$2:$F$" & (row - 1)
    row = 2
    For Each k In variants.Keys
        ws.Cells(row, 7).Value = k
        row = row + 1
    Next k
    ThisWorkbook.Names("BSP_Variant_List").RefersTo = "='" & SHT_BANKS & "'!$G$2:$G$" & Application.Max(2, row - 1)

    ws.Range("A1:G1").Font.Bold = True
    ws.Columns("A:G").AutoFit
    BSP_UI.Range("BSP_Engine_Info").Value = "Engine: " & (ids.Count) & " bank, " & (n - 1) & " format  (diperbarui " & Format(Now, "dd/mm/yyyy hh:mm") & ")"

End Sub

Public Sub BSP_Btn_Detect()
    BSP_Trace "BSP_Btn_Detect"

    Dim errMsg As String, data As Variant, r As Long, k As Long, f As String
    Dim cFile As Long, cBank As Long, cScore As Long, cStatus As Long, cMsg As Long
    Dim cAcc As Long, cName As Long, cPs As Long, cPe As Long, cVariant As Long

    If Not Check_Settings(False) Then Exit Sub
    Write_Manifest
    If Not BSP_Run_Engine("detect", errMsg) Then
        BSP_Msg errMsg, vbCritical, "BSP"
        Exit Sub
    End If
    Load_Bank_Registry

    data = BSP_Read_Csv(BSP_Work_Folder() & "\detect_result.csv")
    If IsEmpty(data) Then Exit Sub
    cFile = BSP_Col(data, "file_path"): cBank = BSP_Col(data, "bank_id"): cVariant = BSP_Col(data, "variant")
    cScore = BSP_Col(data, "score"): cStatus = BSP_Col(data, "status"): cMsg = BSP_Col(data, "message")
    cAcc = BSP_Col(data, "account_no"): cName = BSP_Col(data, "account_name")
    cPs = BSP_Col(data, "period_start"): cPe = BSP_Col(data, "period_end")

    Speed True
    For k = TBL_HDR_ROW + 1 To Last_File_Row()
        f = BSP_UI.Cells(k, COL_FILE).Value
        For r = 2 To UBound(data, 1)
            If LCase(data(r, cFile)) = LCase(f) Then
                BSP_UI.Cells(k, COL_DET_BANK).Value = data(r, cBank) & IIf(data(r, cVariant) <> "", " - " & data(r, cVariant), "")
                BSP_UI.Cells(k, COL_SCORE).Value = data(r, cScore)
                BSP_UI.Cells(k, COL_ACC).NumberFormat = "@"
                BSP_UI.Cells(k, COL_ACC).Value = data(r, cAcc)
                BSP_UI.Cells(k, COL_NAME).Value = data(r, cName)
                BSP_UI.Cells(k, COL_PERIOD).Value = IIf(data(r, cPs) <> "", data(r, cPs) & " - " & data(r, cPe), "")
                BSP_UI.Cells(k, COL_STATUS).Value = Status_Label(CStr(data(r, cStatus)))
                BSP_UI.Cells(k, COL_MSG).Value = data(r, cMsg)
                Color_Status BSP_UI.Cells(k, COL_STATUS), CStr(data(r, cStatus))
                Exit For
            End If
        Next r
    Next k
    Format_File_Table
    Speed False
    BSP_UI.Activate

End Sub

Private Function Status_Label(s As String) As String
    Select Case s
        Case "AUTO": Status_Label = "Terdeteksi otomatis"
        Case "MANUAL": Status_Label = "Dipilih manual"
        Case "AMBIGUOUS": Status_Label = "Ragu - mohon cek"
        Case "UNKNOWN": Status_Label = "Tidak dikenali"
        Case "PASSWORD": Status_Label = "Perlu password"
        Case Else: Status_Label = s
    End Select
End Function

Private Sub Color_Status(c As Range, s As String)
    Select Case s
        Case "OK", "AUTO", "MANUAL": c.Interior.Color = RGB(198, 239, 206)
        Case "WARN", "AMBIGUOUS": c.Interior.Color = RGB(255, 235, 156)
        Case Else: c.Interior.Color = RGB(255, 199, 206)
    End Select
End Sub

Public Sub BSP_Btn_Process()
    BSP_Trace "BSP_Btn_Process"
    Run_Process False
End Sub

Public Sub BSP_Btn_Verify_Only()
    BSP_Trace "BSP_Btn_Verify_Only"
    Run_Process True
End Sub

Private Sub Run_Process(verifyOnly As Boolean)
    BSP_Trace "Run_Process"

    Dim errMsg As String, data As Variant, nOk As Long, nWarn As Long, nErr As Long, r As Long
    Dim cStatus As Long, msg As String, outDir As String

    If Not Check_Settings(Not verifyOnly) Then Exit Sub
    If BSP_Msg(IIf(verifyOnly, "Verifikasi ", "Proses ") & (Last_File_Row() - TBL_HDR_ROW) & " file PDF sekarang?", _
              vbYesNo + vbQuestion, "BSP") = vbNo Then Exit Sub

    Write_Manifest
    Write_Options verifyOnly
    If Not BSP_Run_Engine("parse", errMsg) Then
        BSP_Msg errMsg, vbCritical, "BSP"
        Exit Sub
    End If

    Speed True
    Load_Results
    Load_Verification
    Speed False

    data = BSP_Read_Csv(BSP_Work_Folder() & "\run_result.csv")
    If Not IsEmpty(data) Then
        cStatus = BSP_Col(data, "status")
        For r = 2 To UBound(data, 1)
            Select Case data(r, cStatus)
                Case "OK": nOk = nOk + 1
                Case "WARN": nWarn = nWarn + 1
                Case Else: nErr = nErr + 1
            End Select
        Next r
    End If

    ThisWorkbook.Worksheets(SHT_RESULT).Activate
    msg = "Selesai!" & vbNewLine & vbNewLine & _
          "OK      : " & nOk & vbNewLine & _
          "WARN    : " & nWarn & " (periksa sheet Verifikasi)" & vbNewLine & _
          "ERROR   : " & nErr & IIf(nErr > 0, " (periksa sheet Hasil & Verifikasi)", "")
    BSP_Msg msg, IIf(nErr > 0, vbExclamation, vbInformation), "BSP"

    If Not verifyOnly And BSP_UI.Range("BSP_Open_Folder").Value = "Ya" Then
        If BSP_UI.Range("BSP_Output_Location").Value = OPT_LOC_PDF Then
            outDir = BSP_UI.Range("BSP_Last_Input_Folder").Value
        Else
            outDir = BSP_Clean_Path(BSP_UI.Range("BSP_Output_Folder").Value)
        End If
        If outDir <> "" And Not BSP_SILENT Then Shell "explorer.exe """ & outDir & """", vbNormalFocus
    End If

End Sub

'===== Result sheets =========================================================

Private Sub Load_Results()
    BSP_Trace "Load_Results"

    Dim data As Variant, ws As Worksheet, r As Long, c As Long, i As Long
    Dim cols As Variant, heads As Variant, src As Long

    cols = Array("file_name", "bank_id", "variant", "account_no", "account_name", "currency", "period_start", _
                 "period_end", "rows", "opening", "opening_source", "total_debit", "total_credit", "closing", _
                 "n_error", "n_warn", "status", "output_file", "message", "source_files")
    heads = Array("File PDF", "Bank", "Format", "No. Rekening", "Nama Rekening", "Mata Uang", "Periode Awal", _
                  "Periode Akhir", "Jumlah Transaksi", "Saldo Awal", "Sumber Saldo Awal", "Total Debit", _
                  "Total Kredit", "Saldo Akhir", "Error", "Warning", "Status", "File Output", "Pesan", "File Sumber (gabungan)")

    Set ws = ThisWorkbook.Worksheets(SHT_RESULT)
    ws.Cells.Clear
    ws.Range("A1").Value = "Hasil proses - " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    For i = 0 To UBound(heads)
        ws.Cells(3, i + 1).Value = heads(i)
    Next i

    data = BSP_Read_Csv(BSP_Work_Folder() & "\run_result.csv")
    If IsEmpty(data) Then Exit Sub
    For r = 2 To UBound(data, 1)
        For i = 0 To UBound(cols)
            src = BSP_Col(data, CStr(cols(i)))
            If src > 0 Then
                Select Case cols(i)
                    Case "account_no", "period_start", "period_end"
                        ws.Cells(r + 2, i + 1).NumberFormat = "@"
                        ws.Cells(r + 2, i + 1).Value = data(r, src)
                    Case "rows", "opening", "total_debit", "total_credit", "closing", "n_error", "n_warn"
                        If data(r, src) <> "" Then ws.Cells(r + 2, i + 1).Value = Val(data(r, src))
                    Case "opening_source"
                        ws.Cells(r + 2, i + 1).Value = IIf(data(r, src) = "stated", "Dari PDF", IIf(data(r, src) = "derived", "Dihitung", data(r, src)))
                    Case "output_file"
                        If data(r, src) <> "" Then
                            ws.Hyperlinks.Add Anchor:=ws.Cells(r + 2, i + 1), Address:=CStr(data(r, src)), _
                                              TextToDisplay:=Mid(data(r, src), InStrRev(data(r, src), "\") + 1)
                        End If
                    Case Else
                        ws.Cells(r + 2, i + 1).Value = data(r, src)
                End Select
            End If
        Next i
        Color_Status ws.Cells(r + 2, 17), CStr(data(r, BSP_Col(data, "status")))
    Next r

    With ws.Range(ws.Cells(3, 1), ws.Cells(UBound(data, 1) + 2, UBound(heads) + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(heads) + 1)).Font.Bold = True
    ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(heads) + 1)).Interior.Color = RGB(31, 78, 121)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(heads) + 1)).Font.Color = RGB(255, 255, 255)
    ws.Range("J:N").NumberFormat = "#,##0.00"
    ws.Columns("A:T").AutoFit
    If ws.Columns(1).ColumnWidth > 60 Then ws.Columns(1).ColumnWidth = 60
    If ws.Columns(20).ColumnWidth > 80 Then ws.Columns(20).ColumnWidth = 80
    ws.Cells(UBound(data, 1) + 4, 1).Value = "Log engine: " & BSP_Work_Folder() & "\engine_log.txt"

End Sub

Private Sub Load_Verification()
    BSP_Trace "Load_Verification"

    Dim data As Variant, cons As Variant, ws As Worksheet, r As Long, i As Long, row As Long
    Dim cols As Variant, heads As Variant, src As Long, sev As String

    cols = Array("file_name", "output_file", "excel_row", "date", "description", "debit", "credit", "balance", _
                 "check", "severity", "message")
    heads = Array("File PDF", "File Output", "Baris Excel", "Tanggal", "Deskripsi", "Debit", "Kredit", "Saldo", _
                  "Pemeriksaan", "Tingkat", "Keterangan")

    Set ws = ThisWorkbook.Worksheets(SHT_VERIFY)
    ws.Cells.Clear
    ws.Range("A1").Value = "Hasil verifikasi - ERROR: angka tidak cocok | WARN: perlu dicek | INFO: catatan"
    ws.Range("A1").Font.Bold = True
    For i = 0 To UBound(heads)
        ws.Cells(3, i + 1).Value = heads(i)
    Next i
    row = 3

    data = BSP_Read_Csv(BSP_Work_Folder() & "\verification_detail.csv")
    If Not IsEmpty(data) Then
        For r = 2 To UBound(data, 1)
            row = row + 1
            For i = 0 To UBound(cols)
                src = BSP_Col(data, CStr(cols(i)))
                If src > 0 Then
                    Select Case cols(i)
                        Case "excel_row", "debit", "credit", "balance"
                            If data(r, src) <> "" Then ws.Cells(row, i + 1).Value = Val(data(r, src))
                        Case "date"
                            ws.Cells(row, i + 1).NumberFormat = "@"
                            ws.Cells(row, i + 1).Value = data(r, src)
                        Case Else
                            ws.Cells(row, i + 1).Value = data(r, src)
                    End Select
                End If
            Next i
            sev = data(r, BSP_Col(data, "severity"))
            Color_Status ws.Cells(row, 10), IIf(sev = "INFO", "OK", sev)
        Next r
    End If

    cons = BSP_Read_Csv(BSP_Work_Folder() & "\consolidation_detail.csv")
    If Not IsEmpty(cons) Then
        For r = 2 To UBound(cons, 1)
            row = row + 1
            ws.Cells(row, 1).Value = cons(r, BSP_Col(cons, "file_name"))
            ws.Cells(row, 2).Value = cons(r, BSP_Col(cons, "output_file"))
            ws.Cells(row, 9).Value = "antar-file: " & cons(r, BSP_Col(cons, "check"))
            ws.Cells(row, 10).Value = cons(r, BSP_Col(cons, "severity"))
            ws.Cells(row, 11).Value = cons(r, BSP_Col(cons, "message"))
            Color_Status ws.Cells(row, 10), CStr(cons(r, BSP_Col(cons, "severity")))
        Next r
    End If

    If row = 3 Then
        ws.Cells(4, 1).Value = "Tidak ada temuan - semua statement konsisten."
    Else
        With ws.Range(ws.Cells(3, 1), ws.Cells(row, UBound(heads) + 1))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
        End With
        ws.Range(ws.Cells(3, 1), ws.Cells(row, UBound(heads) + 1)).AutoFilter
    End If
    ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(heads) + 1)).Font.Bold = True
    ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(heads) + 1)).Interior.Color = RGB(31, 78, 121)
    ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(heads) + 1)).Font.Color = RGB(255, 255, 255)
    ws.Range("F:H").NumberFormat = "#,##0.00"
    ws.Columns("A:K").AutoFit
    If ws.Columns(1).ColumnWidth > 50 Then ws.Columns(1).ColumnWidth = 50
    If ws.Columns(5).ColumnWidth > 50 Then ws.Columns(5).ColumnWidth = 50

End Sub

'===== Automated self-test (used by excel\build_xlsm.ps1 -Test) ===============
'pdfList: full paths separated by "|". Returns the collected messages + result statuses.
Public Function BSP_Selftest(pdfList As String, outDir As String, outputMode As String, _
                             Optional ByVal nameStyle As String = "original") As String
    BSP_Trace "BSP_Selftest"

    Dim files() As String, i As Long, r As Long, ws As Worksheet, res As String

    BSP_SILENT = True
    BSP_LAST_MSG = ""
    On Error GoTo Fail

    BSP_Btn_Clear_Files
    files = Split(pdfList, "|")
    For i = 0 To UBound(files)
        r = TBL_HDR_ROW + 1 + i
        BSP_UI.Cells(r, COL_NO).Value = i + 1
        BSP_UI.Cells(r, COL_FILE).Value = files(i)
        BSP_UI.Cells(r, COL_BANK).Value = "AUTO"
    Next i
    BSP_UI.Range("BSP_Output_Folder").Value = outDir
    BSP_UI.Range("BSP_Output_Location").Value = OPT_LOC_CHOSEN
    BSP_UI.Range("BSP_Output_Mode").Value = IIf(outputMode = "consolidated", OPT_CONSOLIDATED, OPT_PER_FILE)
    BSP_UI.Range("BSP_On_Exists").Value = OPT_OVERWRITE
    BSP_UI.Range("BSP_Name_Style").Value = IIf(nameStyle = "standard", OPT_NAME_STANDARD, OPT_NAME_ORIGINAL)

    BSP_Btn_Refresh_Banks
    BSP_Btn_Detect
    For r = TBL_HDR_ROW + 1 To Last_File_Row()
        res = res & "DETECT|" & BSP_UI.Cells(r, COL_DET_BANK).Value & "|" & BSP_UI.Cells(r, COL_STATUS).Value & vbLf
    Next r
    Run_Process False

    Set ws = ThisWorkbook.Worksheets(SHT_RESULT)
    For r = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 17).Value <> "" Then
            res = res & "RESULT|" & ws.Cells(r, 1).Value & "|" & ws.Cells(r, 2).Value & "|" & ws.Cells(r, 9).Value & _
                  "|" & ws.Cells(r, 17).Value & "|" & ws.Cells(r, 18).Value & vbLf
        End If
    Next r
    'leave the workbook clean
    BSP_Btn_Clear_Files
    BSP_UI.Range("BSP_Output_Folder").Value = ""
    BSP_UI.Range("BSP_On_Exists").Value = OPT_SUFFIX
    BSP_UI.Range("BSP_Output_Mode").Value = OPT_PER_FILE
    BSP_UI.Range("BSP_Name_Style").Value = OPT_NAME_ORIGINAL
    ThisWorkbook.Worksheets(SHT_RESULT).Cells.Clear
    ThisWorkbook.Worksheets(SHT_VERIFY).Cells.Clear
    BSP_UI.Activate
    BSP_SILENT = False
    BSP_Selftest = BSP_LAST_MSG & res
    Exit Function

Fail:
    BSP_SILENT = False
    BSP_Selftest = BSP_LAST_MSG & res & "VBA ERROR " & Err.Number & ": " & Err.Description
End Function
