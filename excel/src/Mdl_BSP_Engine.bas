Attribute VB_Name = "Mdl_BSP_Engine"
'===== Bank Statement Parser (BSP) - communication with the R engine
'
' Same pattern as BSA: VBA writes the job as files into a work folder, runs
'   "<R folder>\Rscript.exe" "<engine>\BSP_engine.R" "<work folder>" "<engine>" "<mode>"
' synchronously, then reads the result CSV files back.
Option Explicit

Public Const BSP_ENGINE_ENTRY As String = "BSP_engine.R"

'Silent mode (automated tests): messages are collected instead of shown, questions answer Yes
Public BSP_SILENT As Boolean
Public BSP_LAST_MSG As String

'Trace for automated tests (silent mode only): %TEMP%\BSP_work\vba_trace.txt
Public Sub BSP_Trace(txt As String)
    Dim fn As Integer
    If Not BSP_SILENT Then Exit Sub
    fn = FreeFile
    Open BSP_Work_Folder() & "\vba_trace.txt" For Append As #fn
    Print #fn, Format(Now, "hh:nn:ss") & " " & txt
    Close #fn
End Sub

Public Function BSP_Msg(ByVal prompt As String, Optional ByVal buttons As Long = vbOKOnly, _
                        Optional ByVal title As String = "BSP") As VbMsgBoxResult
    If BSP_SILENT Then
        BSP_LAST_MSG = BSP_LAST_MSG & prompt & vbLf
        BSP_Msg = vbYes
    Else
        BSP_Msg = MsgBox(prompt, buttons, title)
    End If
End Function

'Work folder for job/result files (local temp, avoids OneDrive/network paths)
Public Function BSP_Work_Folder() As String
    Dim p As String
    p = Environ("TEMP") & "\BSP_work"
    If Len(Dir(p, vbDirectory)) = 0 Then MkDir p
    BSP_Work_Folder = p
End Function

'Run the engine. Returns True when the engine finished (engine_done.txt = OK)
Public Function BSP_Run_Engine(mode As String, ByRef errMsg As String) As Boolean

    Dim shell As Object
    Dim cmd As String, rscript As String, engine As String, work As String
    Dim doneFile As String, doneTxt As String

    rscript = BSP_RScript_Path()
    engine = BSP_Clean_Path(BSP_UI.Range("BSP_Engine_Folder").Value)
    work = BSP_Work_Folder()
    doneFile = work & "\engine_done.txt"
    If Len(Dir(doneFile)) > 0 Then Kill doneFile

    cmd = """" & rscript & """ """ & engine & "\" & BSP_ENGINE_ENTRY & """ """ & work & """ """ & engine & """ " & mode

    Application.StatusBar = "BSP: engine R sedang berjalan (" & mode & ")..."
    Set shell = VBA.CreateObject("WScript.Shell")
    Debug.Print cmd
    BSP_Trace "run: " & cmd
    shell.Run cmd, 0, True
    BSP_Trace "engine finished"
    Application.StatusBar = False

    If Len(Dir(doneFile)) = 0 Then
        errMsg = "Engine R tidak selesai. Periksa log: " & work & "\engine_log.txt"
        BSP_Run_Engine = False
        Exit Function
    End If

    doneTxt = BSP_Read_Text(doneFile)
    If Left(doneTxt, 2) = "OK" Then
        BSP_Run_Engine = True
    Else
        errMsg = "Engine R berhenti dengan error:" & vbNewLine & Mid(doneTxt, InStr(doneTxt & vbLf, vbLf) + 1) & _
                 vbNewLine & "Log: " & work & "\engine_log.txt"
        BSP_Run_Engine = False
    End If

End Function

'Full path of Rscript.exe from the chosen R folder (accepts the R home or its bin folder)
Public Function BSP_RScript_Path() As String
    Dim f As String
    f = BSP_Clean_Path(BSP_UI.Range("BSP_R_Folder").Value)
    If Len(Dir(f & "\Rscript.exe")) > 0 Then
        BSP_RScript_Path = f & "\Rscript.exe"
    ElseIf Len(Dir(f & "\bin\Rscript.exe")) > 0 Then
        BSP_RScript_Path = f & "\bin\Rscript.exe"
    Else
        BSP_RScript_Path = ""
    End If
End Function

'Newest R installed under Program Files (used to pre-fill the R folder)
Public Function BSP_Find_R_Folder() As String
    Dim base As String, d As String, best As String
    base = Environ("ProgramFiles") & "\R\"
    If Len(Dir(base, vbDirectory)) = 0 Then Exit Function
    d = Dir(base & "R-*", vbDirectory)
    Do While d <> ""
        If d > best Then best = d
        d = Dir()
    Loop
    If best <> "" Then
        If Len(Dir(base & best & "\bin\Rscript.exe")) > 0 Then BSP_Find_R_Folder = base & best & "\bin"
    End If
End Function

Public Function BSP_Clean_Path(ByVal p As String) As String
    p = Trim(p)
    Do While Right(p, 1) = "\"
        p = Left(p, Len(p) - 1)
    Loop
    BSP_Clean_Path = p
End Function

'===== UTF-8 text files ======================================================

Public Function BSP_Read_Text(path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    BSP_Read_Text = st.ReadText
    st.Close
End Function

Public Sub BSP_Write_Text(path As String, txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub

Public Function BSP_Csv_Field(ByVal v As String) As String
    BSP_Csv_Field = """" & Replace(v, """", """""") & """"
End Function

'Parse a CSV text into a 2D array (1-based rows x columns), row 1 = header.
'Handles quoted fields, doubled quotes and line breaks inside quotes.
Public Function BSP_Parse_Csv(txt As String) As Variant

    Dim rows As New Collection, fields As Collection
    Dim i As Long, n As Long, ch As String, cur As String
    Dim inQ As Boolean, maxCols As Long, r As Long, c As Long
    Dim out() As Variant

    If Left(txt, 1) = ChrW(&HFEFF) Then txt = Mid(txt, 2)
    Set fields = New Collection
    n = Len(txt)
    i = 1
    Do While i <= n
        ch = Mid(txt, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid(txt, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQ = True
                Case ","
                    fields.Add cur
                    cur = ""
                Case vbCr
                    'ignore, handled by vbLf
                Case vbLf
                    fields.Add cur
                    cur = ""
                    rows.Add fields
                    If fields.Count > maxCols Then maxCols = fields.Count
                    Set fields = New Collection
                Case Else
                    cur = cur & ch
            End Select
        End If
        i = i + 1
    Loop
    If cur <> "" Or fields.Count > 0 Then
        fields.Add cur
        rows.Add fields
        If fields.Count > maxCols Then maxCols = fields.Count
    End If

    If rows.Count = 0 Then
        BSP_Parse_Csv = Empty
        Exit Function
    End If

    ReDim out(1 To rows.Count, 1 To maxCols)
    For r = 1 To rows.Count
        For c = 1 To rows(r).Count
            out(r, c) = rows(r)(c)
        Next c
    Next r
    BSP_Parse_Csv = out

End Function

Public Function BSP_Read_Csv(path As String) As Variant
    If Len(Dir(path)) = 0 Then
        BSP_Read_Csv = Empty
    Else
        BSP_Read_Csv = BSP_Parse_Csv(BSP_Read_Text(path))
    End If
End Function

'Column index of a header name in a parsed CSV (0 if missing)
Public Function BSP_Col(data As Variant, header As String) As Long
    Dim c As Long
    For c = LBound(data, 2) To UBound(data, 2)
        If data(1, c) = header Then
            BSP_Col = c
            Exit Function
        End If
    Next c
    BSP_Col = 0
End Function
