Attribute VB_Name = "fs_"
Option Explicit

'Path helpers
'============
    'Get path file extension
    '-----------------------
    'i.e.: "C:\example.txt" ---> .txt
    Function pathExt(pathStr As String) As String
        Dim i As Long
        i = InStrRev(pathStr, ".")
        If i = 0 Then
            pathExt = ""
        Else
            pathExt = Right(pathStr, Len(pathStr) - i + 1)
        End If
    End Function

    'Get path dir
    '------------
    Function pathDir(pathStr As String) As String
        If pathExt(pathStr) = "" Then
            pathDir = pathStr 'no file at the end
        Else
            pathDir = Left(pathStr, InStrRev(pathStr, Application.PathSeparator))
        End If
    End Function

    'Get path filename
    '-----------------
    'Filename without extension
    Function pathFile(pathStr As String) As String
        Dim lenDir As Long
        lenDir = Len(pathDir(pathStr))

        Dim ext As String
        ext = pathExt(pathStr)

        If ext = "" Then
            pathFile = ""
        Else
            pathFile = Mid( _
                            pathStr, _
                            lenDir + 1, _
                            Len(pathStr) - lenDir - Len(ext) _
                            )
        End If
    End Function

    'True/false if path is dir
    '-------------------------
    Function pathIsDir(pathStr As String) As Boolean
        Dim i As Long
        i = InStrRev(pathStr, ".")

        If i = 0 Then
            pathIsDir = True
        Else
            pathIsDir = False
        End If
    End Function

'Dialog helpers
'==============
    'Returns file filter string
    '--------------------------
    Function fileFilter(filePath As String) As String
        'Set file filter
        Dim ext As String
        ext = pathExt(filePath)
        If ext = ".xlsm" Then   'Macro-enabled Excel file
            fileFilter = "Excel Macro-Enabled Workbook (*.xlsm),*.xlsm"
        ElseIf ext = ".xlsx" Then   'Default Excel file
            fileFilter = "Excel Workbook (*.xlsx),*.xlsx"
        ElseIf ext = ".xltm" Then   'Macro-enabled Excel template
            fileFilter = "Excel Macro-Enabled Template (*.xltm),*.xltm"
        ElseIf ext = ".xltx" Then   'Excel template
            fileFilter = "Excel Template (*.xltx),*.xltx"
        ElseIf ext = ".xlam" Then   'Excel Add-in
            fileFilter = "Excel Add-In (*.xlam),*.xlam"
        ElseIf ext = ".xml" Then    'XML file
            fileFilter = "XML Data (*.xml),*.xml"
        ElseIf ext = ".txt" Then    '.txt file
            fileFilter = "Unicode Text (*.txt),*.txt"
        ElseIf ext = ".pdf" Then    '.pdf file
            fileFilter = "PDF (*.pdf),*.pdf"
        ElseIf ext = ".csv" Then    '.csv file
            fileFilter = "CSV (Comma delimited) (*.csv),*.csv"
        Else                        'All other files
            fileFilter = "All files (*.*),*.*"
        End If
    End Function

    'Display save dialog
    '-------------------
    'Returns the path the user selected
    'or if canceled an empty string
    Function saveDialog(filePath As String) As String
        Dim exitLoop As Boolean
        exitLoop = False
        Dim overwrite As Boolean
        Dim savePath As Variant
        savePath = filePath
        Do While exitLoop = False
            'Prompt user to choose savePath
            savePath = Application _
                        .GetSaveAsFilename( _
                                            InitialFileName:=savePath, _
                                            fileFilter:=fileFilter(filePath) _
                                            )
            If savePath = False Then
                'User canceled save
                saveDialog = ""
                exitLoop = True
            ElseIf Dir(savePath) = "" Then
                saveDialog = savePath
                exitLoop = True
            Else
                'File exists
                overwrite = alertYesNo( _
                                        "Press [Yes] to overwite or [No] to choose " & _
                                        "a new destination folder or filename.")
                If overwrite = True Then
                    'Overwrite file
                    saveDialog = savePath
                    exitLoop = True
                End If
            End If
        Loop
    End Function

    'Display open dialog
    '-------------------
    'Returns an array of filePaths or
    'if user cancels returns an empty array
    'btnOK - if true, instead of "Open" the
    'confirm buttok will say "OK"
    Function openDialog( _
                        filePath As String, _
                        Optional ByVal btnOK As Boolean = False _
                        ) As Variant
        'Dialog button
        Dim action As Integer
        If btnOK = True Then
            action = msoFileDialogFilePicker
        Else
            action = msoFileDialogOpen
        End If

        'Filters
        Dim fileExt As String
        fileExt = pathExt(filePath)
        Dim filter As String
        filter = fileFilter(fileExt)
        Dim none As String
        none = "All files (*.*),*.*"

        Dim result As Variant
        Dim i As Long
        With Application.FileDialog(action)
            'Remove default filters
            Call .filters.Clear
            'Add filter
            If filter <> none Then
                Call .filters.Add(filter, "*" & fileExt)
            End If
            Call .filters.Add("All files (*.*),*.*", "*.*")

            If .Show = 0 Then
                'User canceled
                result = Array()
            Else
                ReDim result(1 To .SelectedItems.Count) As String
                For i = 1 To .SelectedItems.Count
                    result(i) = .SelectedItems(i)
                Next i
            End If
        End With

        openDialog = result
    End Function

    'Display choose folder dialog
    '----------------------------
    Function folderDialog(folderPath As String) As String
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = pathDir(folderPath)
            If .Show = 0 Then
                folderDialog = ""
            Else
                folderDialog = .SelectedItems(1)
            End If
        End With
    End Function

'Save
'====
    'Save workbook
    '-------------
    'Default is current workbook
    'filePath should include filename and file extension
    'filetype is derived from filepath
    Function saveWorkbook( _
                        filePath As String, _
                        Optional ByVal wbName As String = "", _
                        Optional ByVal userPrompt As Boolean = False _
                        ) As Boolean
        'Default to active workbook
        If wbName = "" Then
            wbName = ActiveWorkbook.Name
        End If

        'File format
        Dim ext As String
        ext = pathExt(filePath)
        Dim fileFormat As Integer

        If ext = ".xlsm" Then
            fileFormat = xlOpenXMLWorkbookMacroEnabled
        ElseIf ext = ".xltm" Then
            fileFormat = xlOpenXMLTemplateMacroEnabled
        ElseIf ext = ".xlam" Then
            fileFormat = xlOpenXMLAddIn
        ElseIf ext = ".xlsx" Then
            fileFormat = xlOpenXMLWorkbook
        ElseIf ext = ".xltx" Then
            fileFormat = xlOpenXMLTemplate
        ElseIf ext = ".txt" Then
            fileFormat = xlCurrentPlatformText
        ElseIf ext = ".csv" Then
            fileFormat = xlCSV
        Else
            'To-do:
            'More robut way to handle
            'format not on the list
            err.Raise 52
        End If

        'Save workbook
        Dim writePath As Variant

        If userPrompt = True Then
            writePath = saveDialog(filePath)
        Else
            writePath = filePath
        End If

        If writePath = False Then
            'User canceled
            saveWorkbook = False
        Else
            'TO-DO: validate writepath
            'Write file
            Application.DisplayAlerts = False
            Application.Workbooks(wbName).saveAs _
                                fileName:=writePath, _
                                fileFormat:=fileFormat, _
                                ConflictResolution:=xlUserResolution
            Application.DisplayAlerts = True
            saveWorkbook = True
        End If
    End Function

    'Save a (non-binary) file
    '------------------------
    Function saveFile( _
                        filePath As String, _
                        fileContent As String, _
                        Optional ByVal userPrompt = False _
                        ) As Boolean
        Dim writePath As Variant

        If userPrompt = True Then
            writePath = saveDialog(filePath)
        Else
            writePath = filePath
        End If

        If writePath = False Then
            'User canceled
            saveFile = False
        Else
            'TO-DO: validate writepath
            'Write file
            Dim FSO As Object
            Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
            Dim txtStream As Object
            Set txtStream = FSO.CreateTextFile(writePath, True, True)
            txtStream.WriteLine (fileContent)
            txtStream.Close
            saveFile = True
        End If
    End Function

'Copy
'====
    'Can use wildcare paths to copy multiple files
    '(in that case clonePath should be a folder)
    Function cloneFile( _
                        originalPath As String, _
                        clonePath As String _
                        ) As Boolean
        'TO-DO:
        'check if originalPath and clonePath are valid
        'add user prompt feature
        FileCopy originalPath, clonePath
        cloneFile = True
    End Function

'Read
'====
    'Read excel file
    '---------------
    'Returns a collection of sheets
    'Each sheet is a 2d array of cells
    Function readWorkbook( _
                            filePath As String, _
                            Optional ByVal startRow As Long = 1, _
                            Optional ByVal startColumn As Integer = 1, _
                            Optional ByVal endRow As Long = -1, _
                            Optional ByVal endColumn As Integer = -1, _
                            Optional ByVal getFormula As Boolean = False _
                            ) As Collection
        'Hide user prompts
        Application.DisplayAlerts = False

        'Remember current active workbook
        Dim activeWb As String
        activeWb = Application.ActiveWorkbook.Name

        'Open workbook
        Dim openWb As Workbook
        Set openWb = Application _
                    .Workbooks _
                    .Open _
                        (fileName:=filePath)
        'Get data
        Dim result As New Collection
        Dim i As Integer
        Dim sheetName As String
        For i = 1 To openWb.Sheets.Count
            sheetName = openWb.Sheets(i).Name
            result.Add _
                        Key:=sheetName, _
                        Item:=getMatrix( _
                            startRow, _
                            startColumn, _
                            endRow, _
                            endColumn, _
                            sheetName, _
                            openWb.Name, _
                            getFormula)
        Next i

        'Close workbook
        openWb.Close saveChanges:=False

        'Return active workbook
        Application.Workbooks(activeWb).Activate

        'Show user prompts
        Application.DisplayAlerts = True

        'Return data
        Set readWorkbook = result
    End Function

    Function readXML( _
                    filePath As String, _
                    Optional ByVal endRow As Long = -1, _
                    Optional ByVal endColumn As Integer = -1, _
                    Optional ByVal getFormula As Boolean = False _
                    ) As Variant
            'Hide user prompts
            Application.DisplayAlerts = False

            'Remember current active workbook
            Dim activeWb As String
            activeWb = Application.ActiveWorkbook.Name

            'Open workbook
            Dim openWb As Workbook
            Set openWb = Application _
                        .Workbooks _
                        .OpenXML( _
                            fileName:=filePath, _
                            LoadOption:=xlXmlLoadImportToList)
            'Get data
            Dim result As Variant
            result = getMatrix(1, 1, endRow, endColumn, "Sheet1", openWb.Name, getFormula)

            'Close workbook
            openWb.Close saveChanges:=False

            'Return active workbook
            Application.Workbooks(activeWb).Activate

            'Show user prompts
            Application.DisplayAlerts = True

            'Return data
            readXML = result
    End Function

    'Read (non-binary) file
    '----------------------
    Function readFile(filePath As String) As String
        Dim FSO As Object
        Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
        Dim txtStream As Object

        'Open file at filePath, for reading, parse as Unicode
        'https://msdn.microsoft.com/en-us/library/aa265347(v=vs.60).aspx
        Set txtStream = FSO.OpenTextFile(filePath, 1, 1)

        'Return file content
        readFile = txtStream.readAll
    End Function

    'Parse a CSV string
    '------------------
    Private Function parseCSV( _
                    str As String, _
                    delimit As Variant _
                    ) As Variant
        'Clean the string input
            str = normalizeNewLines(str, "\n")
            'Remove leading and trailing linebreaks
            str = trimStr(str, "\n")
            'Remove empty lines
            str = singleStr(str, "\n")

        'Handle multiple delimiters
            If IsArray(delimit) Then
                Dim d As String
                d = delimit(LBound(delimit))
                Dim i As Integer
                For i = LBound(delimit) To UBound(delimit)
                    str = Replace(str, delimit(i), d)
                Next i
                delimit = d
            End If

        'Interpret input string as CSV
            Dim lines As Variant
            lines = Split(str, "\n")

            Dim result As Variant
            ReDim result(UBound(lines) - LBound(lines)) As Variant
            Dim j As Long
            For i = LBound(lines) To UBound(lines)
                result(i) = Split(lines(i), delimit)
                'Remove enclosing ""
                For j = LBound(result(i)) To UBound(result(i))
                   result(i)(j) = trimStr(CStr(result(i)(j)), """")
                Next j
            Next i

        parseCSV = result
    End Function
    Function readCSV(filePath As String, delimiter As Variant) As Variant
        Dim CSV As String
        CSV = readFile(filePath)

        readCSV = parseCSV(CSV, delimiter)
    End Function

    'Read fixed column width file
    '----------------------------
    'Return a matrix (2d array)
    'Base index is 1
    Function readFWC( _
                    filePath As String, _
                    rowWidth, _
                    ParamArray columnWidth() _
                    ) As Variant
        Dim datastr As String
        datastr = readFile(filePath)

        Dim rowCount As Long
        rowCount = Len(datastr) / rowWidth

        Dim result As Variant
        ReDim result(1 To rowCount) As Variant

        Dim i As Long
        Dim j As Long
        Dim w As Long
        Dim values As Variant
        Dim row As String
        Dim indexInRow As Long
        For i = 1 To rowCount
            row = Left(datastr, rowWidth)
            datastr = Right(datastr, Len(datastr) - rowWidth)
            ReDim values(1 To UBound(columnWidth) + 1)
            indexInRow = 1
            For j = LBound(columnWidth) To UBound(columnWidth)
                w = columnWidth(j)
                values(j + 1) = Mid(row, indexInRow, w) 'column value
                indexInRow = indexInRow + w
            Next j

            result(i) = values
        Next i

        readFWC = result
    End Function

'Delete
'======
    Function delFile( _
        filePath As String, _
        Optional askUser As Boolean = True, _
        Optional askMsg As String = "Delete File?" _
        ) As Boolean
        Dim del As Boolean
        del = True
        If askUser = True Then
            askUser = alertYesNo(askMsg & vbNewLine & vbNewLine & filePath)
        End If

        If del = True Then
            Dim FSO As Object
            Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
            With FSO
                If .FileExists(filePath) Then
                    .DeleteFile filePath
                End If
            End With
        End If

        delFile = del
    End Function

List
====
    'Lists files (non-recursive)
    '---------------------------
    Function listFiles(dirPath As String) As Variant
        Dim filePath As String
        Dim files As String

        filePath = Dir(dirPath)
        Do Until filePath = vbNullString
            files = files & filePath & "|" 'path separataor: |
            filePath = Dir()
        Loop

        'Remove trailing path separator
        If Len(files) > 0 Then
            files = Left(files, (Len(files) - 1))
        End If

        listFiles = Split(files, "|")
    End Function

    'Lists subdirectories (non-recursive)
    '------------------------------------
    Function listDirs(dirPath)
        Dim subDirPath As String
        Dim subDirs As String

        subDirPath = Dir(dirPath, vbDirectory)
        Do Until subDirPath = vbNullString
            If pathIsDir(subDirPath) = True Then
                subDirs = subDirs & subDirPath & "|" 'path separataor: |
            End If
            subDirPath = Dir()
        Loop

        'Remove trailing path separator
        If Len(subDirs) > 0 Then
            subDirs = Left(subDirs, (Len(subDirs) - 1))
        End If

        listDirs = Split(subDirs, "|")
    End Function
