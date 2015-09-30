'\\Plataine 2015//
'\\Takes a jobs input file containing ProgramName, PartNumber, FabricName, and QTY and creates a folder output for the kitted files
'\\The four required fields are mapped into a configuration file called "JobCreator.config" it must be located in C:\ProgramData\Plataine
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Module Module1
    'Readconfig globals
    Dim inputJobSheet() As String, kitOutputPath As String, newJobsPath As String, origStyleLib As String, plyNestingFilename As String
    'Column mapping globals
    Dim partNumberInt As Integer, qtyInt As Integer, fabricInt As Integer, programInt As Integer
    'Column headers per input file
    Dim columnHeaders()
    Dim columnCount As Integer, addComma As String
    'Outputfile globals
    Dim newJobsFileName As String
    Sub Main()
        Call ReadConfig()
        addComma = ""
        Try
            Call ReadInputExcel()
        Catch ex As Exception
            'MsgBox(ex.ToString, vbOKOnly)
            Exit Sub
        End Try
    End Sub
    Public Sub ReadInputExcel()
        ' read each of the elements in the input jobs folder Directory.GetFiles(inputjobsheet())
        For Each element As String In inputJobSheet
            Dim app As New Excel.Application
            Dim inputSheet As Excel.Worksheet
            Dim workbook As Excel.Workbook
            workbook = app.Workbooks.Open(element)
            inputSheet = workbook.Worksheets(1)

            Dim countRows As Integer, i As Integer
            i = 1
            columnCount = 0
            countRows = 2

            'how many rows of data are there? Start at 2 to ignore headers
            While (workbook.ActiveSheet.cells(countRows, partNumberInt).value IsNot Nothing)
                countRows = countRows + 1
            End While

            'how many columns? read the header values into an array 
            Do Until workbook.ActiveSheet.cells(1, i).value Is Nothing
                ReDim Preserve columnHeaders(columnCount)
                columnHeaders(columnCount) = workbook.ActiveSheet.cells(1, i).value
                i = i + 1
                columnCount = columnCount + 1
            Loop

            'correct the column count if there is no programs column
            If columnCount < programInt Then
                columnCount = columnCount + 1
                addComma = ","
            End If

            'create our output CSV with the same headers
            Call PrepareOutput()

            'loop through the rows to create input folder string, and then pass the data to the folder functions
            For i = 2 To countRows - 1
                Dim InputFolder As String = origStyleLib & inputSheet.Cells(i, programInt).value & "\" & inputSheet.Cells(i, partNumberInt).value & "\" & inputSheet.Cells(i, fabricInt).value & "\"
                Dim j As Integer, x As Integer
                Dim rowData() As String
                x = 0
                j = 1
                'take the entire row and dump it into an array call rowData
                Do Until j = columnCount + 1
                    ReDim Preserve rowData(x)
                    rowData(x) = workbook.ActiveSheet.cells(i, j).value
                    j = j + 1
                    x = x + 1
                Loop
                'figure out which folder organizer to run by looking at the plynestinginstructions file
                If File.Exists(InputFolder & Path.GetFileNameWithoutExtension(plyNestingFilename) & ".csv") Then
                    AbsoluteFolders(InputFolder & Path.GetFileNameWithoutExtension(plyNestingFilename) & ".csv", rowData, InputFolder)
                Else
                    Dim PlyNestingFile() As String, maxPlies As String
                    If File.Exists(InputFolder & plyNestingFilename) Then
                        PlyNestingFile = File.ReadAllLines(InputFolder & plyNestingFilename)
                        maxPlies = String.Join(".", PlyNestingFile)
                    Else
                        maxPlies = ""
                    End If
                    If CheckAdvanced(maxPlies, InputFolder) = True Then
                        ReturnAdvancedFolders(maxPlies, rowData, InputFolder)
                    Else
                        ReturnFolders(maxPlies, rowData, InputFolder)
                    End If
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app.ActiveWorkbook)
            app.ActiveWorkbook.Close()
        Next
    End Sub
    Public Sub AbsoluteFolders(plynestingfile As String, ByVal rows() As String, inputPath As String)
        Try
            'create destination folders
            Dim path As String = rows(programInt - 1) & "." & rows(partNumberInt - 1) & "-" & rows(fabricInt - 1) & "-Absolute"
            Dim newRoot As String = kitOutputPath & "\" & path
            If (Not Directory.Exists(newRoot)) Then
                Directory.CreateDirectory(newRoot)
            End If
            'read csv file
            Dim csvReader() As String = File.ReadAllLines(plynestingfile)
            Dim line As Integer, plyFile() As String, plyInt As Integer
            For line = 0 To csvReader.Length - 1
                Dim subRoot As String = newRoot & "\" & line + 1 & "-" & csvReader.Length
                If (Not Directory.Exists(subRoot)) Then
                    Directory.CreateDirectory(subRoot)
                End If
                plyFile = Split(csvReader(line), ",")
                Dim styles As String = ""
                For plyInt = 0 To plyFile.Count - 1
                    If File.Exists(subRoot & "\" & System.IO.Path.GetFileName(plyFile(plyInt))) Then
                        If Not File.GetLastWriteTime(inputPath & "\" & plyFile(plyInt)) = File.GetLastWriteTime(subRoot & "\" & System.IO.Path.GetFileName(plyFile(plyInt))) Then
                            File.Copy(inputPath & "\" & plyFile(plyInt), subRoot & "\" & System.IO.Path.GetFileName(plyFile(plyInt)))
                        End If
                    Else
                        File.Copy(inputPath & "\" & plyFile(plyInt), subRoot & "\" & System.IO.Path.GetFileName(plyFile(plyInt)))
                    End If
                    styles = System.IO.Path.GetFileNameWithoutExtension(plyFile(plyInt)) & "&" & styles
                Next
                For z As Integer = 1 To rows(qtyInt - 1)
                    If Right(styles, 1) = "&" Then styles = Left(styles, Len(styles) - 1)
                    AppendJobFile(rows, subRoot, plyInt, styles)
                Next
            Next
        Catch ex As Exception
            AppendJobFile(rows, "MAT NAME ERROR or PLY NESTING CSV ERROR", "Error", "")
        End Try
    End Sub
    Public Sub ReturnAdvancedFolders(plynestingfile As String, ByVal rows() As String, inputPath As String)
        Dim maxnumber() As String = Split(plynestingfile, ",")
        Dim path As String = rows(programInt - 1) & "." & rows(partNumberInt - 1) & "-" & rows(fabricInt - 1) & "-advanced"
        Dim i As Integer, numberCount As Integer, z As Integer
        Try
            Dim kitCounter As Integer = Directory.GetFiles(inputPath, "*.dxf").Count
            Dim kitarray As String() = Directory.GetFiles(inputPath, "*.dxf")
            Array.Sort(kitarray)
            Dim start As Integer, finish As Integer, newRoot As String
            newRoot = kitOutputPath & "\" & path
            If (Not Directory.Exists(newRoot)) Then
                Directory.CreateDirectory(newRoot)
            End If
            start = 0
            finish = 0
            For numberCount = 0 To maxnumber.Count - 1
                Dim subDirectory As String = newRoot & "\" & numberCount + 1 & "-" & maxnumber(numberCount)
                Directory.CreateDirectory(subDirectory)
                start = finish
                finish = start + maxnumber(numberCount)
                If finish > kitCounter Then finish = kitCounter
                Dim styles As String = ""
                For i = start To finish - 1
                    If File.Exists(subDirectory & "\" & System.IO.Path.GetFileName(kitarray(i))) Then
                        If Not File.GetLastWriteTime(kitarray(i)) = File.GetLastWriteTime(subDirectory & "\" & System.IO.Path.GetFileName(kitarray(i))) Then
                            File.Copy(kitarray(i), subDirectory & "\" & System.IO.Path.GetFileName(kitarray(i)))
                        End If
                    Else
                        File.Copy(kitarray(i), subDirectory & "\" & System.IO.Path.GetFileName(kitarray(i)))
                    End If
                    styles = System.IO.Path.GetFileNameWithoutExtension(kitarray(i)) & "&" & styles
                Next
                ' make entries in new job file for each quantity ("split jobs")
                If Right(styles, 1) = "&" Then styles = Left(styles, Len(styles) - 1)
                For z = 1 To rows(qtyInt - 1)
                    AppendJobFile(rows, subDirectory, maxnumber(numberCount), styles)
                Next
            Next
        Catch ex As Exception
            'Call MsgBox("An error occured:" & Chr(13) & Chr(13) & ex.ToString)
            AppendJobFile(rows, "MAT NAME ERROR", "Error", "")
            Exit Sub
        End Try
    End Sub
    Public Sub ReturnFolders(plynestingfile As String, ByVal rows() As String, inputPath As String)
        Dim plySplit() As String = Split(plynestingfile, ",")
        If plynestingfile = "" Then plynestingfile = "0"
        Dim maxnumber As Integer = CInt(plynestingfile)
        If plySplit.Count > 1 Then plynestingfile = "all"
        Dim path As String = rows(programInt - 1) & "." & rows(partNumberInt - 1) & "-" & rows(fabricInt - 1) & "-" & maxnumber
        If maxnumber = 0 Or IsNothing(maxnumber) Then
            maxnumber = 100000
        End If
        Try
            Dim kitCounter As Integer = Directory.GetFiles(inputPath, "*.dxf").Count
            Dim kitarray As String() = Directory.GetFiles(inputPath, "*.dxf")
            Array.Sort(kitarray)
            Dim i As Integer, j As Integer, start As Integer, finish As Integer, newRoot As String, z As Integer
            i = 1
            If kitCounter > maxnumber Then
                ' move first (maxnumber) plies here since kitcounter > maxnumber
                Dim divide = Math.Ceiling(kitCounter / maxnumber)
                newRoot = kitOutputPath & "\" & path & "." & divide
                If (Not Directory.Exists(newRoot)) Then
                    Directory.CreateDirectory(newRoot)
                End If
                For i = 1 To divide
                    Dim subDirectory As String = newRoot & "\" & i & "-" & maxnumber
                    Directory.CreateDirectory(subDirectory)
                    start = (i - 1) * maxnumber
                    finish = i * maxnumber
                    If finish > kitCounter Then finish = kitCounter
                    Dim styles As String = ""
                    For j = start To finish - 1
                        If File.Exists(subDirectory & "\" & System.IO.Path.GetFileName(kitarray(j))) Then
                            If Not File.GetLastWriteTime(kitarray(j)) = File.GetLastWriteTime(subDirectory & "\" & System.IO.Path.GetFileName(kitarray(j))) Then
                                File.Delete(subDirectory & "\" & System.IO.Path.GetFileName(kitarray(j)))
                                File.Copy(kitarray(j), subDirectory & "\" & System.IO.Path.GetFileName(kitarray(j)))
                            End If
                        Else
                            File.Copy(kitarray(j), subDirectory & "\" & System.IO.Path.GetFileName(kitarray(j)))
                        End If
                        styles = System.IO.Path.GetFileNameWithoutExtension(kitarray(j)) & "&" & styles
                    Next
                    ' make entries in new job file for each quantity ("split jobs")
                    If Right(styles, 1) = "&" Then styles = Left(styles, Len(styles) - 1)
                    For z = 1 To rows(qtyInt - 1)
                        AppendJobFile(rows, subDirectory, maxnumber, styles)
                    Next
                Next
            Else
                ' move all plies from kit here since kitcounter < maxnumber
                newRoot = kitOutputPath & "\" & path & "Kit"
                If (Not Directory.Exists(newRoot)) Then
                    Directory.CreateDirectory(newRoot)
                End If
                Dim styles As String = ""
                For j = 0 To kitCounter - 1
                    If File.Exists(newRoot & "\" & System.IO.Path.GetFileName(kitarray(j))) Then
                        If Not File.GetLastWriteTime(kitarray(j)) = File.GetLastWriteTime(newRoot & "\" & System.IO.Path.GetFileName(kitarray(j))) Then
                            File.Delete(newRoot & "\" & System.IO.Path.GetFileName(kitarray(j)))
                            File.Copy(kitarray(j), newRoot & "\" & System.IO.Path.GetFileName(kitarray(j)))
                        End If
                    Else
                        File.Copy(kitarray(j), newRoot & "\" & System.IO.Path.GetFileName(kitarray(j)))
                    End If
                    styles = System.IO.Path.GetFileNameWithoutExtension(kitarray(j)) & "&" & styles
                Next
                ' make entries in new job file for each quantity ("split jobs")
                If Right(styles, 1) = "&" Then styles = Left(styles, Len(styles) - 1)
                For z = 1 To rows(qtyInt - 1)
                    AppendJobFile(rows, newRoot, "all", styles)
                Next
            End If
        Catch ex As Exception
            'Call MsgBox("An error occured:" & Chr(13) & Chr(13) & ex.ToString)
            AppendJobFile(rows, "MAT NAME ERROR", "ERROR", "")
            Exit Sub
        End Try
    End Sub
    Public Sub AppendJobFile(ByVal passedRows() As String, newStylePath As String, maxPlies As String, styleList As String)
        ''put some stuff here that appends the new job to the end of the newjobs.csv
        If (Not Directory.Exists(newJobsPath)) Then
            Directory.CreateDirectory(newJobsPath)
        End If
        Dim sw As New StreamWriter(newJobsFileName, True)
        'count commas for headers and compare with data. Make new variable to add them below
        Dim i As Integer, commas As String
        For i = 0 To columnCount - passedRows.Length
            commas = commas & ","
        Next
        sw.WriteLine(String.Join(",", passedRows) & commas & newStylePath & "," & maxPlies & "," & styleList)
        sw.Close()
    End Sub
    Public Sub PrepareOutput()
        'prepare csv output
        If (Not Directory.Exists(newJobsPath)) Then
            Directory.CreateDirectory(newJobsPath)
        End If
        newJobsFileName = newJobsPath & DateTime.Now.ToString("MM-dd.HH-mm") & "-JobCreatorOutput.csv"
        If (File.Exists(newJobsFileName)) Then newJobsFileName = newJobsPath & DateTime.Now.ToString("MM-dd.HH-mm-ss") & "-JobCreatorOutput.csv"
        Dim sw As New StreamWriter(newJobsFileName, True)
        sw.WriteLine(String.Join(", ", columnHeaders) & addComma & ",New Ply Folder Location,Plys Per Job,StyleList")
        sw.Close()
    End Sub
    Public Function CheckAdvanced(input As String, validateFolder As String) As Boolean
        Dim test() As String, testSum As Integer, l As Integer
        test = Split(input, ",")
        If test.Length > 1 Then
            For l = 0 To test.Count - 1
                testSum = testSum + test(l)
            Next
            If Directory.GetFiles(validateFolder, "*.dxf").Count = testSum Then
                CheckAdvanced = True
            Else
                CheckAdvanced = False
            End If
        Else
            CheckAdvanced = False
        End If
    End Function
    Public Sub ReadConfig()
        Try
            Dim configFile() As String = File.ReadAllLines("C:\ProgramData\Plataine\JobCreator.config")
            For Each line As String In configFile
                Dim setting() As String = Split(line, "=")
                If UCase(setting(0)) = "INPUTFILE" Then
                    inputJobSheet = Directory.GetFiles(setting(1))
                ElseIf UCase(setting(0)) = "KITOUTPUTPATH" Then
                    kitOutputPath = setting(1).ToString
                    If Right(kitOutputPath, 1) = "\" Then kitOutputPath = Left(kitOutputPath, Len(kitOutputPath) - 1)
                ElseIf UCase(setting(0)) = "NEWJOBSPATH" Then
                    newJobsPath = setting(1).ToString
                    If Not Right(newJobsPath, 1) = "\" Then newJobsPath = newJobsPath & "\"
                ElseIf UCase(setting(0)) = "STYLELIBRARY" Then
                    origStyleLib = setting(1).ToString
                    If Not Right(origStyleLib, 1) = "\" Then origStyleLib = origStyleLib & "\"
                ElseIf UCase(setting(0)) = "PLYNESTINGFILENAME" Then
                    plyNestingFilename = setting(1).ToString
                    If Not Path.GetExtension(plyNestingFilename) = ".txt" Then Path.ChangeExtension(plyNestingFilename, ".txt")
                    'column mappings:
                ElseIf UCase(setting(0)) = "STYLEDIRECTORY" Then
                    programInt = CInt(setting(1).ToString)
                ElseIf UCase(setting(0)) = "PARTNUMBER" Then
                    partNumberInt = CInt(setting(1).ToString)
                ElseIf UCase(setting(0)) = "FABRICNAME" Then
                    fabricInt = CInt(setting(1).ToString)
                ElseIf UCase(setting(0)) = "QTY" Then
                    qtyInt = CInt(setting(1).ToString)
                End If
            Next
            If IsNothing(plyNestingFilename) Or IsNothing(origStyleLib) Or IsNothing(newJobsPath) Or IsNothing(kitOutputPath) Or IsNothing(inputJobSheet) _
                Or IsNothing(partNumberInt) Or IsNothing(qtyInt) Or IsNothing(fabricInt) Or IsNothing(programInt) Then
                Call MsgBox("Your config file is invalid. Must be of the form:" _
                           & Chr(13) & "inputfile=pathtojobs\" _
                           & Chr(13) & "kitoutputpath=path" _
                           & Chr(13) & "newjobspath=path\" _
                           & Chr(13) & "StyleLibrary=path\" _
                           & Chr(13) & "plynestingfilename=filename.txt" _
                           & Chr(13) & "Config location must be: C:\ProgramData\Plataine\JobCreator.config")
                End
            End If
        Catch ex As Exception
            Call MsgBox("Your config file is missing, or missing required column mappings." _
                               & Chr(13) & "Config location must be: C:\ProgramData\Plataine\JobCreator.config")
            End
        End Try
    End Sub
End Module
