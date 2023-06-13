Attribute VB_Name = "Módulo3"
Sub ReportByCPV()
    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim wsDestino As Worksheet
    Dim strFolderPath As String, strConn As String
    Dim files() As Variant, file As Variant
    Dim header As Boolean, hasIncidentID As Boolean, hasSubmitter As Boolean
    Dim strSheetName As String
    SharepointAddress = "https://nsg.sharepoint.com/sites/GBS/OtherDocuments/Service Centre Reports/Monthly reports/Remedy Closed Incidents Monthly Report/"

    Set objNet = CreateObject("WScript.Network")
    Set FS = CreateObject("Scripting.FileSystemObject")
    objNet.MapNetworkDrive "A:", SharepointAddress
    
    Set objFolder = FS.GetFolder("A:")
    
    strFolderPath = "A:"
    
    strAssignedGroup = "Assigned Group"
        
    GetExcelFilesInFolder strFolderPath, files()
    
    Set wsDestino = ThisWorkbook.Sheets("Planilha1")
    
    wsDestino.Cells.Clear
    
   
    For Each file In files
    If Not InStr(1, file, "Archive", vbTextCompare) > 0 Then
        hasIncidentID = False
        hasReportedSrc = False
        hasSubmitter = False
    
        ' Connection string for Excel workbooks
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & file & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    
        ' Open connection to the workbook
        Set conn = CreateObject("ADODB.Connection")
        conn.Open strConn
        
        Dim schemaTable As ADODB.Recordset
        Set schemaTable = conn.OpenSchema(adSchemaTables)

        strSheetName = schemaTable.Fields("TABLE_NAME").Value
        Set rs = CreateObject("ADODB.Recordset")
               
        rs.Open "SELECT * FROM [" & strSheetName & "]", conn
        
        ' Verifica se a primeira coluna contém valor
        hasF1Value = False
        If Not rs.EOF Then
            hasF1Value = (Not IsNull(rs.Fields(0).Value))
        End If
        
        If Not hasF1Value Then
            rng = "B2:AF"
            Else
            rng = "A2:AF"
            End If
        
        If Not IsEmpty(strSheetName) Then
            Dim arrHeader() As String
            ' Obtém o header da primeira planilha
            Dim strHeader As String
            strHeader = GetHeaderFromSheet(conn, strSheetName)
            Dim HasFColumn As Boolean
            HasFColumn = False
            
            ' Check if the headers returned are F* columns
            If Left(strHeader, 1) = "F" Then
                
                HasFColumn = True
                ' Adjust the range to retrieve headers from the second line onwards
                Dim rangeString As String
                rangeString = Replace(strSheetName, "'", "") & rng
                strHeader = GetHeaderFromRange(conn, rangeString)
            End If
                    
            arrHeader = Split(strHeader, ",")
            
                For i = LBound(arrHeader) To UBound(arrHeader)
                If arrHeader(i) = "Incident ID" Then
                    hasIncidentID = True
                End If
                If arrHeader(i) = "Reported Source" Then
                    hasReportedSrc = True
                End If
                If arrHeader(i) = "Submitter" Then
                    hasSubmitter = True
                End If
            
             Next i
            
            If WorksheetFunction.CountA(wsDestino.Range("A1:AE1")) = 0 Then
                For i = LBound(arrHeader) To UBound(arrHeader)
                      wsDestino.Cells(1, i + 1).Value = arrHeader(i)
                Next i
            End If
        End If
        
        Dim strSQL As String
        
        If hasIncidentID = True And hasSubmitter = True And HasFColumn = False Then
            strSQL = "SELECT [Incident ID], [Submit Date], [Submitter], [Reported Source], [Full Name], [Country], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Categorization Tier 1], [Categorization Tier 2], [Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category], [Resolution Category Tier 2], [Resolution Category Tier 3], [Closure Product Category Tier1], [Closure Product Category Tier2], [Closure Product Category Tier3], [Status], [Last Resolved Date], [Last Modified Date], [progress], [Service Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Assigned Group] = 'Brazil Back Desk Remote' OR [Assigned Group] = 'South America Front Desk' OR [Assigned Group] = 'South America Service Delivery' " & _
            "UNION " & _
            "SELECT [Incident ID], [Submit Date], [Submitter], [Reported Source], [Full Name], [Country], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Categorization Tier 1], [Categorization Tier 2], [Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category], [Resolution Category Tier 2], [Resolution Category Tier 3], [Closure Product Category Tier1], [Closure Product Category Tier2], [Closure Product Category Tier3], [Status], [Last Resolved Date], [Last Modified Date], [progress], [Service Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Country] ='Argentina' OR [Country] ='Brazil' OR [Country] ='Chile';"
        ElseIf hasIncidentID = True And hasReportedSrc = False And HasFColumn = False Then
            strSQL = "SELECT [Incident ID], [Submit Date], [Created By], Null AS [Reported Source], [Name], [Site Group], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Operational  Categorization Tier 1], [Operational  Categorization Tier 2], [Operational  Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category Tier 1], [Resolution Category Tier 2], [Resolution Category Tier 3], [Resolution Product Category Tier1], [Resolution Product Category Tier2], [Resolution Product Category Tier3], [Status], [Incident Last Resolved Date], [Last Modified Date], [Progress], [Incident Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Assigned Group] = 'Brazil Back Desk Remote' OR [Assigned Group] = 'South America Front Desk' OR [Assigned Group] = 'South America Service Delivery' " & _
            "UNION " & _
            "SELECT [Incident ID], [Submit Date], [Created By], Null AS [Reported Source], [Name], [Site Group], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Operational  Categorization Tier 1], [Operational  Categorization Tier 2], [Operational  Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category Tier 1], [Resolution Category Tier 2], [Resolution Category Tier 3], [Resolution Product Category Tier1], [Resolution Product Category Tier2], [Resolution Product Category Tier3], [Status], [Incident Last Resolved Date], [Last Modified Date], [Progress], [Incident Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Site Group] ='Argentina' OR [Site Group] ='Brazil' OR [Site Group] ='Chile';"
        ElseIf hasIncidentID = True And hasReportedSrc = True And HasFColumn = False Then
            strSQL = "SELECT [Incident ID], [Submit Date], [Created By], [Reported Source], [Name], [Site Group], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Operational  Categorization Tier 1], [Operational  Categorization Tier 2], [Operational  Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category Tier 1], [Resolution Category Tier 2], [Resolution Category Tier 3], [Resolution Product Category Tier1], [Resolution Product Category Tier2], [Resolution Product Category Tier3], [Status], [Incident Last Resolved Date], [Last Modified Date], [Progress], [Incident Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Assigned Group] = 'Brazil Back Desk Remote' OR [Assigned Group] = 'South America Front Desk' OR [Assigned Group] = 'South America Service Delivery' " & _
            "UNION " & _
            "SELECT [Incident ID], [Submit Date], [Created By], [Reported Source], [Name], [Site Group], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Operational  Categorization Tier 1], [Operational  Categorization Tier 2], [Operational  Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category Tier 1], [Resolution Category Tier 2], [Resolution Category Tier 3], [Resolution Product Category Tier1], [Resolution Product Category Tier2], [Resolution Product Category Tier3], [Status], [Incident Last Resolved Date], [Last Modified Date], [Progress], [Incident Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Site Group] ='Argentina' OR [Site Group] ='Brazil' OR [Site Group] ='Chile';"
        ElseIf hasIncidentID = False And hasReportedSrc = True And HasFColumn = False Then
            strSQL = "SELECT [Incident Number], [Submit Date], [Submitter], [Reported Source], [Full Name], [Country], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Categorization Tier 1], [Categorization Tier 2], [Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category], [Resolution Category Tier 2], [Resolution Category Tier 3], [Closure Product Category Tier1], [Closure Product Category Tier2], [Closure Product Category Tier3], [Status], [Last Resolved Date], [Last Modified Date], [progress], [Service Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Assigned Group] = 'Brazil Back Desk Remote' OR [Assigned Group] = 'South America Front Desk' OR [Assigned Group] = 'South America Service Delivery' " & _
            "UNION " & _
            "SELECT [Incident Number], [Submit Date], [Submitter], [Reported Source], [Full Name], [Country], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Categorization Tier 1], [Categorization Tier 2], [Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category], [Resolution Category Tier 2], [Resolution Category Tier 3], [Closure Product Category Tier1], [Closure Product Category Tier2], [Closure Product Category Tier3], [Status], [Last Resolved Date], [Last Modified Date], [progress], [Service Type], [Resolved 30 min], [Resolved 60 min] FROM [" & strSheetName & "] WHERE [Country] ='Argentina' OR [Country] ='Brazil' OR [Country] ='Chile';"
        Else
            strSQL = "SELECT [Incident ID], [Submit Date], [Created By], [Reported Source], [Name], [Site Group], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Operational  Categorization Tier 1], [Operational  Categorization Tier 2], [Operational  Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category Tier 1], [Resolution Category Tier 2], [Resolution Category Tier 3], [Resolution Product Category Tier1], [Resolution Product Category Tier2], [Resolution Product Category Tier3], [Status], [Incident Last Resolved Date], [Last Modified Date], [Progress], [Incident Type], [Resolved 30 min], [Resolved 60 min] FROM [" & Replace(strSheetName, "'", "") & rng & "] WHERE [Assigned Group] = 'Brazil Back Desk Remote' OR [Assigned Group] = 'South America Front Desk' OR [Assigned Group] = 'South America Service Delivery' " & _
            "UNION " & _
            "SELECT [Incident ID], [Submit Date], [Created By], [Reported Source], [Name], [Site Group], [Site], [Summary], [Priority], [Urgency], [Assigned Group], [Assignee], [Operational  Categorization Tier 1], [Operational  Categorization Tier 2], [Operational  Categorization Tier 3], [Product Categorization Tier 1], [Product Categorization Tier 2], [Product Categorization Tier 3], [Resolution Category Tier 1], [Resolution Category Tier 2], [Resolution Category Tier 3], [Resolution Product Category Tier1], [Resolution Product Category Tier2], [Resolution Product Category Tier3], [Status], [Incident Last Resolved Date], [Last Modified Date], [Progress], [Incident Type], [Resolved 30 min], [Resolved 60 min] FROM [" & Replace(strSheetName, "'", "") & rng & "]  WHERE [Site Group] ='Argentina' OR [Site Group] ='Brazil' OR [Site Group] ='Chile';"
        End If
        
        ' Executa a consulta SQL e obtém um objeto Recordset
        Debug.Print strSQL
        Debug.Print file; strSheetName
                

           Set rs = conn.Execute(strSQL)
            
            
            Dim lastRow As Long
            lastRow = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row + 1
            'rs.MoveFirst
            wsDestino.Cells(lastRow, 1).CopyFromRecordset rs

        rs.Close
        conn.Close
        End If

    Next file
    
    ' Determine a última linha e a última coluna com dados preenchidos
    últimaLinha = wsDestino.Cells(Rows.Count, 1).End(xlUp).Row
    últimaColuna = wsDestino.Cells(1, columns.Count).End(xlToLeft).Column
    
    ' Cria um dicionário para rastrear os valores encontrados nas colunas desejadas
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Percorre as linhas da planilha e exclui as duplicatas quando o valor da coluna "Status" for diferente de "Closed" ou "Resolved"
    For i = últimaLinha To 2 Step -1
        If wsDestino.Cells(i, 3).Value <> "Closed" And wsDestino.Cells(i, 3).Value <> "Resolved" Then
            Dim chave As String
            chave = wsDestino.Cells(i, 1).Value & "|" & wsDestino.Cells(i, 2).Value & "|" & wsDestino.Cells(i, 3).Value
            If Not dict.Exists(chave) Then
                dict(chave) = True
            Else
                wsDestino.Rows(i).Delete
            End If
        End If
    Next i
    
    ' Determine o novo intervalo após a exclusão das linhas duplicadas
    últimaLinha = wsDestino.Cells(Rows.Count, 1).End(xlUp).Row
    últimaColuna = wsDestino.Cells(1, columns.Count).End(xlToLeft).Column
    
    ' Define o intervalo com base nas últimas linhas e colunas
    Set intervalo = wsDestino.Range(wsDestino.Cells(1, 1), wsDestino.Cells(últimaLinha, últimaColuna))
    
    ' Cria a tabela a partir do intervalo definido
    Set tabela = wsDestino.ListObjects.Add(xlSrcRange, intervalo, , xlYes)
    
    ' Define o nome da tabela
    tabela.name = "Tabela1"
    
    ' Formata a tabela como uma tabela do Excel
    tabela.TableStyle = "TableStyleMedium2"
    
    ' Ajusta a largura das colunas da tabela
    intervalo.columns.AutoFit
    
    ' Ocultar linhas de grade da planilha
    wsDestino.Activate
    ActiveWindow.DisplayGridlines = False
        
    objNet.RemoveNetworkDrive "A:"
    Set objNet = Nothing
    Set FS = Nothing
End Sub

Function GetHeaderFromSheet(conn As ADODB.Connection, sheetName As String) As String
    Dim rs As ADODB.Recordset
    Set rs = conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, sheetName))
    Dim header As String
    Dim columnDict As Object
    Set columnDict = CreateObject("Scripting.Dictionary")
    
    ' Check if the first line is empty
    Dim isFirstLineEmpty As Boolean
    isFirstLineEmpty = True
    
    If Not rs.EOF Then
        Dim columnName As String
        columnName = rs.Fields("COLUMN_NAME").Value
        
        ' Check if the first column is empty
        If columnName <> "" Then
            isFirstLineEmpty = False
        End If
    End If
    
    If isFirstLineEmpty Then
        ' Move to the second line
        rs.MoveNext
    End If
    
    While Not rs.EOF
        Dim ordinalPosition As Long
        ordinalPosition = rs.Fields("ORDINAL_POSITION").Value
        'Dim columnName As String
        columnName = rs.Fields("COLUMN_NAME").Value
        
        columnDict.Add ordinalPosition, columnName
        rs.MoveNext
    Wend
    
    rs.Close
    
    ' Get the column names in the order of appearance in the header
    Dim i As Long
    For i = 1 To columnDict.Count
        header = header & columnDict(i)
        
        If i < columnDict.Count Then
            header = header & ","
        End If
    Next i
    
    GetHeaderFromSheet = header
End Function

Function GetHeaderFromRange(conn As ADODB.Connection, rangeString As String) As String
    Dim rs As ADODB.Recordset
    Set rs = conn.Execute("SELECT * FROM [" & rangeString & "] WHERE 1=0;")
    Dim header As String
    Dim columnDict As Object
    Set columnDict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To rs.Fields.Count
        columnDict.Add i, rs.Fields(i - 1).name
    Next i
    
    rs.Close
    
    ' Get the column names in the order of appearance in the header
    Dim j As Long
    For j = 1 To columnDict.Count
        header = header & columnDict(j)
        
        If j < columnDict.Count Then
            header = header & ","
        End If
    Next j
    
    GetHeaderFromRange = header
End Function

Sub GetExcelFilesInFolder(strFolder As String, ByRef files() As Variant)
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer
    
    ' Create a new instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get the folder object associated with the folder path
    Set objFolder = objFSO.GetFolder(strFolder)
    
    ' Initialize the file counter to 0
    i = 0
    
    ' Loop through each file in the folder
    For Each objFile In objFolder.files
        ' Check if the file is an Excel file
        If LCase(Right(objFile.name, 4)) = ".xls" Or LCase(Right(objFile.name, 5)) = ".xlsx" Then
            ' Add the file path to the array
            ReDim Preserve files(i)
            files(i) = objFile.Path
            
            ' Increment the file counter
            i = i + 1
        End If
    Next objFile
    
    ' Call the GetExcelFilesInSubfolders subroutine to get Excel files in subfolders
    GetExcelFilesInSubfolders objFolder, files, i
    
    ' Cleanup
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
End Sub

Sub GetExcelFilesInSubfolders(ByVal objFolder As Object, ByRef files() As Variant, ByRef i As Integer)
    Dim objSubFolder As Object
    Dim objFile As Object
    Dim j As Integer
    
    ' Loop through each subfolder in the folder
    For Each objSubFolder In objFolder.SubFolders
        ' Initialize the file counter for this subfolder to 0
        j = 0
        
        ' Loop through each file in the subfolder
        For Each objFile In objSubFolder.files
            ' Check if the file is an Excel file
            If LCase(Right(objFile.name, 4)) = ".xls" Or LCase(Right(objFile.name, 5)) = ".xlsx" Then
                ' Add the file path to the array
                ReDim Preserve files(i + j)
                files(i + j) = objFile.Path
                
                ' Increment the file counter for this subfolder
                j = j + 1
            End If
        Next objFile
        
        ' Add the file counter for this subfolder to the main file counter
        i = i + j
        
        ' Call the GetExcelFilesInSubfolders subroutine recursively to get Excel files in sub-subfolders
        GetExcelFilesInSubfolders objSubFolder, files, i
    Next objSubFolder
End Sub




