Attribute VB_Name = "DOWNLOADER"

Private Function SpecifyDownloadFolder(DOWNLOAD_DIRECTORY, FILE_NAME, URL, EXTENSION)



    Dim filename As String, myFolder As Object
    Dim htmlas As Object, htmla As Object, html As Object
    Dim stream As Object, tempArr As Variant
    Dim fileSource As String

    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fileSource = URL
    
    Dim http As New ServerXMLHTTP60, htmldoc As New HTMLDocument
    With http
        .Open "GET", fileSource, False
        .Send
    End With

    Set stream = CreateObject("ADODB.Stream")

    stream.Open
    stream.Type = 1
    stream.Write http.ResponseBody
    

    On Error GoTo err:
    stream.SaveToFile (DOWNLOAD_DIRECTORY & "\" & FILE_NAME & "." & EXTENSION)
    
err:
    stream.Close
    

    
End Function

Public Function MyMkDir(sPath As String)
    Dim iStart          As Integer
    Dim aDirs           As Variant
    Dim sCurDir         As String
    Dim I               As Integer
 
    If sPath <> "" Then
        aDirs = Split(sPath, "\")
        If Left(sPath, 2) = "\\" Then
            iStart = 3
        Else
            iStart = 1
        End If
 
        sCurDir = Left(sPath, InStr(iStart, sPath, "\"))
 
        For I = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(I) & "\"
            If Dir(sCurDir, vbDirectory) = vbNullString Then
                MkDir sCurDir
            End If
        Next I
    End If
End Function


Private Function Col_Letter(lngCol As Variant)
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function





Sub DOWNLOAD_PHOTOS_SELECTION()

    Dim DESTINO

    DESTINO = ThisWorkbook.Path & "\FOTOS\" & ActiveSheet.Name

    
    If Selection.Cells.Rows.Count > 1 Then
        
        RF = Selection.Cells(Selection.Cells.Rows.Count, 1).Row
        RI = Selection.Cells(1, 1).Row
        
        COLU = Col_Letter(Selection.Cells.Column)
        
        Dim celda As Range
    
        For Each celda In Range(COLU & RI & ":" & COLU & RF).SpecialCells(xlCellTypeVisible)
            
            celda.Select
            FILA = ActiveCell.Row
            Range("B" & FILA).Select
            
            If ActiveCell <> "" Then
            
                TRAMITE = ActiveCell
                
                MyMkDir (DESTINO & "\" & TRAMITE)
                
                
                For I = 0 To 10
                
                    foto = Range("D" & FILA).Offset(0, I)
                    encabezado = Range("D2").Offset(0, I)
                    
                    If InStr(1, foto, "https://", vbTextCompare) = 1 Then
                    
                        SpecifyDownloadFolder DESTINO & "\" & TRAMITE, encabezado, foto, "jpg"
                    
                    End If
                    
                Next
                         
                
            End If
            
            Range("A" & FILA) = "OK"
            
        Next
        
        
    ElseIf ActiveCell <> "" Then
        
        FILA = ActiveCell.Row
        Range("B" & FILA).Select
                
        TRAMITE = ActiveCell
        
                
        MyMkDir (DESTINO & "\" & TRAMITE)
        
        
        For I = 0 To 10
        
            foto = Range("D" & FILA).Offset(0, I)
            encabezado = Range("D2").Offset(0, I)
            
            If InStr(1, foto, "https://", 1) = 1 Then
            
                SpecifyDownloadFolder DESTINO & "\" & TRAMITE, encabezado, foto, "jpg"
            
            End If
            
        Next
        
        Range("A" & FILA) = "OK"
     
    End If
End Sub

