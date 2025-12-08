Attribute VB_Name = "PDFtoTxt"
'Attribute VB_Name = "PDFtoTxt"
Public Function ProcessPDF(fromPDFpath As String, toTXTpath As String) As String
    'Dim pdfPath As String
    'Dim txtPath As String

    ' Specify the paths
    'pdfPath = RFPpath

Debug.Print "1: " & fromPDFpath

    ' Remove previous version of the text file if it exists
    On Error Resume Next
    If Len(Dir(toTXTpath)) > 0 Then
        Kill toTXTpath
Debug.Print "Killed toTXTpath: " & toTXTpath
    End If
    On Error GoTo 0

    ' Use MuPDF to extract text from the PDF
    ProcessPDF = ExtractTextFromPDF(fromPDFpath, toTXTpath)
End Function



Public Function ExtractTextFromPDF(pdfPath As String, txtPath As String) As String
    Dim command As String
    Dim engine As String
    Dim fileContents As String
    Dim fso As Object
    Dim fileExists As Boolean
    Dim startTime As Double
    Dim fileNum As Integer
    Dim line As String
    Dim lastSize As Long
    Dim currentSize As Long
    Dim buffer As String ' Use buffer to avoid large string concatenation issues
    
    ' Path to MuPDF's mutool.exe
    engine = GetRelativePath("\PDFreader\mupdf-1.24.0-windows\mutool.exe")
    'engine = GetRelativePath("\Tesseract-OCR\tesseract.exe")
    

    ' Use mutool to convert PDF to text
    command = engine & " draw -F txt -o """ & txtPath & """ """ & pdfPath & """"
    
    'command = engine & " " & pdfPath & " " & txtPath & " -l eng pdf"
    ' command = engine & " " & pdfPath & " " & txtPath & " -l jpn pdf"
'command = engine & Chr(34) & pdfPath & Chr(34) & " " & Chr(34) & Left(txtPath, InStrRev(txtPath, ".") - 1) & Chr(34) & " -l jpn pdf"
    
    Debug.Print "Command: " & command

    ' Execute the command
    Shell command, vbHide

    ' Wait for the text file to be generated, up to 30 seconds, checking file size stability
    Set fso = CreateObject("Scripting.FileSystemObject")
    startTime = Timer
    lastSize = 0
    
    Do
        ' Check if the file exists
        If fso.fileExists(txtPath) Then
            ' Get the current file size
            currentSize = FileLen(txtPath)
            
            ' Check if the file size has stabilized (i.e., it's no longer increasing)
            If currentSize = lastSize And currentSize > 0 Then
                Exit Do  ' File size hasn't changed, assume it's done writing
            End If
            
            ' Update lastSize for the next comparison
            lastSize = currentSize
        End If
        
        Debug.Print "Waiting for file size to stabilize... Current size: " & currentSize
        Application.Wait Now + TimeValue("00:00:01")  ' Wait 1 second
        
        ' Timeout after 30 seconds
        If Timer - startTime > 30 Then
            MsgBox "Timed out waiting for the PDF to be converted.", vbExclamation
            Exit Function
        End If
    Loop

    ' Read the extracted text file line by line
    fileNum = FreeFile
    On Error GoTo ErrorHandler
    
    Debug.Print "3 txtPath: " & txtPath

    ' Open file
    Open txtPath For Input As #fileNum

    ' Initialize file contents
    buffer = ""  ' Reset the buffer for storing chunks of data
    fileContents = ""  ' Reset file contents

    ' Read file line by line and concatenate to buffer
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        
        ' Check if the buffer is growing too large; if so, concatenate to fileContents
        If Len(buffer) > 10000 Then  ' Arbitrary limit to prevent overflow
            fileContents = fileContents & buffer
            buffer = ""  ' Reset buffer after writing to fileContents
        End If
        
        buffer = buffer & line & vbCrLf
    Loop

    ' Add any remaining buffer content to fileContents
    fileContents = fileContents & buffer

    Debug.Print "Extracted Text: " & fileContents

    ' Close the file
    Close #fileNum

    ' Return the extracted text
    ExtractTextFromPDF = fileContents
    On Error GoTo 0

    Exit Function

ErrorHandler:
    MsgBox "Error in ExtractTextFromPDF: " & Err.Description
    ExtractTextFromPDF = ""
    If fileNum > 0 Then Close #fileNum

End Function

