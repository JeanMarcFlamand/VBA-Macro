Attribute VB_Name = "PutSpecificFileinWorkbook"
Option Explicit

Sub getFilesandPutinWbk()
    Dim TheSelectedFiles() As String
    Dim TheSelectedFile As Variant
    TheSelectedFiles = GetSelectedFilesDir("I:\Documents\Jean-Marc\sawe canada", "_", "*.xls*")
    
    'looping into the array
    For Each TheSelectedFile In TheSelectedFiles
        ' Will not change the array value
        MsgBox "Selected WB Excel file is: " & TheSelectedFile
    Next TheSelectedFile

End Sub

Public Function GetSelectedFilesDir(ByVal sPath As String, SingleCaracterExclusionCriteria As String, _
                                    Optional ByVal sFilter As String) As String()

    'dynamic array for names
    Dim aFileNames() As String
    ReDim aFileNames(0)
    Dim myPos As Long

    Dim sFile As String
    Dim nCounter As Long

    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    If sFilter = "" Then
        sFilter = "*.*"
    End If

    'call with path "initializes" the dir function and returns the first file
    sFile = Dir(sPath & sFilter)

    'call it until there is no filename returned
    Do While sFile <> ""
    
        'Exclude file that match the Single Caracter Exclusion Criteria
        myPos = InStr(sFile, SingleCaracterExclusionCriteria) ' Returns 0

        If myPos = 0 Then
        
            'store the file name in the array
            aFileNames(nCounter) = sFile
            Debug.Print sFile
            
            'make sure your array is large enough for another
            nCounter = nCounter + 1
            If nCounter > UBound(aFileNames) Then
                'preserve the values and grow by reasonable amount for performance
                ReDim Preserve aFileNames(UBound(aFileNames) + 255)
            End If
        End If
        'subsequent calls without param return next file
        sFile = Dir
 
    Loop

    'truncate the array to correct size
    If nCounter < UBound(aFileNames) Then
        ReDim Preserve aFileNames(0 To nCounter - 1)
    End If

    'return the array of file names
    GetSelectedFilesDir = aFileNames()

End Function


