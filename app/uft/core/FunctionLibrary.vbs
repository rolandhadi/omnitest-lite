Class FunctionLibrary

Private m_oFSO
Private m_libraries

Public Function load_libraries(sRootDir, sExtension, bRecursive)
    Dim oFolder, oFile, sFExt
    sFExt = LCase(sExtension)
    Set oFolder = m_oFSO.GetFolder(sRootDir)
    For Each oFile In oFolder.Files
        If LCase(m_oFSO.GetExtensionName(oFile.name)) = sFExt Then
            f_list = Split(Replace(oFile.ParentFolder.Path, ROOT_FOLDER & "app\uft\library\", ""), "\")
            array_push m_libraries, Array(oFile.Path, Join(f_list, "\"), Split(oFile.name, ".")(0))
        End If
    Next
    If bRecursive Then
        Dim oSubFolder
        For Each oSubFolder In oFolder.SubFolders
            load_libraries oSubFolder, sExtension, True
        Next
    End If
End Function

Public Function get_libraries()
    get_libraries = m_libraries
End Function

Private Sub Class_Initialize()
    Set m_oFSO = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    Set m_oFSO = Nothing
End Sub

End Class