'Option Explicit
'Sub aa()
    Dim oEncoder, oFile, oFSO 'As Object
    Dim oStream, sSourceFile
    Dim sDest, sFileOut, oEncFile
    
    Set oEncoder = CreateObject("Scripting.Encoder")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oFile = oFSO.GetFile("input.vbs")
    Set oStream = oFile.OpenAsTextStream(1)
    sSourceFile = oStream.ReadAll
    oStream.Close
    
    sDest = oEncoder.EncodeScriptFile(".vbs", sSourceFile, 0, "")
    sFileOut = "output.vbe"
    Set oEncFile = oFSO.CreateTextFile(sFileOut)
    oEncFile.Write sDest
    oEncFile.Close
'End Sub