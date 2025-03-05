Sub Main()

    Dim parser As New Parser
    'set attribute path
    parser.Init "C:\path\to\your\file.pdb"
    'load all data
    parser.LoadPDB
    ' Filter the atoms from the data
    parser.FindAtoms
    ' Write the filtered data to an Excel sheet
    Parser.WriteDataToSheet
    
End Sub
