
Sub Separated_Export()
 
    Dim expopt As StructExportOptions
    Set expopt = CreateStructExportOptions
    expopt.UseColorProfile = False
    Dim expflt As ExportFilter
    
    Dim aNaMe As Integer
    aNaMe = 0
    Dim sNaMe As String
     
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.CreateSelection
   
    
    
    For Each s In OrigSelection
        s.CreateSelection
        sNaMe = "J:\for CAM\exp" + CStr(aNaMe) + ".ai"
        Set expflt = ActiveDocument.ExportEx(sNaMe, cdrAI, cdrSelection, expopt)
        With expflt
            .Version = 2 ' FilterAILib.aiVersion8
            .TextAsCurves = False
            
            .ConvertSpotColors = False
            .UseColorProfile = False
            .SimulateOutlines = False
            .SimulateFills = False
            .IncludePlacedImages = True
            .IncludePreview = False
            .Finish
        End With
        aNaMe = aNaMe + 1
    Next s
    
    OrigSelection.CreateSelection
End Sub



Sub Separated_stretch11()
 
    Dim expopt As StructExportOptions
    Set expopt = CreateStructExportOptions
    expopt.UseColorProfile = False
    Dim expflt As ExportFilter
    
    Dim aNaMe As Integer
    aNaMe = 0
    Dim sNaMe As String
     
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.CreateSelection
   
    
    
    For Each s In OrigSelection
        s.CreateSelection
        ActiveDocument.ReferencePoint = cdrCenter
        s.Stretch 1.1
    Next s
    
    OrigSelection.CreateSelection
End Sub
