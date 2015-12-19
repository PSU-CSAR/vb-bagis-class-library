Imports ESRI.ArcGIS.Catalog

Public Class BA_GxFilterFeatureServers
    Implements IGxObjectFilter

    Public Function CanChooseObject(ByVal catObj As IGxObject, ByRef result As ESRI.ArcGIS.Catalog.esriDoubleClickResult) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanChooseObject
        Dim canChoose As Boolean = False
        If InStr(catObj.Name, ".") <> 0 Then
            Dim ext As String = GetFileExt(catObj.Name)
            If ext = "FeatureServer" Then
                canChoose = True
            End If
        End If
        If ((catObj.Category = "Folder Connection")) Then
            Return False
        ElseIf (canChoose = True) Then
            Return True
        End If
        Return False
    End Function

    Public Function CanDisplayObject(ByVal Location As ESRI.ArcGIS.Catalog.IGxObject) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanDisplayObject
        'Always show folder connections so you can get to the feature service
        If ((Location.Category = "Folder Connection")) Then
            Return True
        End If
        Dim canDisplay As Boolean = False
        Dim ext As String
        If InStr(Location.Name, ".") <> 0 Then
            ext = GetFileExt(Location.Name)
            'Debug.Print(Location.Name)
            If ext = "FeatureServer" Then
                canDisplay = True
            End If
        Else
            canDisplay = True
        End If
        Return canDisplay
    End Function

    Public Function CanSaveObject(Location As ESRI.ArcGIS.Catalog.IGxObject, newObjectName As String, ByRef objectAlreadyExists As Boolean) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanSaveObject
        Return True
    End Function

    Public ReadOnly Property Description As String Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.Description
        Get
            Return "Feature Services (*.FeatureServer)"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.Name
        Get
            Return "FeatureServer"
        End Get
    End Property

    Function GetFileExt(ByVal strFileWPath As String)

        Try
            Return Right(strFileWPath, Len(strFileWPath) - InStrRev(strFileWPath, "."))
        Catch ex As Exception
            MsgBox("The following error has occured." & vbCrLf & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Source: GetFileExt" & vbCrLf & _
                "Error Description: " & Err.Description, _
                vbCritical, "An Error has Occured!")
            Return Nothing
        End Try
    End Function
End Class
