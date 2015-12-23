Imports ESRI.ArcGIS.Catalog

Public Class BA_GxFilterFeatureServers
    Implements IGxObjectFilter

    Public Function CanChooseObject(ByVal catObj As IGxObject, ByRef result As ESRI.ArcGIS.Catalog.esriDoubleClickResult) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanChooseObject
        'Can only choose feature services
        Dim category As String = Trim(catObj.Category)
        If category = BA_EnumDescription(GxFilterCategory.FeatureService) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CanDisplayObject(ByVal Location As ESRI.ArcGIS.Catalog.IGxObject) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanDisplayObject
        'Always show folder connections and ArcGIS Server so you can get to the feature service
        Dim category As String = Trim(Location.Category)
        Select Case category
            Case BA_EnumDescription(GxFilterCategory.FolderConnection)
                Return True
            Case BA_EnumDescription(GxFilterCategory.ArcGisServer)
                Return True
            Case BA_EnumDescription(GxFilterCategory.ArcGisServerFolder)
                Return True
            Case BA_EnumDescription(GxFilterCategory.FeatureService)
                Return True
            Case Else
                Return False
        End Select
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
End Class
