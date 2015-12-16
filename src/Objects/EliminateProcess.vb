Imports System.Xml.Serialization
Imports BAGIS_ClassLibrary

Public Class EliminateProcess

    'Which method will be used for eliminating features
    Dim m_selectionMethod As String
    'We are selecting by hru area
    Dim m_selectByHruArea As Boolean
    Dim m_polyArea As Double
    Dim m_polyAreaUnits As MeasurementUnit
    'We are selecting by percentile
    Dim m_selectByPercentile As Boolean
    Dim m_areaPercent As Double
    Dim m_polygonsEliminated As Long
    'We are selecting by poly area
    Dim m_selectByPolyArea As Boolean


    ' Constructor for selecting by hru or poly area
    Sub New(ByVal selectionMethod As String, ByVal selectByHruArea As Boolean, ByVal polyArea As Double, _
            ByVal polyAreaUnits As MeasurementUnit, ByVal polygonsEliminated As Long, ByVal selectByPolyArea As Boolean)
        m_selectionMethod = selectionMethod
        m_selectByHruArea = selectByHruArea
        m_polyArea = polyArea
        m_polyAreaUnits = polyAreaUnits
        m_polygonsEliminated = polygonsEliminated
        m_selectByPolyArea = selectByPolyArea
    End Sub

    ' Constructor for selecting by percentile
    Sub New(ByVal selectionMethod As String, ByVal selectByPercentile As Boolean, ByVal polyArea As Double, _
            ByVal areaPercent As Double, ByVal polyAreaUnits As MeasurementUnit, ByVal polygonsEliminated As Long)
        m_selectionMethod = selectionMethod
        m_selectByPercentile = selectByPercentile
        m_polyArea = polyArea
        m_areaPercent = areaPercent
        m_polyAreaUnits = polyAreaUnits
        m_polygonsEliminated = polygonsEliminated
    End Sub

    ' Required for de-serialization. Do not use.
    Sub New()
        MyBase.New()
    End Sub

    Property SelectionMethod() As String
        Get
            Return m_selectionMethod
        End Get
        Set(ByVal value As String)
            m_selectionMethod = value
        End Set
    End Property

    Property SelectByHruArea() As Boolean
        Get
            Return m_selectByHruArea
        End Get
        Set(ByVal value As Boolean)
            m_selectByHruArea = value
        End Set
    End Property

    Property PolygonArea() As Double
        Get
            Return m_polyArea
        End Get
        Set(ByVal value As Double)
            m_polyArea = value
        End Set
    End Property

    <XmlIgnore()> Property PolygonAreaUnits() As MeasurementUnit
        Get
            Return m_polyAreaUnits
        End Get
        Set(ByVal value As MeasurementUnit)
            m_polyAreaUnits = value
        End Set
    End Property

    Public Property PolygonAreaUnitsText() As String
        Get
            Return BA_EnumDescription(m_polyAreaUnits)
        End Get
        Set(ByVal value As String)
            m_polyAreaUnits = BA_GetMeasurementUnit(value)
        End Set
    End Property

    Property SelectByPercentile() As Boolean
        Get
            Return m_selectByPercentile
        End Get
        Set(ByVal value As Boolean)
            m_selectByPercentile = value
        End Set
    End Property

    Property AreaPercent() As Double
        Get
            Return m_areaPercent
        End Get
        Set(ByVal value As Double)
            m_areaPercent = value
        End Set
    End Property

    Property PolygonsEliminated() As Long
        Get
            Return m_polygonsEliminated
        End Get
        Set(ByVal value As Long)
            m_polygonsEliminated = value
        End Set
    End Property

    ReadOnly Property AreaInSqKm() As Double
        Get
            Select Case m_polyAreaUnits
                Case MeasurementUnit.SquareKilometers
                    Return m_polyArea
                Case MeasurementUnit.SquareMiles
                    Return m_polyArea / BA_SQKm_To_SQMile
                Case MeasurementUnit.Acres
                    Return m_polyArea / BA_SQKm_To_ACRE
                Case Else
                    Return 0
            End Select
        End Get
    End Property

    Property SelectByPolyArea() As Boolean
        Get
            Return m_selectByPolyArea
        End Get
        Set(ByVal value As Boolean)
            m_selectByPolyArea = value
        End Set
    End Property

End Class
