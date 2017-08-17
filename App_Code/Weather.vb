Imports Microsoft.VisualBasic

Class Weather
    Private _year As Integer
    Private _month As Integer
    Private _day As Integer
    Private _srad As Integer
    Private _tmax As Integer
    Private _tmin As Integer
    Private _prcp As Integer

    Property Year() As Integer
        Get
            Return _year
        End Get
        Set(ByVal value As Integer)
            _year = value
        End Set
    End Property

    Property Month() As Integer
        Get
            Return _month
        End Get
        Set(ByVal value As Integer)
            _month = value
        End Set
    End Property
    Property Day() As Integer
        Get
            Return _day
        End Get
        Set(ByVal value As Integer)
            _day = value
        End Set
    End Property

    Property Srad() As Single
        Get
            Return _srad
        End Get
        Set(ByVal value As Single)
            _srad = value
        End Set
    End Property
    Property Tmax() As Single
        Get
            Return _tmax
        End Get
        Set(ByVal value As Single)
            _tmax = value
        End Set
    End Property
    Property Tmin() As Single
        Get
            Return _tmin
        End Get
        Set(ByVal value As Single)
            _tmin = value
        End Set
    End Property
    Property Prcp() As Single
        Get
            Return _prcp
        End Get
        Set(ByVal value As Single)
            _prcp = value
        End Set
    End Property

End Class
