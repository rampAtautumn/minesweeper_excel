Class Duck
    ' Properties
    Private m_Name As String
    Private m_Age As Integer
    Private m_Color As String

    ' Constructor
    Public Sub New(name As String, age As Integer, color As String)
        m_Name = name
        m_Age = age
        m_Color = color
    End Sub

    ' Methods
    Public Function Quack() As String
        Quack = "Quack!"
    End Function

    Public Function GetDetails() As String
        GetDetails = "Name: " & m_Name & ", Age: " & m_Age & ", Color: " & m_Color
    End Function
End Class