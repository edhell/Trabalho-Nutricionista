Public Class ItemNutricional
    Public id As Integer
    Public descricao As String
    Public quantidade As Double
    Public medida As String
    Public vd As Integer

    Public Sub New()
    End Sub

    Public Sub New(id As Integer, descricao As String)
        Me.id = id
        Me.descricao = descricao
    End Sub

    Public Sub New(id As Integer, descricao As String, quantidade As Double, medida As String, vd As Integer)
        Me.New(id, descricao)
        Me.quantidade = quantidade
        Me.medida = medida
        Me.vd = vd
    End Sub
End Class
