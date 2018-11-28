Public Class CategoriaProduto
    Public id As Integer
    Public descricao As String

    Public Sub New()
    End Sub

    Public Sub New(id As Integer, descricao As String)
        Me.id = id
        Me.descricao = descricao
    End Sub
End Class
