Public Sub Foo()

    Dim aFoo As Variant
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    Set db = DBEngine(0)(0)
    Set rst = db.OpenRecordset("tblFoo")

    With rst
        .MoveLast
        .MoveFirst
        aFoo = .GetRows(.RecordCount)
    End With

    rst.Close
    db.Close

End Sub