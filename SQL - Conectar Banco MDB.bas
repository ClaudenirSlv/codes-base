Option Explicit

Function DBConnection

    Dim Arq     as String 'será o arquivo do banco de dados.

    Arq = "caminho do banco de dados"

    DBConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &  Arq & ";Persist Security Info=False;"


End Function

Sub ExecuteSQL

    ' Atribuição de variáveis
    ' -----------------------
    Dim sql         as String
    Dim cnn         as New ADODB.Connection
    Dim RS          as New ADODB.Recordset
    Dim Fd          as ADODB.Field

    Set cnn = New ADODB.Connection

    'Abrir Conexão

    cnn.Open DBConnection

    set RS = New ADODB.Recordset
    sql = RetornaSQL(1)

    RS.Open sql, cnn

    'Verifica se há dados no Recordset
    '---------------------------------
    If RS.EOF = False Then



    End If

    RS.Close
    cnn.Close



End Sub