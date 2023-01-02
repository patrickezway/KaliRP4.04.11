Attribute VB_Name = "MBDD"
Option Explicit

Private Function makeQuery(params As Variant) As rdoQuery

    Dim ip As Long
    Dim sql As String
    Dim qry As rdoQuery
    Dim rs As rdoResultset

    If UBound(params) < 0 Then
        DB_getVal = Nothing
        Exit Function
    End If
    
    sql = params(0)
    Set qry = Odbc_Cnx.CreateQuery("", sql)
    For ip = 1 To UBound(params)
        qry(ip - 1) = params(ip)
    Next ip
    
    Set makeQuery = qry
    
End Function

Private Function makeResultset(params As Variant) As rdoResultset

    Dim qry As rdoQuery
    Dim rs As rdoResultset
    
    Set qry = makeQuery(Array(params)(0))
    Set rs = qry.OpenResultset
    Set qry = Nothing
    
    Set makeResultset = rs
    
End Function

Private Function makeDICT(rs As rdoResultset) As Variant

    Dim rc As rdoColumn
    Dim dict As New Scripting.Dictionary
    
'    Debug.Print "rs.rdoColumns.Count", rs.rdoColumns.Count
    For Each rc In rs.rdoColumns
        If IsNull(rs(rc.Name)) Then
            'Debug.Print rc.Name, "isNull"
        End If
        dict.Add rc.Name, IIf(IsNull(rs(rc.Name)), "", (rs(rc.Name)))
    Next
'Debug.Print Join(dict.Items, "|")
    Set makeDICT = dict
    
End Function

Public Function DB_execute(ParamArray params() As Variant) As Variant

    Dim qry As rdoQuery
    
    Set qry = makeQuery(Array(params)(0))
    qry.Execute
    DB_execute = qry.RowsAffected
    Set qry = Nothing
    
End Function

Public Function DB_getVal(ParamArray params() As Variant) As Variant

    Dim rs As rdoResultset
    
    Set rs = makeResultset(Array(params)(0))
    If rs.rdoColumns.Count = 0 Then
        Call LOG_WARN("DB_getVal", "rs.rdoColumns.Count=0")
    End If
    If rs.rdoColumns.Count > 1 Then
        Call LOG_DBG("DB_getVal", "rs.rdoColumns.Count=" & rs.rdoColumns.Count)
    End If
        
    DB_getVal = (rs(0))
    rs.Close
    
End Function

Public Function DB_getDICT(ParamArray params() As Variant) As Variant

    Dim rs As rdoResultset

    Set rs = makeResultset(Array(params)(0))
    Set DB_getDICT = makeDICT(rs)
    rs.Close
    Set rs = Nothing
    
End Function

Public Function DB_getDICTS(ParamArray params() As Variant) As Variant

    Dim rs As rdoResultset
    Dim dicts As New Scripting.Dictionary
    
    Set rs = makeResultset(Array(params)(0))
    While Not rs.EOF
        dicts.Add rs.AbsolutePosition, makeDICT(rs)
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    Set DB_getDICTS = dicts
    
End Function
