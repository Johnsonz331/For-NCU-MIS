Sub ImportAndQueryData_department()
    Dim conn As Object
    Dim rs As Object
    Dim folderPath As String
    Dim sql As String
    Dim i As Integer
    
    ' 1. 設定檔案路徑 (本活頁簿所在的資料夾)
    folderPath = ThisWorkbook.Path & "\"
    
    ' 2. 建立 ADO 連線 (將資料夾視為資料庫，CSV 視為資料表)
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & folderPath & ";" & _
              "Extended Properties=""Text;HDR=YES;FMT=Delimited"""

    ' 3. SQL 查詢
    sql = "SELECT T3.department, SUM(T1.usage_hours) AS Total_Hours " & _
          "FROM ([usage.csv] AS T1 " & _
          "INNER JOIN [students.csv] AS T2 ON T1.student_id = T2.student_id) " & _
          "INNER JOIN [labs.csv] AS T3 ON T2.lab_id = T3.lab_id " & _
          "GROUP BY T3.department " & _
          "ORDER BY SUM(T1.usage_hours) DESC"

    ' 4. 執行查詢
    Set rs = conn.Execute(sql)

    ' 5. 將結果輸出到工作表
    With ThisWorkbook.Sheets(1)
        .Cells.ClearContents ' 清除舊資料
        
        ' 寫入欄位名稱
        For i = 0 To rs.Fields.Count - 1
            .Cells(1, i + 1).Value = rs.Fields(i).Name
        Next i
        
        ' 寫入資料內容
        .Range("A2").CopyFromRecordset rs
    End With

    ' 6. 關閉連線
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    MsgBox "資料查詢完成！", vbInformation
End Sub





