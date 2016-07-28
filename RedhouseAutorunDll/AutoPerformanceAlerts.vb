Imports HS.Platform
Imports hsopPlatform
Imports System
Imports System.Collections
Imports System.Data


Public Class AutoPerformanceAlerts : Implements HS.Platform.IAutoRunCode


    Public Sub Deal(ByRef pst As CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As AutoRunCodeData, ByRef datFreq As FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements IAutoRunCode.Deal

        Dim PFDefineDat As PFDefine = New PFDefine()
        SLog.Crucial("运行自动生成提醒邮件-" + DateAndTime.Now().ToString())
        If PFDefineDat.GetPFDefine(pst, PFDefineDat) Then
            SLog.Crucial("运行自动生成提醒邮件：" + DateAndTime.Now().ToString())
            GenAlerts(pst, PFDefineDat, "目标未提交")
            GenAlerts(pst, PFDefineDat, "目标未核准")
        End If

    End Sub
   
    Public Function GenAlerts(ByRef pst As CmsPassport, ByVal PFDefineDat As PFDefine, ByVal strAlertType As String) As Boolean
        Dim resid As Long = 420130498195
        Dim reccount As Long
        Dim lngRecid As Long
        Dim lngPnid As Long
        Dim lngPnid2 As Long
        Dim strSql As String = ""
        Dim ht2alert As Hashtable = New Hashtable()
        Dim dt As DataTable = New DataTable()
        Dim ds As DataSet = New DataSet()
        If strAlertType = "目标未提交" Then
            strSql = "select * from ct420130498195 where C3_420150922019='FY2014' and isnull(C3_420953811304,'N')<>'Y' and (convert(char(8),dateadd(day,5,getdate()),112)) =convert(char(8),C3_422371025108,112) or convert(char(8),dateadd(day,3,getdate()),112) =convert(char(8),C3_422371025108,112) or convert(char(8),dateadd(day,5,getdate()),112) =convert(char(8),C3_422371025108,112) and (convert(char(8),getdate(),112)=convert(char(8),C3_422371025108,112))"
        ElseIf strAlertType = "目标未核准" Then
            strSql = "select * from ct420130498195 where C3_420150922019='FY2014' and isnull(C3_420953811304,'N')='Y' and isnull(C3_420976746773,'N')<>'Y' and dateadd(hour," + PFDefineDat.CheckAlertH.ToString() + ",isnull(C3_421418663228,'1900-1-1')) =getdate()"
        End If

        Try
            dt = CmsTable.GetDatasetForHostTable(pst, resid, False, "", "", strSql, 0, 0, reccount, "", False).Tables(0)
            For i As Long = 0 To dt.Rows.Count - 1
                ht2alert.Clear()
                lngRecid = Convert.ToInt64(dt.Rows(i)("Rec_id"))
                'C3_420148203323 人员编号 
                lngPnid = Convert.ToInt64(dt.Rows(i)("C3_420148203323"))
                lngPnid2 = Convert.ToInt64(dt.Rows(i)("C3_420309049071"))
                ht2alert.Add("C3_421271414457", PFDefineDat.PFYearName)
                ht2alert.Add("C3_421270674475", lngPnid)
                ht2alert.Add("C3_421270900560", lngPnid2)
                ht2alert.Add("C3_420989391756", strAlertType)
                If GenAlerts(pst, PFDefineDat, ht2alert) Then
                    Dim dbs As CmsDbStatement = Nothing
                    If strAlertType = "目标未提交" Then
                        strSql = "update ct420130498195 set C3_421418659958=getdate() where rec_id= " + lngRecid.ToString()
                    ElseIf strAlertType = "目标未核准" Then
                        strSql = "update ct420130498195 set C3_421418663228=getdate() where rec_id= " + lngRecid.ToString()
                    End If
                    CmsDbStatement.Execute(pst.Dbc, strSql, True, Nothing)
                End If
            Next
        Catch ex As Exception
            SLog.Crucial("生成提醒记录错误：" + ex.Message.ToString())
            Return False
        End Try


        'ds.Tables(0)
        'C3_420150922019 财年名称 
        ' C3_420953811304 提交目标 
        'C3_420976746773核准目标
        'C3_421418659958 自动邮件提醒目标未提交日期 日期 8  显示                        
        '88 编辑 C3_421418663228 自动邮件提醒目标未核准日期 

        Return True
    End Function
 
    Public Function GenAlerts(ByRef pst As CmsPassport, ByVal PFDefineDat As PFDefine, ByVal ht As Hashtable) As Boolean
        Dim resid As Long = 420989179080
        Dim cp As CmsTableParam = New CmsTableParam()
        Try
            CmsTable.AddRecord(pst, resid, ht, cp)
        Catch ex As Exception
            SLog.Crucial("添加邮件提醒错误：" + ex.Message.ToString())
            Return False
        End Try

        Return True
    End Function

End Class
'    C3_420161949106 财年 文字 20  显示                        
'2 编辑 C3_420161974487 目标填写开始日期 日期 8  显示                        
'3 编辑 C3_420161981213 目标填写结束日期 日期 8  显示                        
'4 编辑 C3_420162001996 评价开始日期 日期 8  显示                        
'5 编辑 C3_420162012173 评价结束日期 日期 8  显示                        
'6 编辑 C3_420162027612 是否启用 文字 1  显示 是否项(Y/)                      
'7 编辑 C3_421419471144 绩效目标提交邮件提醒间隔 整数 8  显示                        
''8 编辑 C3_421419509494 绩效目标核准邮件提醒间隔 
'420989179080
'C3_421271414457 财年名称 文字 20  显示                        
'2 编辑 C3_421270674475 人员编号 整数 8  显示                        
'3 编辑 C3_421272059001 人员名称 文字 200  显示   计算                    
'4 编辑 C3_421270900560 直评人编号 整数 8  显示                        
'5 编辑 C3_420989220971 目标邮箱 文字 800  显示   计算                    
'6 编辑 C3_420989241735 提醒内容 文字 800  显示   计算                    
'7 编辑 C3_420989391756 邮件提醒类别 文字 20  显示 下拉字典                      
'8 编辑 C3_421270791803 人员邮箱 文字 200  显示   计算                    
'9 编辑 C3_421270800388 主管邮箱 文字 200  显示   计算                    
'10 编辑 C3_421271699717 是否发送 文字 1  显示                        
'11 编辑 C3_421274449870 主管名称 文字 200  显示   计算                    


