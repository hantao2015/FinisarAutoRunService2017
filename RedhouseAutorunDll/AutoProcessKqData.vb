Imports HS.Platform
Imports hsopPlatform
Imports System
Imports System.Collections
Imports System.Data
Imports System.Threading
Imports FinisarAutorun
Public Class AutoProcessKqData : Implements HS.Platform.IAutoRunCode
    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        SLog.Crucial("运行自动考勤数据分析-" + DateAndTime.Now().ToString())
        AutoProcessKqData(pst)
    End Sub
    Public Sub AutoProcessKqData(ByRef pst As HS.Platform.CmsPassport)
        Dim dtStart As DateTime = DateAndTime.Now
        Dim dtEnd As DateTime = DateAndTime.Now
        Dim strProc As String = "[dbo].[Pro_attdailyrpt_process_withoutinitial]  "
        Try
            Dim hs As Hashtable = New Hashtable
            If Not CmsTable.IsRecordExist(pst, 424358078333, "C3_424358188666", "Y") Then
                SLog.Crucial("当前考勤期间不存在无法自动数据分析-" + DateAndTime.Now().ToString())
                Return

            End If
            hs = CmsTable.GetRecordHashtableByUniqueColumn(pst, 424358078333, "C3_424358188666", "Y", False)
            dtStart = Convert.ToDateTime(hs("C3_424358162925"))
            AutoProcessKqData(pst, dtStart, dtEnd, strProc, "true")
        Catch ex As Exception

            SLog.Err("AutoProcessKqData", ex)
        End Try
     
    End Sub
    Public Sub AutoProcessKqData(ByRef pst As HS.Platform.CmsPassport, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal strProc As String, ByVal skipnormalrpt As String)
        Dim rp As New CmsTableParam
        Dim lngResid As Long = 227186227531
        Dim rtncmstable As New CmsTableReturn
        Dim ds As New DataSet

        Dim dt As New DataTable
        Dim strSQl As String = "select * from ct227186227531 "
        Dim lngReccount As Long
        Dim strErr As String = ""
        Dim strExcSql As String = ""
        Dim strPnid As String
        '查询当前人员档案的人员

        Try
            dt = CmsTable.GetDatasetForHostTable(pst, lngResid, False, "", "", strSQl, 0, 0, lngReccount, "", True).Tables(0)
            For i As Long = 0 To dt.Rows.Count - 1
                Thread.Sleep(100)
                Try
                    strPnid = Convert.ToString(dt.Rows(i)("C3_305737857578"))
                    strExcSql = strProc + " " + strPnid + ",'" + dtStart.ToString() + "','" + dtEnd.ToString() + "','0','false','true','" + skipnormalrpt + "',''"
                    CmsDbStatement.Execute(pst.Dbc, strExcSql, True, Nothing)
                Catch ex As Exception
                    SLog.Err("AutoProcessKqData2-数据处理" + ex.Message + "strSQl=" + strSQl)
                End Try


            Next
        Catch ex As Exception
            SLog.Err("AutoProcessKqData1-查询员工档案失败" + ex.Message + "strSQl=" + strSQl)
        End Try
    End Sub
End Class
Public Class AutoProcessKqDataSynProcess : Implements HS.Platform.IAutoRunCode
    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        SLog.Crucial("运行考勤数据分析同步处理" + DateAndTime.Now().ToString())
        KqDataSynProcess(pst)
    End Sub
    Public Sub KqDataSynProcess(ByRef pst As HS.Platform.CmsPassport)
        Try

            Dim strSql = "select * from CT425176359095 where isnull(C3_425176423317,'N') <> 'Y'"
            Dim intCount As Integer = 0
            Dim lngRecID As Long = 0
            Dim strProce As String = "EXECUTE [dbo].[pro_attprocesskqdata_syn]   "
            Dim strExcuteSql = ""

            Dim ds As DataSet = CmsTable.GetDatasetForHostTable(pst, 425176359095, False, "", "", strSql, 0, 0, intCount, "", False)
            For i As Integer = 0 To ds.Tables(0).DefaultView.Count - 1
                Thread.Sleep(500)
                Try
                    lngRecID = Convert.ToInt64(ds.Tables(0).DefaultView.Item(i)("REC_ID"))
                    strExcuteSql = strProce + " " + lngRecID.ToString()
                    CmsDbStatement.Execute(pst.Dbc, strExcuteSql, True, Nothing)
                Catch ex As Exception
                    SLog.Err("KqDataSynProcess2-" + ex.Message.ToString(), ex, False)
                End Try

            Next



        Catch ex As Exception
            SLog.Err("KqDataSynProcess1-" + ex.Message.ToString(), ex, False)
        End Try
    End Sub
End Class

Public Class AutoProcessKqDataASynProcess : Implements HS.Platform.IAutoRunCode
    Public apst As CmsPassport
    Public Delegate Sub del_startThreads()
    Public AmtaBatch As Integer = 5
    Public Sub startThreads_complete()
        If CmsParameter.GetParamOfBool(apst.Dbc, "debugmode") Then
            SLog.Crucial("运行数据结算异步处理-ver3.01-end" + DateAndTime.Now().ToString())
        End If
        dealstart()
    End Sub
    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        SLog.Crucial("运行数据结算异步处理-ver3.01-Deal" + DateAndTime.Now().ToString())
        apst = pst
        dealstart()
    End Sub
    Public Sub dealstart()
        If CmsParameter.GetParamOfBool(apst.Dbc, "debugmode") Then
            SLog.Crucial("运行数据结算异步处理-ver3.01-starting" + DateAndTime.Now().ToString())
        End If

        Dim run As del_startThreads
        Dim Result As Long = 0
        Thread.Sleep(5000)
        run = New del_startThreads(AddressOf startThreads)
        run.BeginInvoke(AddressOf startThreads_complete, Result)
    End Sub
    Public Sub startThreads()
        Dim athread1 As Thread = New Thread(New ThreadStart(AddressOf KqDataASynProcessTread))
        Dim athread2 As Thread = New Thread(New ThreadStart(AddressOf KqMonthlyCalASynProcessTread))
        Dim athread3 As Thread = New Thread(New ThreadStart(AddressOf SalaryMonthlyCalASynProcessTread))

        Try
            athread1.IsBackground = True
            athread1.Start()
        Catch ex As Exception
            SLog.Err("运行数据结算异步处理1-" + ex.Message.ToString(), ex, False)
        End Try
        Try
            athread2.IsBackground = True
            athread2.Start()
        Catch ex As Exception
            SLog.Err("运行数据结算异步处理2-" + ex.Message.ToString(), ex, False)
        End Try
        Try
            athread3.IsBackground = True
            athread3.Start()
        Catch ex As Exception
            SLog.Err("运行数据结算异步处理3-" + ex.Message.ToString(), ex, False)
        End Try
        athread1.Join()
        athread2.Join()
        athread3.Join()
    End Sub
    Public Sub KqDataASynProcessTread()
        KqDataASynProcess(apst)
    End Sub
    Public Sub KqMonthlyCalASynProcessTread()
        Dim aKqDataASynProcessformonthlyrpt As KqDataASynProcessformonthlyrpt = New KqDataASynProcessformonthlyrpt()
        aKqDataASynProcessformonthlyrpt.KqDataASynProcessformonthlyrpt(apst)

    End Sub
    Public Sub SalaryMonthlyCalASynProcessTread()
        Dim aKqDataASynProcessformonthlyrpt As KqDataASynProcessformonthlyrpt = New KqDataASynProcessformonthlyrpt()
        aKqDataASynProcessformonthlyrpt.SalaryMonthlyASynCal(apst)

    End Sub

    Public Sub KqDataASynProcess(ByRef pst As HS.Platform.CmsPassport)
        Try
            If CmsParameter.GetParamOfInt(pst.Dbc, "SYS_DailyKQ_Threads") > 0 Then
                AmtaBatch = CmsParameter.GetParamOfInt(pst.Dbc, "SYS_DailyKQ_Threads")
            End If


            Dim strtop As String = "Top " + AmtaBatch.ToString()

            Dim strSql = "select " + strtop + "   * from CT424442156355 where isnull(C3_439574015461,'N')='Y' and isnull(C3_424442209948,'N') <> 'Y' and isnull(C3_424442209948,'N')<>'P'"
            Dim intCount As Integer = 0
            Dim lngRecID As Long = 0
            Dim lngResid As Long = 424442156355
            Dim hs As Hashtable = New Hashtable()
            Dim strProce As String = "EXECUTE [dbo].[pro_attprocesskqdata] "
            Dim strExcuteSql = ""
            'Dim strWhere As String

            Dim rp As CmsTableParam = New CmsTableParam()



            Dim ds As DataSet = CmsTable.GetDatasetForHostTable(pst, 424442156355, False, "", "", strSql, 0, 0, intCount, "", False)
            Dim aThread(ds.Tables(0).DefaultView.Count) As Thread
            For i As Integer = 0 To ds.Tables(0).DefaultView.Count - 1

                Try
                    Thread.Sleep(500)
                    lngRecID = Convert.ToInt64(ds.Tables(0).DefaultView.Item(i)("REC_ID"))
                    hs.Clear()
                    hs.Add("C3_424442209948", "P")
                    rp.EnableAutoEmailSend = False
                    CmsTable.EditRecord(pst, lngResid, lngRecID, hs, rp)
                    strExcuteSql = strProce + " " + lngRecID.ToString()
                    CmsDbStatement.m_dbc = pst.Dbc
                    CmsDbStatement.m_strSql = strExcuteSql

                    aThread(i) = New Thread(New ThreadStart(AddressOf CmsDbStatement.ThreadExecute))
                    aThread(i).IsBackground = True
                    aThread(i).Start()

                Catch ex As Exception
                    SLog.Err("KqDataASynProcess2-KqDataASynProcess" + ex.Message.ToString(), ex, False)
                End Try

            Next
            ' strWhere = "isnull(C3_424442209948,'N') <> 'Y' "
            ' intCount = CmsDbStatement.Count(pst.Dbc, "CT424442156355", strWhere, Nothing)

        Catch ex As Exception
            SLog.Err("KqDataASynProcess1-" + ex.Message.ToString(), ex, False)
            KqDataASynProcess(pst)
        End Try

    End Sub
   
End Class
Public Class AutoProcessKqMonthlyData : Implements HS.Platform.IAutoRunCode


    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        SLog.Crucial("AutoProcessKqMonthlyData-运行自动考勤月度结算-" + DateAndTime.Now().ToString())
        AutoProcessKqData(pst)
    End Sub
    Public Sub AutoProcessKqData(ByRef pst As HS.Platform.CmsPassport)
        Dim strYearmonth = ""
        Try
            Dim hs As Hashtable = New Hashtable
            If Not CmsTable.IsRecordExist(pst, 424358078333, "C3_424358188666", "Y") Then
                SLog.Crucial("AutoProcessKqMonthlyData-当前考勤期间不存在无法自动结算月报-" + DateAndTime.Now().ToString())
                Return

            End If
            hs = CmsTable.GetRecordHashtableByUniqueColumn(pst, 424358078333, "C3_424358188666", "Y", False)
            strYearmonth = Convert.ToString(hs("C3_424358155202"))
            AutoProcessKqData(pst, strYearmonth)
        Catch ex As Exception

            SLog.Err("AutoProcessKqMonthlyData-AutoProcessKqData", ex)
        End Try


    End Sub
    Public Sub AutoProcessKqData(ByRef pst As HS.Platform.CmsPassport, ByVal yearmonth As String)
        Dim rp As New CmsTableParam
        Dim lngResid As Long = 227186227531
        Dim rtncmstable As New CmsTableReturn
        Dim ds As New DataSet

        Dim dt As New DataTable
        Dim strSQl As String = "select * from ct227186227531 "
        Dim lngReccount As Long
        Dim strErr As String = ""
        Dim strCountSql As String = "select count(*) from Attendance_MonthlyRecord where pnid="
        Dim strPnid As String
        Dim strwhere As String = ""
        Dim lngresid2add As Long = 311025002785
        Dim lngrecid As Long
        Dim hs As Hashtable = New Hashtable()
        Dim hs2save As Hashtable = New Hashtable()

        '查询当前人员档案的人员

        Try
            dt = CmsTable.GetDatasetForHostTable(pst, lngResid, False, "", "", strSQl, 0, 0, lngReccount, "", True).Tables(0)
            For i As Long = 0 To dt.Rows.Count - 1
                Thread.Sleep(100)
                Try
                    strPnid = Convert.ToString(dt.Rows(i)("C3_305737857578"))
                    
                    strCountSql = "select count(*) from Attendance_MonthlyRecord where pnid=" + strPnid + "  and yearmonth=" + yearmonth
                    hs2save.Clear()
                    hs.Clear()
                    If CmsDbStatement.CountBySql(pst.Dbc, strCountSql, Nothing) = 0 Then
                        hs2save.Add("PNID", strPnid)
                        hs2save.Add("YEARMONTH", yearmonth)
                        rp.EnableAutoEmailSend = False
                        CmsTable.AddRecord(pst, lngresid2add, hs2save, rp)
                    Else
                        strwhere = " pnid=" + strPnid + "  and yearmonth=" + yearmonth
                        hs = CmsTable.GetRecordHashtableByUniqueWhere(pst, lngresid2add, strwhere, False)
                        lngrecid = Convert.ToInt64(hs("REC_ID"))
                        rp.EnableAutoEmailSend = False
                        CmsTable.EditRecord(pst, lngresid2add, lngrecid, hs, rp)
                    End If

                Catch ex As Exception
                    SLog.Err("AutoProcessKqMonthlyData2-数据处理" + ex.Message + "strSQl=" + strSQl)
                End Try


            Next
        Catch ex As Exception
            SLog.Err("AutoProcessKqMonthlyData1-查询员工档案失败" + ex.Message + "strSQl=" + strSQl)
        End Try
    End Sub
End Class

Public Class AutoProcessKqMonthlyDataByDaily : Implements HS.Platform.IAutoRunCode

    Private Delegate Sub deleAutoProcessKqData(ByVal pst As CmsPassport, ByVal hs As Hashtable)

    Public apst As CmsPassport
    Public Delegate Sub del_startThreads()
    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        apst = pst
        startThreads()
        SLog.Crucial("AutoProcessKqMonthlyDataByDaily-任务管理-Deal" + DateAndTime.Now().ToString())
    End Sub
    Public Sub startThreads()
        If CmsParameter.GetParamOfBool(apst.Dbc, "debugmode") Then
            SLog.Crucial("AutoProcessKqMonthlyDataByDaily-任务管理-start" + DateAndTime.Now().ToString())
        End If

        'Dim athread1 As Thread = New Thread(New ThreadStart(AddressOf AutoProcessKqData1))
        'athread1.Start()
        Dim run As del_startThreads
        Dim Result As Long = 0
        Thread.Sleep(5000)
        run = New del_startThreads(AddressOf AutoProcessKqData1)
        run.BeginInvoke(AddressOf startThreads_complete, Result)

    End Sub
    Public Sub startThreads_complete()
        If CmsParameter.GetParamOfBool(apst.Dbc, "debugmode") Then
            SLog.Crucial("AutoProcessKqMonthlyDataByDaily-任务管理-complete" + DateAndTime.Now().ToString())
        End If
        startThreads()
    End Sub
    Public Sub AutoProcessKqData1()
        Dim strYearmonth = ""

        Try
            '读取结算任务
            Dim rp As New CmsTableParam
            Dim lngResid As Long = 426607306969
            Dim strSQl As String = "select top 1 * from CT426607306969 where  isnull(C3_427156560908,'')<>'计算完毕' and isnull(C3_427156560908,'')<>'正在添加数据分析任务...'"
            Dim lngReccount As Long
            Dim lngRecid As Long
            Dim dt As New DataTable
            Dim hs As New Hashtable
            Dim Result As Long = 0
            dt = CmsTable.GetDatasetForHostTable(apst, lngResid, False, "", "", strSQl, 0, 0, lngReccount, "", True).Tables(0)
            For i As Integer = 0 To dt.Rows.Count - 1
                lngRecid = Convert.ToInt64(dt.Rows(i)("REC_ID"))
                hs = CmsTable.GetRecordHashtableByRecID(apst, lngResid, lngRecid)
                'Dim run As deleAutoProcessKqData
                'run = New deleAutoProcessKqData(AddressOf AutoProcessKqData1)
                'run.BeginInvoke(apst, hs, AddressOf AutoProcessKqData_complete, Result)
                AutoProcessKqData1(apst, hs)

            Next
        Catch ex As Exception
            SLog.Err("AutoProcessKqMonthlyDataByDaily-AutoProcessKqData", ex)
        End Try



    End Sub
    Public Sub AutoProcessKqData1(ByVal pst As HS.Platform.CmsPassport, ByVal hs As Hashtable)
        Dim rp As New CmsTableParam
        Dim lngResid As Long = 227186227531
        Dim rtncmstable As New CmsTableReturn
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim strSQl As String = "select * from CT429373729010 "
        Dim lngReccount As Long
        Dim strErr As String = ""
        Dim strCountSql As String = "select count(*) from CT425820361991 where C3_425820547208="
        Dim strPnid As String = ""
        Dim strwhere As String = ""
        Dim lngresid2add As Long = 424442156355
        Dim lngrecid As Long
        Dim yearmonth As String
        Dim hs2save As Hashtable = New Hashtable()
        Dim hs2save2 As Hashtable = New Hashtable()
        Dim kqdef As Hashtable = New Hashtable()
        Dim lngresidkqdef As Long = 424358078333
     

        Try
            Dim cp As New CmsTableParam
            cp.UseSpecifiedCrtID = True
            cp.UseSpecifiedEdtID = True
            Dim lngdeptid As Long = Convert.ToInt64(hs("C3_426607433242"))

            yearmonth = Convert.ToString(hs("C3_426609812350"))
            kqdef = CmsTable.GetRecordHashtableByUniqueColumn(pst, lngresidkqdef, "C3_424358155202", yearmonth)

            lngrecid = Convert.ToInt64(hs("REC_ID"))
            'C3_429373825045
            strSQl = "select distinct C3_429373824779,C3_429373825045,C3_429373825263,C3_429374111210 from view_kqprocesslist where   C3_429788235937='Y' and  isnull(C3_429373825045,'')<>'' and  C3_429373875109='" + yearmonth + "' and (C3_429374330280=" + lngdeptid.ToString() + "  or C3_429374330108=" + lngdeptid.ToString() + "  or  C3_429374331639=" + lngdeptid.ToString() + "  or  C3_429374332264=" + lngdeptid.ToString() + " or C3_429374332795=" + lngdeptid.ToString() + ") "
            dt = CmsTable.GetDatasetForHostTable(pst, lngResid, False, "", "", strSQl, 0, 0, lngReccount, "", True).Tables(0)
            hs2save2.Add("C3_427156560908", "正在添加数据分析任务...")
            hs2save2.Add("C3_427156561111", "共" + dt.Rows.Count.ToString() + "人次")
            hs2save2.Add("C3_427156587581", DateAndTime.Now())
            hs2save2.Add("C3_431786850001", dt.Rows.Count.ToString())
            hs2save2.Add("REC_EDTID", hs("REC_EDTID"))
            cp.EnableAutoEmailSend = False
            CmsTable.EditRecord(pst, 426607306969, lngrecid, hs2save2, cp)
            Dim sqlDelete As String = "delete CT424442156355 where C3_431785972216=" + Convert.ToString(hs("C3_431772278147"))
            CmsDbStatement.Execute(pst.Dbc, sqlDelete)
          
            For i As Long = 0 To dt.Rows.Count - 1
                hs2save2("C3_427156561111") = "共" + dt.Rows.Count.ToString() + "人次，当前添加第" + (i + 1).ToString() + "人"

                Thread.Sleep(2)
                Try
                    hs2save.Clear()
                    strPnid = Convert.ToString(dt.Rows(i)("C3_429373824779"))
                    '人员编号
                    hs2save.Add("C3_424442177569", strPnid)
                    '开始日期
                    hs2save.Add("C3_424442188226", kqdef("C3_424358162925"))
                    '结束日期
                    hs2save.Add("C3_424442196223", kqdef("C3_424358171558"))
                    '分析批次
                    hs2save.Add("C3_431785972216", hs("C3_431772278147"))
                    hs2save.Add("C3_424442209948", "N")
                    'hs2save.Add("C3_427765426593", "Y")
                    hs2save.Add("C3_427765443987", "N")
                    '工号
                    hs2save.Add("C3_425176076517", dt.Rows(i)("C3_429373825045"))
                    '姓名
                    hs2save.Add("C3_425176076751", dt.Rows(i)("C3_429373825263"))
                    '所属部门
                    hs2save.Add("C3_425176076892", dt.Rows(i)("C3_429374111210"))
                    '分析月份
                    hs2save.Add("C3_431889009156", yearmonth)
                    '分析申请用户
                    hs2save.Add("REC_CRTID", hs("REC_CRTID"))
                    hs2save.Add("REC_EDTID", hs("REC_EDTID"))
                    '是否日报分析
                    hs2save.Add("C3_439574015461", hs("C3_439572859350"))
                    '是否月报汇总
                    hs2save.Add("C3_427765426593", hs("C3_439572864319"))

                    '是否计算薪资
                    hs2save.Add("C3_439573991099", hs("C3_439572866663"))
                    cp.EnableAutoEmailSend = False
                    CmsTable.AddRecord(pst, lngresid2add, hs2save, cp)
                Catch ex As Exception
                    SLog.Err("AutoProcessKqMonthlyDataByDaily-AutoProcessKqData2-数据处理" + ex.Message + "strSQl=" + strSQl)
                End Try
                cp.EnableAutoEmailSend = False
                CmsTable.EditRecord(pst, 426607306969, lngrecid, hs2save2, cp)
            Next
            hs2save2("C3_427156561111") = "任务添加完毕"
            cp.EnableAutoEmailSend = False
            CmsTable.EditRecord(pst, 426607306969, lngrecid, hs2save2, cp)
            '开始循环检查是否完成
            hs2save2.Clear()
            hs2save2.Add("C3_431787287414", "正在计算...")
            While True

                Thread.Sleep(500)
                cp.EnableAutoEmailSend = False
                CmsTable.EditRecord(pst, 426607306969, lngrecid, hs2save2, cp)
                Dim hsstate As Hashtable = CmsTable.GetRecordHashtableByRecID(pst, 426607306969, lngrecid)
                If hsstate("C3_431786850001") = hsstate("C3_431786866562") Then
                    hs2save2("C3_431787287414") = "计算完毕"
                    Exit While
                End If
                hs2save2("C3_431787287414") = "正在计算..."
            End While
            cp.EnableAutoEmailSend = False
            CmsTable.EditRecord(pst, 426607306969, lngrecid, hs2save2, cp)
        Catch ex As Exception
            SLog.Err("AutoProcessKqMonthlyDataByDaily-AutoProcessKqData2-查询员工档案失败" + ex.Message + "strSQl=" + strSQl)
        End Try
    End Sub
End Class
Public Class KqDataASynProcessformonthlyrpt : Implements HS.Platform.IAutoRunCode
    Dim strYearmonth As String = ""
    Dim pnid As Long = 0
    Dim pstparm As CmsPassport = New CmsPassport()
    Dim hsparm As Hashtable = New Hashtable()
    Dim AmtaBatch As Integer = 5
    Private Delegate Sub delemonthlyrpt(ByVal pst As CmsPassport, ByVal pnid As Long, ByVal yearmonth As String, ByVal hs As Hashtable)
    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        SLog.Crucial("运行考勤数据异步处理-ver2.01" + DateAndTime.Now().ToString())
        KqDataASynProcessformonthlyrpt(pst)
    End Sub
    Public Sub SalaryMonthlyASynCal(ByRef pst As HS.Platform.CmsPassport)
        Try

            Dim hstemp As Hashtable = New Hashtable()
            Try

                If Not CmsTable.IsRecordExist(pst, 424358078333, "C3_424358188666", "Y") Then
                    SLog.Crucial("SalaryMonthlyASynCal-当前考勤期间不存在无法自动结算薪资-" + DateAndTime.Now().ToString())
                    Return

                End If
                hstemp = CmsTable.GetRecordHashtableByUniqueColumn(pst, 424358078333, "C3_424358188666", "Y", False)
                strYearmonth = Convert.ToString(hstemp("C3_424358155202"))

            Catch ex As Exception

                SLog.Err("SalaryMonthlyASynCal", ex)
                Return
            End Try
            If CmsParameter.GetParamOfInt(pst.Dbc, "SYS_MonthlySalary_Threads") > 0 Then
                AmtaBatch = CmsParameter.GetParamOfInt(pst.Dbc, "SYS_MonthlySalary_Threads")
            End If

            Dim strtop As String = "Top " + AmtaBatch.ToString()
            Dim strSql = "select  " + strtop + "  * from CT424442156355 where isnull(C3_424442209948,'N') ='Y' and  isnull(C3_427765443987,'N')='Y' and C3_439573991099='Y' and (isnull(C3_439574424344,'N')<>'P' and isnull(C3_439574424344,'N')<>'Y')  order by C3_424442177569 desc"
            Dim intCount As Integer = 0
            Dim lngRecID As Long = 0
            Dim lngResid As Long = 424442156355
            Dim hs As Hashtable = New Hashtable()

            Dim strExcuteSql = ""

            Dim hashkey As Hashtable = New Hashtable()
            hashkey.Clear()
            Dim rp As CmsTableParam = New CmsTableParam()
           
                Dim ds As DataSet = CmsTable.GetDatasetForHostTable(pst, 424442156355, False, "", "", strSql, 0, 0, intCount, "", False)

                Dim aThread(ds.Tables(0).DefaultView.Count - 1) As Thread

                Dim acalsalary(ds.Tables(0).DefaultView.Count - 1) As calsalary
                For i As Integer = 0 To ds.Tables(0).DefaultView.Count - 1

                    Try


                        lngRecID = Convert.ToInt64(ds.Tables(0).DefaultView.Item(i)("REC_ID"))

                        acalsalary(i) = New calsalary()

                        acalsalary(i).hsparm = CmsTable.GetRecordHashtableByRecID(pst, 424442156355, lngRecID)
                        acalsalary(i).strygno = Convert.ToString(ds.Tables(0).DefaultView.Item(i)("C3_425176076517"))
                        acalsalary(i).pstparm = pst
                        Dim stryearmonth2 As String = ""
                        If ds.Tables(0).DefaultView.Item(i)("C3_431889009156") = Nothing Then
                            stryearmonth2 = ""
                        Else
                            stryearmonth2 = Convert.ToString(ds.Tables(0).DefaultView.Item(i)("C3_431889009156"))
                        End If

                        If stryearmonth2 = "" Then
                            acalsalary(i).strYearmonth = strYearmonth
                        Else
                            acalsalary(i).strYearmonth = stryearmonth2
                        End If

                        hsparm("C3_439574424344") = "P"
                        hsparm("C3_427847943186") = DateAndTime.Now
                        CmsDbStatement.AddOrEditRecord(pst.Dbc, "CT424442156355", hsparm, "REC_ID=" + lngRecID.ToString(), False, True)


                    Catch ex As Exception
                        SLog.Err("SalaryMonthlyASynCal-" + ex.Message.ToString(), ex, False)
                    End Try
                Next

                Dim aThread1(acalsalary.Length) As Thread
                For k As Integer = 0 To acalsalary.Length - 1
                Thread.Sleep(1)
                    aThread1(k) = New Thread(New ThreadStart(AddressOf acalsalary(k).calsalary))
                    aThread1(k).IsBackground = True
                    aThread1(k).Start()
                Next


        Catch ex As CmsException
            SLog.Err("SalaryMonthlyASynCal-" + ex.Message.ToString(), ex, False)
            KqDataASynProcessformonthlyrpt(pst)
        Catch ex As Exception
            SLog.Err("SalaryMonthlyASynCal-" + ex.Message.ToString(), ex, False)
            KqDataASynProcessformonthlyrpt(pst)
        End Try
    End Sub
    Public Sub KqDataASynProcessformonthlyrpt(ByRef pst As HS.Platform.CmsPassport)
        Try

            Dim hstemp As Hashtable = New Hashtable()
            Try

                If Not CmsTable.IsRecordExist(pst, 424358078333, "C3_424358188666", "Y") Then
                    SLog.Crucial("KqDataASynProcessformonthlyrpt-当前考勤期间不存在无法自动结算月报-" + DateAndTime.Now().ToString())
                    Return

                End If
                hstemp = CmsTable.GetRecordHashtableByUniqueColumn(pst, 424358078333, "C3_424358188666", "Y", False)
                strYearmonth = Convert.ToString(hstemp("C3_424358155202"))

            Catch ex As Exception

                SLog.Err("KqDataASynProcessformonthlyrpt-AutoProcessKqData", ex)
                Return
            End Try
            If CmsParameter.GetParamOfInt(pst.Dbc, "SYS_MonthlyKQ_Threads") > 0 Then
                AmtaBatch = CmsParameter.GetParamOfInt(pst.Dbc, "SYS_MonthlyKQ_Threads")
            End If
            Dim strtop As String = "Top " + AmtaBatch.ToString()
            Dim strSql = "select  " + strtop + "  * from CT424442156355 where isnull(C3_424442209948,'N') ='Y' and C3_427765426593='Y' and (isnull(C3_427765443987,'N')<>'P' and isnull(C3_427765443987,'N')<>'Y')  order by C3_424442177569 desc"
            Dim intCount As Integer = 0
            Dim lngRecID As Long = 0
            Dim lngResid As Long = 424442156355
            Dim hs As Hashtable = New Hashtable()

            Dim strExcuteSql = ""

            Dim hashkey As Hashtable = New Hashtable()
            hashkey.Clear()
            Dim rp As CmsTableParam = New CmsTableParam()


            Dim ds As DataSet = CmsTable.GetDatasetForHostTable(pst, 424442156355, False, "", "", strSql, 0, 0, intCount, "", False)

            Dim aThread(ds.Tables(0).DefaultView.Count - 1) As Thread

            Dim acalmonthlyrpt2(ds.Tables(0).DefaultView.Count - 1) As calmonthlyrpt2
            For i As Integer = 0 To ds.Tables(0).DefaultView.Count - 1

                Try


                    lngRecID = Convert.ToInt64(ds.Tables(0).DefaultView.Item(i)("REC_ID"))

                    acalmonthlyrpt2(i) = New calmonthlyrpt2()

                    acalmonthlyrpt2(i).hsparm = CmsTable.GetRecordHashtableByRecID(pst, 424442156355, lngRecID)
                    acalmonthlyrpt2(i).pnid = Convert.ToInt64(ds.Tables(0).DefaultView.Item(i)("C3_424442177569"))
                    acalmonthlyrpt2(i).pstparm = pst
                    Dim stryearmonth2 As String = ""
                    If ds.Tables(0).DefaultView.Item(i)("C3_431889009156") = Nothing Then
                        stryearmonth2 = ""
                    Else
                        stryearmonth2 = Convert.ToString(ds.Tables(0).DefaultView.Item(i)("C3_431889009156"))
                    End If

                    If stryearmonth2 = "" Then
                        acalmonthlyrpt2(i).strYearmonth = strYearmonth
                    Else
                        acalmonthlyrpt2(i).strYearmonth = stryearmonth2
                    End If

                    hsparm("C3_427765443987") = "P"
                    hsparm("C3_427847943186") = DateAndTime.Now
                    CmsDbStatement.AddOrEditRecord(pst.Dbc, "CT424442156355", hsparm, "REC_ID=" + lngRecID.ToString(), False, True)


                Catch ex As Exception
                    SLog.Err("KqDataASynProcess2-KqDataASynProcessformonthlyrpt" + ex.Message.ToString(), ex, False)
                End Try
            Next

            Dim aThread1(acalmonthlyrpt2.Length) As Thread
            For k As Integer = 0 To acalmonthlyrpt2.Length - 1
                Thread.Sleep(100)

                aThread1(k) = New Thread(New ThreadStart(AddressOf acalmonthlyrpt2(k).calmonthlyrpt))
                aThread1(k).IsBackground = True
                aThread1(k).Start()
            Next


        Catch ex As CmsException
            SLog.Err("KqDataASynProcessformonthlyrpt-" + ex.Message.ToString(), ex, False)
            KqDataASynProcessformonthlyrpt(pst)
        Catch ex As Exception
            SLog.Err("KqDataASynProcessformonthlyrpt-" + ex.Message.ToString(), ex, False)
            KqDataASynProcessformonthlyrpt(pst)
        End Try


    End Sub
End Class
Public Class calmonthlyrpt2
    Public strYearmonth As String = ""
    Public pnid As Long = 0
    Public pstparm As CmsPassport = New CmsPassport()
    Public hsparm As Hashtable = New Hashtable()
    Public Sub calmonthlyrpt()
        Dim yearmonth As String = strYearmonth
        Dim strPnid As String = pnid.ToString()
        Dim apstparm As CmsPassport = pstparm
        Dim ahsparm As Hashtable = hsparm
        Dim dt As New DataTable
        Dim strErr As String = ""
        Dim strCountSql As String = "select count(*) from Attendance_MonthlyRecord where pnid="
        Dim strwhere As String = ""
        Dim lngresid2add As Long = 311025002785
        Dim lngrecid As Long
        Dim lngresidTaskTable As Long = 424442156355
        Dim hs As Hashtable = New Hashtable()
        Dim hs2save As Hashtable = New Hashtable()
        Dim rp As CmsTableParam = New CmsTableParam()
        Try
            strCountSql = "select count(*) from Attendance_MonthlyRecord where pnid=" + strPnid + "  and yearmonth=" + yearmonth
            hs2save.Clear()
            hs.Clear()
            If CmsDbStatement.CountBySql(apstparm.Dbc, strCountSql, Nothing) = 0 Then
                hs2save.Add("PNID", strPnid)
                hs2save.Add("YEARMONTH", yearmonth)
                Try
                    rp.EnableAutoEmailSend = False
                    CmsTable.AddRecord(apstparm, lngresid2add, hs2save, rp)
                Catch ex1 As Exception
                    SLog.Err("calmonthlyrpt1-" + ".月报记录人员编号=" + strPnid + ".考勤月份=" + yearmonth, ex1)
                End Try
            Else
                strwhere = " pnid=" + strPnid + "  and yearmonth=" + yearmonth
                hs = CmsTable.GetRecordHashtableByUniqueWhere(apstparm, lngresid2add, strwhere, False)
                lngrecid = Convert.ToInt64(hs("REC_ID"))
                Try
                    rp.EnableAutoEmailSend = False
                    CmsTable.EditRecord(apstparm, lngresid2add, lngrecid, hs, rp)
                Catch ex2 As Exception
                    SLog.Err("calmonthlyrpt2-" + ".月报记录人员编号=" + strPnid + ".考勤月份=" + yearmonth, ex2)
                End Try
            End If
            lngrecid = Convert.ToInt64(ahsparm("REC_ID"))

            ahsparm = CmsTable.GetRecordHashtableByRecID(apstparm, lngresidTaskTable, lngrecid)
            ahsparm("C3_427765443987") = "Y"
            ahsparm("C3_427765456967") = DateAndTime.Now
            Try
                CmsDbStatement.AddOrEditRecord(apstparm.Dbc, "CT424442156355", ahsparm, "REC_ID=" + lngrecid.ToString(), False, True)
            Catch ex3 As Exception
                SLog.Err("calmonthlyrpt3-任务记录表", ex3)
            End Try
        Catch ex4 As Exception

            SLog.Err("calmonthlyrpt4-", ex4)

        End Try
    End Sub
End Class
Public Class calsalary
    Public strYearmonth As String = ""
    Public strygno As String = 0
    Public pstparm As CmsPassport = New CmsPassport()
    Public hsparm As Hashtable = New Hashtable()
    Public Sub calsalary()
        Dim yearmonth As String = strYearmonth
        Dim strPnid As String = strygno
        Dim apstparm As CmsPassport = pstparm
        Dim ahsparm As Hashtable = hsparm
        Dim dt As New DataTable
        Dim strErr As String = ""
        Dim strCountSql As String = "select count(*) from CT433706772782 where C3_433706818283="
        Dim strwhere As String = ""
        Dim lngresid2add As Long = 433706772782
        Dim lngrecid As Long
        Dim lngresidTaskTable As Long = 424442156355
        Dim hs As Hashtable = New Hashtable()
        Dim hs2save As Hashtable = New Hashtable()
        Dim rp As CmsTableParam = New CmsTableParam()
        Try

            strCountSql = "select count(*) from CT433706772782 where C3_433706818455='" + strPnid + "'  and C3_433706819172=" + yearmonth
            hs2save.Clear()
            hs.Clear()
            If CmsDbStatement.CountBySql(apstparm.Dbc, strCountSql, Nothing) = 0 Then
                hs2save.Add("C3_433706818455", strPnid)
                hs2save.Add("C3_433706819172", yearmonth)
                rp.EnableAutoEmailSend = False
                CmsTable.AddRecord(apstparm, lngresid2add, hs2save, rp)
            Else
                strwhere = " C3_433706818455='" + strPnid + "'  and C3_433706819172=" + yearmonth
                hs = CmsTable.GetRecordHashtableByUniqueWhere(apstparm, lngresid2add, strwhere, False)
                lngrecid = Convert.ToInt64(hs("REC_ID"))
                rp.EnableAutoEmailSend = False
                CmsTable.EditRecord(apstparm, lngresid2add, lngrecid, hs, rp)
            End If
            lngrecid = Convert.ToInt64(ahsparm("REC_ID"))

            ahsparm = CmsTable.GetRecordHashtableByRecID(apstparm, lngresidTaskTable, lngrecid)
            ahsparm("C3_439574424344") = "Y"
            ahsparm("C3_427765456967") = DateAndTime.Now
            CmsDbStatement.AddOrEditRecord(apstparm.Dbc, "CT424442156355", ahsparm, "REC_ID=" + lngrecid.ToString(), False, True)
        Catch ex As Exception

            SLog.Err("calsalary-", ex)
        End Try
    End Sub
End Class