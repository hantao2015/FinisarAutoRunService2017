Imports HS.Platform
Imports hsopPlatform
Imports System
Imports System.Collections
Imports System.Data
Public Class AutoGeneratePFTable : Implements HS.Platform.IAutoRunCode



    Public Sub Deal(ByRef pst As HS.Platform.CmsPassport, strRunTime As String, lngRunTimeRecID As Long, ByRef datAr As HS.Platform.AutoRunCodeData, ByRef datFreq As HS.Platform.FrequencyData, ByRef alistServiceIDOnly As ArrayList) Implements HS.Platform.IAutoRunCode.Deal
        Dim PFDefineDat As PFDefine = New PFDefine()
        SLog.Crucial("运行自动生成绩效表格-" + DateAndTime.Now().ToString())
        If PFDefineDat.GetPFDefine(pst, PFDefineDat) Then
            dealwithPFDefineDat(pst, PFDefineDat)
        End If
    End Sub
    Public Sub dealwithPFDefineDat(ByRef pst As HS.Platform.CmsPassport, ByVal PFDefineDat As PFDefine)
        'If DateTime.Now >= PFDefineDat.PFPlanfillStartdate And DateTime.Now <= PFDefineDat.PFPlanfillenddate Then
        If PFDefineDat.PFISStart = "Y" And PFDefineDat.PFAutoGenerateTable = "Y" Then
            Dim rp As New CmsTableParam
            Dim lngResid As Long = 227186227531
            Dim rtncmstable As New CmsTableReturn
            Dim ds As New DataSet

            Dim dt As New DataTable
            Dim strSQl As String = "select * from ct227186227531 where C3_420125636966='Y'"
            Dim lngReccount As Long
            Dim strErr As String = ""

            Dim strPnid As String
            '查询当前人员档案中补贴发放的人员

            Try
                dt = CmsTable.GetDatasetForHostTable(pst, lngResid, False, "", "", strSQl, 0, 0, lngReccount, "", True).Tables(0)
                For i As Long = 0 To dt.Rows.Count - 1

                    strPnid = Convert.ToString(dt.Rows(i)("C3_305737857578"))
                    GenFpform(pst, strPnid, PFDefineDat)
                Next
            Catch ex As Exception
                SLog.Err("查询员工档案失败" + ex.Message + "strSQl=" + strSQl)
            End Try
        End If
        'End If
    End Sub
    Private Function GenFpform(ByVal pst As CmsPassport, ByVal strPnid As String, ByVal PFDefineDat As PFDefine) As Boolean
        Dim hf As New Hashtable
        Dim bthf As New Hashtable
        Dim lngResid As Long = 420130498195
        Dim strWhere As String = "C3_420150922019='" + PFDefineDat.PFYearName + "' and C3_420148203323=" + strPnid
        Dim rp As New CmsTableParam
        Try
            hf = CmsTable.GetRecordHashtableByUniqueWhere(pst, lngResid, strWhere)
            If hf.Count = 0 Then
                '如果不存在那么添加记录。
                bthf.Add("C3_420150922019", PFDefineDat.PFYearName)
                bthf.Add("C3_420148203323", strPnid)

                CmsTable.AddRecord(pst, lngResid, bthf, rp)
            Else
                Return True
            End If
        Catch ex As Exception
            SLog.Err("生成绩效表格失败" + ex.Message + "strPnid=" + strPnid + "year=" = PFDefineDat.PFYearName)
            Return False
        End Try
        Return True
    End Function
End Class
