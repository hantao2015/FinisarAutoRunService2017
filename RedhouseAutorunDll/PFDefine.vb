Imports HS.Platform
Imports hsopPlatform
Imports System
Imports System.Collections
Imports System.Data
Public Class PFDefine
    '    C3_420161949106 财年 文字 20  显示                        
    '2 编辑 C3_420161974487 目标填写开始日期 日期 8  显示                        
    '3 编辑 C3_420161981213 目标填写结束日期 日期 8  显示                        
    '4 编辑 C3_420162001996 评价开始日期 日期 8  显示                        
    '5 编辑 C3_420162012173 评价结束日期 日期 8  显示                        
    '6 编辑 C3_420162027612 是否启用 文字 1  显示 是否项(Y/)                      
    '7 编辑 C3_421419471144 绩效目标提交邮件提醒间隔 整数 8  显示                        
    '8 编辑 C3_421419509494 绩效目标核准邮件提醒间隔 
    Public PFYearName As String
    Public PFPlanfillStartdate As Date
    Public PFPlanfillenddate As Date
    Public PFStartdate As Date
    Public PFEnddate As Date
    Public PFISStart As String
    Public submitAlertH As Long
    Public CheckAlertH As Long
    Public PFAutoGenerateTable As String = "N"

    Public Function GetPFDefine(ByRef pst As CmsPassport, ByRef PFDefineData As PFDefine) As Boolean
        Dim ht As Hashtable = New Hashtable()
        Dim resid As Long = 420161931474
        Dim Sql As String = "C3_420162027612='Y'"
        Try
            ht = CmsTable.GetRecordHashtableByUniqueWhere(pst, resid, Sql)
            PFDefineData.PFYearName = Convert.ToString(ht("C3_420161949106"))
            PFDefineData.PFPlanfillStartdate = Convert.ToDateTime(ht("C3_420161974487"))
            PFDefineData.PFPlanfillenddate = Convert.ToDateTime(ht("C3_420161981213"))
            PFDefineData.PFStartdate = Convert.ToDateTime(ht("C3_420162001996"))
            PFDefineData.PFEnddate = Convert.ToDateTime(ht("C3_420162012173"))
            PFDefineData.submitAlertH = Convert.ToInt32(ht("C3_421419471144"))
            PFDefineData.CheckAlertH = Convert.ToInt32(ht("C3_421419509494"))
            PFDefineData.PFISStart = Convert.ToString(ht("C3_420162027612"))
            PFDefineData.PFAutoGenerateTable = Convert.ToString(ht("C3_421697434723"))
        Catch ex As Exception
            SLog.Crucial("读取财年评估定义错误：" + ex.Message.ToString())
            Return False
        End Try

        Return True
    End Function



End Class
