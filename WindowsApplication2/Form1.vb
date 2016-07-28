Imports Microsoft.VisualBasic
Imports HS.Platform
Imports hsopPlatform
Imports System.Collections
Imports System.Windows.Forms
Imports FinisarAutorun
Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim pst As CmsPassport = New CmsPassport()
        CmsEnvironment.InitForClientApplication(Application.StartupPath, "tset.log", True)
        pst = CmsPassport.GenerateCmsPassportBySysuser()
        Dim aAutocodeData As AutoRunCodeData = New AutoRunCodeData()
        Dim af As FrequencyData = New FrequencyData()
        Dim alist As ArrayList = New ArrayList()
        Dim AutoGenBTRecord = New FinisarAutorun.AutoPerformanceAlerts()
        Dim autoGenpftable = New FinisarAutorun.AutoGeneratePFTable()
        Dim PFDefineDat As PFDefine = New PFDefine()
        'Dim kqprocess1 As AutoProcessKqDataSynProcess = New AutoProcessKqDataSynProcess()
        'Dim kqprocess2 As KqDataASynProcessformonthlyrpt = New KqDataASynProcessformonthlyrpt()
        'Try
        '    kqprocess2.KqDataASynProcessformonthlyrpt(pst)
        'Catch ex As Exception
        '    kqprocess2.KqDataASynProcessformonthlyrpt(pst)
        'End Try

       


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim pst As CmsPassport = New CmsPassport()
        CmsEnvironment.InitForClientApplication(Application.StartupPath, "test.log", True)
        pst = CmsPassport.GenerateCmsPassportBySysuser()
        Dim kqprocess1 As AutoProcessKqMonthlyDataByDaily = New AutoProcessKqMonthlyDataByDaily()
        kqprocess1.apst = pst
        kqprocess1.startThreads()
        Dim kqprocess2 As AutoProcessKqDataASynProcess = New AutoProcessKqDataASynProcess()
        kqprocess2.apst = pst
        kqprocess2.dealstart()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim pst As CmsPassport = New CmsPassport()
        CmsEnvironment.InitForClientApplication(Application.StartupPath, "tset.log", True)
        pst = CmsPassport.GenerateCmsPassportBySysuser()

        Dim kqprocess2 As AutoProcessKqMonthlyDataByDaily = New AutoProcessKqMonthlyDataByDaily()
        kqprocess2.apst = pst
        kqprocess2.startThreads()

    End Sub

End Class
