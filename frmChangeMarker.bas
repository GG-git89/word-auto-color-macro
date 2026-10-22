Attribute VB_Name = "frmChangeMarker"
Attribute VB_Base = "0{A2ABBF37-921A-4C36-B3C6-53BF1C7EBCF3}{3C50197E-A3BF-44EA-A536-9376EE58616D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim RevName As String

    With cmbRevision
        .Clear
        .AddItem "Rev.A - 绿色"
        .AddItem "Rev.B - 黄色"
        .AddItem "Rev.C - 青色"
        .AddItem "Rev.D - 粉红"
        .AddItem "Rev.E - 蓝色"
        .AddItem "Rev.F - 橙色"
        .AddItem "Rev.G - 灰色"
    End With

    RevName = 获取当前版本()

    Select Case RevName
        Case "Rev.A"
            cmbRevision.ListIndex = 0
        Case "Rev.B"
            cmbRevision.ListIndex = 1
        Case "Rev.C"
            cmbRevision.ListIndex = 2
        Case "Rev.D"
            cmbRevision.ListIndex = 3
        Case "Rev.E"
            cmbRevision.ListIndex = 4
        Case "Rev.F"
            cmbRevision.ListIndex = 5
        Case "Rev.G"
            cmbRevision.ListIndex = 6
        Case Else
            cmbRevision.ListIndex = 0
    End Select

    Call 更新当前版本显示
    Call 更新颜色块
    Call 更新修改状态

End Sub


Private Sub cmbRevision_Change()

    Dim RevName As String

    Select Case cmbRevision.ListIndex
        Case 0
            RevName = "Rev.A"
        Case 1
            RevName = "Rev.B"
        Case 2
            RevName = "Rev.C"
        Case 3
            RevName = "Rev.D"
        Case 4
            RevName = "Rev.E"
        Case 5
            RevName = "Rev.F"
        Case 6
            RevName = "Rev.G"
        Case Else
            Exit Sub
    End Select

    Call 设置版本(RevName)

    Call 更新当前版本显示

End Sub


Private Sub 更新当前版本显示()

    Dim RevName As String

    RevName = 获取当前版本()

    If RevName = "" Then
        lblCurrent.Caption = "当前版本：未选择"
    Else
        lblCurrent.Caption = "当前版本：" & RevName
    End If

End Sub


Private Sub cmdStart_Click()

    Call 开始工程修改

    Call 更新修改状态

End Sub


Private Sub cmdFinish_Click()

    Call 完成工程修改

    Call 更新当前版本显示

    Call 更新修改状态

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub
Private Sub 更新颜色块()

    Dim RevName As String

    RevName = 获取当前版本()

    Select Case RevName

        Case "Rev.A"
            lblColor.BackColor = RGB(0, 255, 0)

        Case "Rev.B"
            lblColor.BackColor = RGB(255, 255, 0)

        Case "Rev.C"
            lblColor.BackColor = RGB(0, 255, 255)

        Case "Rev.D"
            lblColor.BackColor = RGB(255, 153, 204)

        Case "Rev.E"
            lblColor.BackColor = RGB(0, 0, 255)

        Case "Rev.F"
            lblColor.BackColor = RGB(255, 0, 0)

        Case "Rev.G"
            lblColor.BackColor = RGB(191, 191, 191)

        Case Else
            lblColor.BackColor = RGB(255, 255, 255)

    End Select

End Sub
Private Sub 更新修改状态()

    If ActiveDocument.TrackRevisions = True Then

        lblStatus.Caption = "状态：修改中"

    Else

        lblStatus.Caption = "状态：未开始修改"

    End If

End Sub
