Public Class Main

    '实例化Converter类
    Dim objConverter As New Converter
    '声明全局变量
    Dim strLastModify As String

    '程序加载及预处理
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '遍历所有的单位类型，为每个单位类型对应TableLayoutPanel内的TextBox绑定事件处理程序
        For Each unitset As String In [Enum].GetNames(GetType(Converter.UnitSet))
            AddTextBoxTextChangedHandlers("tlp" & unitset)
            AddTextBoxDoubleClickHandlers("tlp" & unitset)
        Next

        txtLocalATM.Text = objConverter.ATM
        txtViscosityDensity.Text = "1000"
    End Sub

    '绑定TextBox的TextChanged事件处理程序
    Private Sub AddTextBoxTextChangedHandlers(ByVal tlpName As String)
        Dim objTextBox As TextBox
        For Each objTextBox In GetTableLayoutPanel(tlpName).Controls.OfType(Of TextBox)
            '为每个TextBox分配objTextBox_TextChanged事件
            AddHandler objTextBox.TextChanged, AddressOf objTextBox_TextChanged
        Next
    End Sub

    '绑定TextBox的DoubleClick事件处理程序
    Private Sub AddTextBoxDoubleClickHandlers(ByVal tlpName As String)
        Dim objTextBox As TextBox
        For Each objTextBox In GetTableLayoutPanel(tlpName).Controls.OfType(Of TextBox)
            '为每个TextBox分配objTextBox_DoubleClick事件
            AddHandler objTextBox.DoubleClick, AddressOf objTextBox_DoubleClick
        Next
    End Sub

    'DoubleClick事件处理程序
    Private Sub objTextBox_DoubleClick(sender As Object, e As EventArgs)
        CType(sender, TextBox).Copy()
    End Sub

    'TextChanged事件处理程序
    Private Sub objTextBox_TextChanged(sender As Object, e As EventArgs)
        Dim objTextBox As TextBox
        Dim objTableLayoutPanel As TableLayoutPanel
        Dim strTableLayoutPanelName As String

        objTextBox = CType(sender, TextBox)

        If InputValidate(objTextBox) Then
            UpdateDisplay(sender)
        End If

        objTableLayoutPanel = CType(objTextBox.Parent, TableLayoutPanel)
        strTableLayoutPanelName = objTableLayoutPanel.Name

        If strTableLayoutPanelName = "tlpDynamicViscosity" Or strTableLayoutPanelName = "tlpKinematicViscosity" Then
            ConvertViscosity(strTableLayoutPanelName)
        End If
    End Sub

    '文本框输入校验，返回布尔值
    Private Function InputValidate(sender As TextBox) As Boolean
        Dim strText As String = sender.Text
        Dim dblValue As Double

        If strText <> "" Then
            Try
                dblValue = CType(strText, Double)
                Return True
            Catch ex As Exception
                strText = strText & "0"
                Try
                    dblValue = CType(strText, Double)
                    Return False
                Catch ex2 As Exception
                    sender.Clear()
                    Return False
                End Try
            End Try
        Else
            Return True
        End If
    End Function

    '获取当前的TableLayoutPanel
    Private Function GetTableLayoutPanel(ByVal tlpName As String) As TableLayoutPanel
        Dim objTableLayoutPanel As TableLayoutPanel
        objTableLayoutPanel = CType(Me.Controls.Find(tlpName, True)(0), TableLayoutPanel)
        Return objTableLayoutPanel
    End Function

    '进行单位转换并更新显示
    Private Sub UpdateDisplay(sender As Object)

        Dim strTableLayoutPanelName As String
        Dim strUnitSet As String
        Dim objTextBox As TextBox
        Dim strTextBoxName As String
        Dim strUnit As String
        Dim strValue As String

        strTableLayoutPanelName = CType(sender, Control).Parent.Name
        strUnitSet = strTableLayoutPanelName.Remove(0, 3)
        strTextBoxName = CType(sender, Control).Name
        strUnit = strTextBoxName.Replace("txt" & strUnitSet & "_", "")
        strValue = CType(sender, Control).Text

        For Each objTextBox In GetTableLayoutPanel(strTableLayoutPanelName).Controls.OfType(Of TextBox)
            If objTextBox IsNot CType(sender, TextBox) Then
                '解除其它控件的objTextBox_TextChanged事件绑定
                RemoveHandler objTextBox.TextChanged, AddressOf objTextBox_TextChanged
                If strValue <> "" Then
                    Dim strTargetUnit As String
                    strTargetUnit = objTextBox.Name.Replace("txt" & strUnitSet & "_", "")
                    objTextBox.Text = objConverter.Convert(strUnit, strValue, strUnitSet, strTargetUnit)
                Else
                    objTextBox.Text = ""
                End If
                '重建其它控件的objTextBox_TextChanged事件绑定
                AddHandler objTextBox.TextChanged, AddressOf objTextBox_TextChanged
            End If
        Next

    End Sub

    '响应大气压设置文本框的TextChanged事件
    Private Sub txtLocalATM_TextChanged(sender As Object, e As EventArgs) Handles txtLocalATM.TextChanged
        Try
            objConverter.ATM = txtLocalATM.Text
        Catch ex As Exception
            MessageBox.Show("此处只能输入数字", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtLocalATM.Text = objConverter.ATM
        End Try
        UpdateDisplay(txtPressure_Pa)
    End Sub

    '进行动力粘度和运动粘度之间的转换
    Private Sub ConvertViscosity(tlpName As String)
        Dim dblDensity As Double
        If txtViscosityDensity.Text <> "" Then
            dblDensity = CType(txtViscosityDensity.Text, Double)
        Else
            Exit Sub
        End If

        If tlpName = "tlpDynamicViscosity" Then
            RemoveHandler txtKinematicViscosity_m2__s.TextChanged, AddressOf objTextBox_TextChanged
            Dim dblDynamicViscosity_Pa_s As Double
            Try
                dblDynamicViscosity_Pa_s = CType(txtDynamicViscosity_Pa_s.Text, Double)
                txtKinematicViscosity_m2__s.Text = CType(dblDynamicViscosity_Pa_s / dblDensity, String)
            Catch ex As Exception
                txtKinematicViscosity_m2__s.Text = ""
            End Try
            UpdateDisplay(txtKinematicViscosity_m2__s)
            strLastModify = "tlpDynamicViscosity"
            AddHandler txtKinematicViscosity_m2__s.TextChanged, AddressOf objTextBox_TextChanged
        Else
            RemoveHandler txtDynamicViscosity_Pa_s.TextChanged, AddressOf objTextBox_TextChanged
            Dim dblKinematicViscosity_m2__s As Double
            Try
                dblKinematicViscosity_m2__s = CType(txtKinematicViscosity_m2__s.Text, Double)
                txtDynamicViscosity_Pa_s.Text = CType(dblKinematicViscosity_m2__s * dblDensity, String)
            Catch ex As Exception
                txtDynamicViscosity_Pa_s.Text = ""
            End Try
            UpdateDisplay(txtDynamicViscosity_Pa_s)
            strLastModify = "tlpKinematicViscosity"
            AddHandler txtDynamicViscosity_Pa_s.TextChanged, AddressOf objTextBox_TextChanged
        End If
    End Sub

    '响应粘度标签页密度文本框的TextChanged事件
    Private Sub txtViscosityDensity_TextChanged(sender As Object, e As EventArgs) Handles txtViscosityDensity.TextChanged
        Dim objTextBox As TextBox
        objTextBox = CType(sender, TextBox)

        If InputValidate(objTextBox) Then
            If strLastModify = "tlpDynamicViscosity" Then
                objTextBox_TextChanged(txtDynamicViscosity_Pa_s, Nothing)
            ElseIf strLastModify = "tlpKinematicViscosity" Then
                objTextBox_TextChanged(txtKinematicViscosity_m2__s, Nothing)
            End If
        End If
    End Sub

    '响应工具栏"退出"的单击事件
    Private Sub tsbExit_Click(sender As Object, e As EventArgs) Handles tsbExit.Click
        Me.Close()
    End Sub

    '打开关于界面
    Private Sub tsbAbout_Click(sender As Object, e As EventArgs) Handles tsbAbout.Click
        Using objAbout As New About
            objAbout.ShowDialog(Me)
        End Using
    End Sub

End Class
