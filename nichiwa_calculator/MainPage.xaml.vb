' 空白ページのアイテム テンプレートについては、http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409 を参照してください

Imports System.Globalization
Imports Windows.UI
''' <summary>
''' それ自体で使用できる空白ページまたはフレーム内に移動できる空白ページ。
''' </summary>
Public NotInheritable Class MainPage

    ''' <summary>
    ''' クラス変数
    ''' </summary>
    Private strNum As String             '入力変数
    Private dblValue As Double           '計算用変数
    Private strOpera As String           '演算子変数
    Private previousMode As Mode         '状態変数
    Private dblResultValue As Double     '計算結果変数

    ''' <summary>
    ''' 'クラス定数
    ''' </summary>
    Private Const CON_STR_PLUS As String = "+"
    Private Const CON_STR_MINUS As String = "-"
    Private Const CON_STR_TIME As String = "×"
    Private Const CON_STR_DIVISION As String = "÷"
    Private Const CON_STR_AC As String = "AC"
    Private Const CON_STR_ZERO As String = "0"
    Private Const CON_STR_POINT As String = "."
    Private Const CON_STR_FORMAT As String = "#,##" 'フォーマット

    ''' <summary>
    ''' 列挙型画面状態
    ''' </summary>
    Private Enum Mode
        init 　　 　　'初期状態
        clear　　　   'クリア状態
        digit_noDec   '演算子なし
        digit_withDec '演算子あり
        afterOpera    '演算子直後
        aftercalc     '計算後

    End Enum

    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します

        '初期化
        txtNum.IsEnabled = False
        txtNum.Text = 0
        strNum = String.Empty
        dblValue = 0
        previousMode = Mode.init
    End Sub

    ''' <summary>
    ''' クリアボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnClear_Click(sender As Object, e As RoutedEventArgs) Handles btnClear.Click

        '初期化処理
        strNum = String.Empty
        txtNum.Text = 0
        btnClear.Content = CON_STR_AC
        dblValue = 0
        dblResultValue = 0
        strOpera = String.Empty
        previousMode = Mode.init

    End Sub

    ''' <summary>
    ''' 数字ボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DigitClick(sender As Object, e As RoutedEventArgs) Handles btnZero.Click, btnOne.Click, btnTwo.Click, btnThree.Click, btnFour.Click, btnFive.Click, btnSix.Click, btnSeven.Click, btnEight.Click, btnNine.Click

        Dim clickBtn As Button = CType(sender, Button)

        '画面状態判定
        Select Case previousMode

            Case Mode.init

                If clickBtn.Content <> CON_STR_ZERO Then

                    UpdateDisplayNoDec(clickBtn.Content)
                    previousMode = Mode.digit_noDec

                End If

            Case Mode.digit_noDec
                UpdateDisplayNoDec(clickBtn.Content)
                previousMode = Mode.digit_noDec

            Case Mode.afterOpera
                UpdateDisplayWithDec(clickBtn.Content)
                previousMode = Mode.digit_withDec

            Case Mode.digit_withDec

                UpdateDisplayNoDec(clickBtn.Content)

            Case Mode.aftercalc

                UpdateDisplayNoDec(clickBtn.Content)

                previousMode = Mode.digit_withDec
        End Select


    End Sub

    ''' <summary>
    ''' ポイントボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PointClick(sender As Object, e As RoutedEventArgs) Handles btnPoint.Click

        Dim clickBtn As Button = CType(sender, Button)

        If strNum.Length = 0 Then

            UpdateDisplayWithDec(clickBtn.Content)
            previousMode = Mode.digit_noDec

        Else
            If strNum.Contains(CON_STR_POINT) = False Then
                UpdateDisplayNoDec(clickBtn.Content)
            End If

        End If

    End Sub

    ''' <summary>
    ''' 演算子ボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BinalyClick(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click, btnSubtr.Click, btnMulti.Click, btnDivi.Click

        Dim opClickBtn As Button = CType(sender, Button)

        strOpera = opClickBtn.Content

        '画面状態判定
        Select Case previousMode

            Case Mode.init

                previousMode = Mode.afterOpera

            Case Mode.digit_noDec

                dblValue = Convert.ToDouble(strNum)

                previousMode = Mode.afterOpera

            Case Mode.digit_withDec

                CalValue()

                UpdateDisplayWithDec(dblResultValue)

                dblValue = Convert.ToDouble(strNum)

                previousMode = Mode.afterOpera

            Case Mode.aftercalc

                If Double.TryParse(strNum, dblValue) = True Then
                    dblValue = Convert.ToDouble(strNum)
                Else
                    dblValue = 0
                End If

                previousMode = Mode.afterOpera

        End Select

    End Sub

    ''' <summary>
    ''' イコールボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EqualClick(sender As Object, e As RoutedEventArgs) Handles btnEqual.Click

        If previousMode <> Mode.init Then
            If previousMode = Mode.digit_withDec Or previousMode = Mode.aftercalc Then
                CalValue()
                UpdateDisplayWithDec(dblResultValue)
                previousMode = Mode.aftercalc
            End If
        End If
    End Sub

    ''' <summary>
    ''' パーセントボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ParcentClick(sender As Object, e As RoutedEventArgs) Handles btnParcent.Click

        '画面状態判定
        Select Case previousMode

            Case Mode.digit_noDec
                CalcuPercant()
            Case Mode.aftercalc
                CalcuPercant()
            Case Mode.digit_withDec
                CalcuPercant()
            Case Else

        End Select

        '最大少数桁数チェック
        CheckDecimal()

    End Sub

    Private Sub CalcuPercant()

        strNum = (Convert.ToDouble(strNum) / 100).ToString
        UpdateDisplayWithDec(strNum)

    End Sub

    ''' <summary>
    ''' 画面更新（追加）
    ''' </summary>
    ''' <param name="upKey"></param>
    Private Sub UpdateDisplayNoDec(ByVal upKey As String)

        If strNum.Length < 9 Then
            txtNum.Text = String.Empty
            strNum = String.Concat(strNum, upKey)
            txtNum.Text = strNum.ToString
        End If

    End Sub

    ''' <summary>
    ''' 画面更新（新規）
    ''' </summary>
    ''' <param name="upKey"></param>
    Private Sub UpdateDisplayWithDec(ByVal upKey As String)

        strNum = upKey
        If strNum = CON_STR_POINT Then

            strNum = String.Concat(CON_STR_ZERO, strNum.ToString)
            txtNum.Text = strNum.ToString

        Else
            txtNum.Text = strNum.ToString

        End If


    End Sub

    ''' <summary>
    ''' 計算メソッド
    ''' </summary>
    Private Sub CalValue()

        If previousMode <> Mode.init Then

            Select Case strOpera
                Case CON_STR_PLUS
                    dblResultValue = dblValue + (Convert.ToDouble(strNum))
                Case CON_STR_MINUS
                    dblResultValue = dblValue - (Convert.ToDouble(strNum))
                Case CON_STR_TIME
                    dblResultValue = dblValue * (Convert.ToDouble(strNum))
                Case CON_STR_DIVISION
                    dblResultValue = dblValue / (Convert.ToDouble(strNum))
            End Select

            '最大桁数チェック
            CheckMaxLength()

            '最小桁数チェック
            CheckMinLength()

            '最大少数桁数チェック
            CheckDecimal()

        End If

    End Sub

    ''' <summary>
    ''' 最大表示桁数（整数）チェック
    ''' </summary>
    Private Sub CheckMaxLength()

        If dblResultValue > 999999999 Then
            dblResultValue = 999999999
        End If

    End Sub


    ''' <summary>
    ''' 最小表示桁数（整数）チェック
    ''' </summary>
    Private Sub CheckMinLength()

        If dblResultValue < -999999999 Then
            dblResultValue = -999999999
        End If

    End Sub

    ''' <summary>
    ''' 最大表示桁数（少数）チェック
    ''' </summary>
    Private Sub CheckDecimal()

        Dim dblNb As Double
        dblNb = Math.Ceiling(Math.Log10(dblResultValue))

        Select Case dblNb

            Case 2
                dblResultValue = Math.Round(dblResultValue, 7)
            Case 3
                dblResultValue = Math.Round(dblResultValue, 6)
            Case 4
                dblResultValue = Math.Round(dblResultValue, 5)
            Case 5
                dblResultValue = Math.Round(dblResultValue, 4)
            Case 6
                dblResultValue = Math.Round(dblResultValue, 3)
            Case 7
                dblResultValue = Math.Round(dblResultValue, 2)
            Case 8
                dblResultValue = Math.Round(dblResultValue, 1)
            Case 9
                dblResultValue = Math.Round(dblResultValue, 0)
            Case Else
                dblResultValue = Math.Round(dblResultValue, 8)

        End Select

    End Sub

End Class
