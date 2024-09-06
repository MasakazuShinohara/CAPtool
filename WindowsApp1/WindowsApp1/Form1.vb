Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class Form1
    Dim WithEvents captureTimer As New Timer()
    Dim captureInProgress As Boolean = False
    Dim capturedFiles As New List(Of String)

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowRect(hWnd As IntPtr, ByRef rect As RECT) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As Integer, vk As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Private Const MOD_NONE As Integer = &H0
    Private Const VK_ESCAPE As Integer = &H1B

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ホットキーの登録
        RegisterHotKey(Me.Handle, 1, MOD_NONE, VK_ESCAPE)

        ' フォームの最大化と最小化のボタンを非表示にする
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
    End Sub


    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H312 Then ' WM_HOTKEY
            If m.WParam.ToInt32() = 1 Then ' ESCキー
                captureInProgress = False
                captureTimer.Stop()
                StartExcelAndInsertImages()
                Me.Close()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        UnregisterHotKey(Me.Handle, 1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.WindowState = FormWindowState.Minimized
        captureInProgress = True
        captureTimer.Interval = 100
        captureTimer.Start()
    End Sub

    Private Sub captureTimer_Tick(sender As Object, e As EventArgs) Handles captureTimer.Tick
        If captureInProgress AndAlso Control.MouseButtons = MouseButtons.Left Then
            CaptureWindow()
        End If
    End Sub

    Private Sub CaptureWindow()
        Try
            Dim hWnd As IntPtr = GetForegroundWindow()
            Dim rect As New RECT()
            GetWindowRect(hWnd, rect)
            Dim width As Integer = rect.Right - rect.Left
            Dim height As Integer = rect.Bottom - rect.Top

            Dim windowGrab As New Bitmap(width, height)
            Dim g As Graphics = Graphics.FromImage(windowGrab)
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, New Size(width, height))

            ' CheckBox1にチェックが入っている場合のみ、マウスカーソルの位置に赤い円を描画
            If CheckBox1.Checked Then
                Dim cursorPosition As Point = Cursor.Position
                Dim redPen As New Pen(Color.Red, 2)
                g.DrawEllipse(redPen, cursorPosition.X - rect.Left - 10, cursorPosition.Y - rect.Top - 10, 20, 20)
            End If

            ' ファイル名に現在の日時を付与
            Dim fileName As String = Application.StartupPath & "\" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".png"
            windowGrab.Save(fileName, ImageFormat.Png)
            capturedFiles.Add(fileName)
        Catch ex As Exception
            ' エラーメッセージの表示をやめる
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        captureInProgress = False
        captureTimer.Stop()
        StartExcelAndInsertImages()
        Me.Close()
    End Sub

    Private Sub StartExcelAndInsertImages()
        Try
            ' Excelアプリケーションに接続
            Dim excelApp As Excel.Application = Marshal.GetActiveObject("Excel.Application")
            excelApp.Visible = True

            ' 新しいブックを作成
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Excel.Worksheet = workbook.Sheets(1)

            ' 画像をA列に貼り付け
            Dim row As Integer = 1
            For Each file In capturedFiles
                Dim img As Excel.Picture = worksheet.Pictures().Insert(file)
                img.Top = worksheet.Cells(row, 1).Top
                img.Left = worksheet.Cells(row, 1).Left

                ' 画像の高さをセルの高さに換算してrowを更新
                Dim imgHeight As Integer = img.Height
                Dim cellHeight As Double = worksheet.Rows(row).Height
                row += Math.Ceiling(imgHeight / cellHeight) + 1 ' 画像が重ならないように調整

                ' 画像の選択を解除
                img.Select(False)
            Next

            ' ファイル名に現在の日時を付与して保存
            Dim savePath As String = Application.StartupPath & "\CAP" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xlsx"
            workbook.SaveAs(savePath)
        Catch ex As Exception
            ' エラーメッセージの表示をやめる
        End Try
    End Sub


End Class
