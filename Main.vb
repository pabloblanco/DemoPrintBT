Imports System
Imports System.Data
Imports System.Windows.Forms
Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices

Public Class Main
    Public sound As New Sound("\Windows\alarm3.wav")
    Dim itemCounter As Integer = 0
    Dim cvsPath As String = Nothing
    Dim intReturnValue As Integer
    Public flagPrint As Boolean = False         ' Bandera para indicar que ya se imprimio y abrio el puerto


    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call

    End Sub

    ' Boton salir
    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click

        Application.Exit()

    End Sub

    Sub SendToSerial(ByVal filename As String)
        Dim p As System.IO.Ports.SerialPort
        Dim fs As IO.StreamReader
        Dim portname As String = TextBox2.Text
        Dim sendValue As String = ""

        Label4.Text = "Imprimiendo..."
        Me.Refresh()

        Try
            p = New System.IO.Ports.SerialPort(portname)
            p.WriteTimeout = 500
            p.ReadTimeout = 500

            p.BaudRate = 9600
            p.DataBits = 8
            p.Parity = Ports.Parity.None
            p.StopBits = Ports.StopBits.One
            'p.Handshake = Ports.Handshake.None      ' Try None. 

            Try
                p.Open()
                ' Comandos de la impresoras

                Dim CharHeight(2) As Byte
                CharHeight(0) = &H1B ' ESC
                CharHeight(1) = &H56 ' !   
                CharHeight(2) = &H1 ' Alto Ancho
                p.Write(CharHeight, 0, 3)

                'Dim CharWidht(2) As Byte
                'CharWidht(0) = &H1B ' ESC
                'CharWidht(1) = &H55 ' !   
                'CharWidht(2) = &H1 ' Alto Ancho
                'p.Write(CharWidht, 0, 3)

                'Dim CharHeig_Width(2) As Byte
                'CharHeig_Width(0) = &H1B ' ESC
                'CharHeig_Width(1) = &H57 ' !   
                'CharHeig_Width(2) = &H2 ' Alto Ancho
                'p.Write(CharHeig_Width, 0, 3)

                'Dim CharSpace(2) As Byte
                'CharSpace(0) = &H1B ' GS 29
                'CharSpace(1) = &H70 ' ! 33  
                'CharSpace(2) = &H0 ' Alto Ancho
                'p.Write(CharSpace, 0, 3)

                fs = New IO.StreamReader(filename)
                Do Until fs.EndOfStream
                    sendValue = fs.ReadLine
                    p.WriteLine(sendValue)
                    Thread.Sleep(100)
                Loop

                fs.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                fs = Nothing
                Thread.Sleep(100)
                If p.IsOpen Then p.Close()
                Thread.Sleep(1000)
                p.Dispose()
                p = Nothing
            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Label4.Text = "Fin Impresion"
        Me.Refresh()
        p = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        SendToSerial(TextBox1.Text)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog1.FileName = TextBox1.Text
        If Me.OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub

End Class
