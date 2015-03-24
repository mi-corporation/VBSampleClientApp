Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        InitMiFormsComponent()

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MiFormsComponent1 As MiCo.MiForms.MiFormsComponent
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MiFormsComponent1 = New MiCo.MiForms.MiFormsComponent
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2})
        Me.MenuItem1.Text = "&File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "&Load Form"
        '
        'MiFormsComponent1
        '
        Me.MiFormsComponent1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MiFormsComponent1.Location = New System.Drawing.Point(0, 0)
        Me.MiFormsComponent1.Name = "MiFormsComponent1"
        Me.MiFormsComponent1.Size = New System.Drawing.Size(292, 266)
        Me.MiFormsComponent1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.MiFormsComponent1)
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "VB Sample"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub InitMiFormsComponent()
        MiFormsComponent1.RecognitionResourcePath = "c:\program files (x86)\mi-co\mi-forms sdk\res"
        MiFormsComponent1.LicensePath = MiCo.MiForms.License.DefaultLicensePath()
        MiFormsComponent1.Init()
    End Sub
    Private Sub SetCanvasSize(ByVal fWidth As Double, ByVal fHeight As Double)
        Dim nBorderH As Integer = Me.Size.Height - MiFormsComponent1.Size.Height
        Dim nBorderW As Integer = Me.Size.Width - MiFormsComponent1.Size.Width

        Dim nWidth As Integer = (CInt(fWidth) * 72.0 + CDbl(nBorderH))
        Dim nHeight As Integer = (CInt(fHeight) * 72 + CDbl(nBorderW))

        Me.Size = New System.Drawing.Size(nWidth, nHeight)
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim xloadedform As MiCo.MiForms.LoadedForm

        OpenFileDialog1.Filter = "Mi-Forms Templates (*.mft)|*.mft"
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            xloadedform = MiFormsComponent1.LoadFormFromFile(OpenFileDialog1.FileName)

            If (xloadedform Is Nothing) Then
                Return
            End If

            ' set the size of the component appropriately to display the page at 72 dpi
            Dim xpage As MiCo.MiForms.FormPage = CType(xloadedform.Form.Pages(0), MiCo.MiForms.FormPage)
            SetCanvasSize(xpage.Size.Width / 2.54, xpage.Size.Height / 2.54)

            MiFormsComponent1.ActiveForm = xloadedform
        End If
    End Sub
End Class
