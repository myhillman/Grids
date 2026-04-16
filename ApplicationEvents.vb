Partial Friend Class MyApplication
    Inherits Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase
    Private Sub MyApplication_Startup(sender As Object,
                                  e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) _
                                  Handles Me.Startup

        InitHttp()   ' runs BEFORE Form1 is created
    End Sub
End Class

