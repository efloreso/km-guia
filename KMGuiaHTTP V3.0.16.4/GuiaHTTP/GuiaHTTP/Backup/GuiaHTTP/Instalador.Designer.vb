<System.ComponentModel.RunInstaller(True)> Partial Class Instalador
    Inherits System.Configuration.Install.Installer

    'Installer reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    'Se puede modificar usando el Diseñador de componentes.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ProcesoInstalador = New System.ServiceProcess.ServiceProcessInstaller
        Me.ServicioInstalador = New System.ServiceProcess.ServiceInstaller
        '
        'ProcesoInstalador
        '
        Me.ProcesoInstalador.Account = System.ServiceProcess.ServiceAccount.LocalSystem
        Me.ProcesoInstalador.Password = Nothing
        Me.ProcesoInstalador.Username = Nothing
        '
        'ServicioInstalador
        '
        Me.ServicioInstalador.Description = "Recibe información Transaccional"
        Me.ServicioInstalador.DisplayName = "Key Monitor Guia en Linea"
        Me.ServicioInstalador.ServiceName = "Key Monitor Guia en Linea"
        Me.ServicioInstalador.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        '
        'Instalador
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ProcesoInstalador, Me.ServicioInstalador})

    End Sub
    Friend WithEvents ProcesoInstalador As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServicioInstalador As System.ServiceProcess.ServiceInstaller

End Class

'Partial Class Instalador
'    Inherits System.ComponentModel.Component

'    <System.Diagnostics.DebuggerNonUserCode()> _
'    Public Sub New(ByVal container As System.ComponentModel.IContainer)
'        MyClass.New()

'        'Requerido para la compatibilidad con el Diseñador de composiciones de clases Windows.Forms
'        If (container IsNot Nothing) Then
'            container.Add(Me)
'        End If

'    End Sub


'    <System.Diagnostics.DebuggerNonUserCode()> _
'  Public Sub New()
'        MyBase.New()

'        'El Diseñador de componentes requiere esta llamada.
'        InitializeComponent()

'    End Sub
'    'Component reemplaza a Dispose para limpiar la lista de componentes.
'    <System.Diagnostics.DebuggerNonUserCode()> _
'    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
'        Try
'            If disposing AndAlso components IsNot Nothing Then
'                components.Dispose()
'            End If
'        Finally
'            MyBase.Dispose(disposing)
'        End Try
'    End Sub

'    'Requerido por el Diseñador de componentes
'    Private components As System.ComponentModel.IContainer

'    'NOTA: el Diseñador de componentes requiere el siguiente procedimiento
'    'Se puede modificar usando el Diseñador de componentes.
'    'No lo modifique con el editor de código.
'    <System.Diagnostics.DebuggerStepThrough()> _
'    Private Sub InitializeComponent()
'        Me.ProcesoInstalador = New System.ServiceProcess.ServiceProcessInstaller
'        Me.ServicioInstalador = New System.ServiceProcess.ServiceInstaller
'        '
'        'ProcesoInstalador
'        '
'        Me.ProcesoInstalador.Account = System.ServiceProcess.ServiceAccount.LocalSystem
'        Me.ProcesoInstalador.Password = Nothing
'        Me.ProcesoInstalador.Username = Nothing
'        '
'        'ServicioInstalador
'        '
'        Me.ServicioInstalador.Description = "Guia para key Monitor"
'        Me.ServicioInstalador.DisplayName = "Guia HTTP"
'        Me.ServicioInstalador.ServiceName = "Guia HTTP"
'        Me.ServicioInstalador.StartType = System.ServiceProcess.ServiceStartMode.Automatic

'    End Sub
'    Friend WithEvents ProcesoInstalador As System.ServiceProcess.ServiceProcessInstaller
'    Friend WithEvents ServicioInstalador As System.ServiceProcess.ServiceInstaller

'End Class
