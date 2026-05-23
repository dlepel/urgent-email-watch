<#
.SYNOPSIS
    Renders a Windows toast notification. Self-contained -- no external modules.

.DESCRIPTION
    Show-Toast.ps1 displays a desktop notification using the built-in
    Windows.UI.Notifications WinRT API. If that API is unavailable (older
    Windows builds, or running outside an interactive desktop session), it
    falls back to a simple WPF popup window in the bottom-right corner.

    No BurntToast or any other third-party module is required.

.PARAMETER Title
    The notification title (bold, first line).

.PARAMETER Message
    The notification body text.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File Show-Toast.ps1 `
        -Title "URGENT: Bank alert" -Message "From: My Bank`nLarge transaction"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Title,

    [Parameter(Mandatory = $true)]
    [string]$Message
)

$ErrorActionPreference = 'Stop'

function Show-WpfFallback {
    param([string]$Title, [string]$Message)

    # Simple WPF popup in the bottom-right corner. Auto-dismisses after 12s
    # or on click. Used only when the WinRT toast API is not available.
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    $window = New-Object System.Windows.Window
    $window.WindowStyle = 'None'
    $window.AllowsTransparency = $true
    $window.Background = 'Transparent'
    $window.ShowInTaskbar = $false
    $window.Topmost = $true
    $window.Width = 380
    $window.SizeToContent = 'Height'
    $window.ResizeMode = 'NoResize'

    $screen = [System.Windows.SystemParameters]::WorkArea
    $window.Left = $screen.Right - $window.Width - 16
    $window.Top = $screen.Bottom - 160

    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = New-Object System.Windows.CornerRadius(8)
    $border.Background = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromRgb(28, 28, 30))
    $border.Padding = New-Object System.Windows.Thickness(16)
    $border.BorderThickness = New-Object System.Windows.Thickness(1)
    $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromRgb(80, 80, 84))

    $stack = New-Object System.Windows.Controls.StackPanel

    $titleBlock = New-Object System.Windows.Controls.TextBlock
    $titleBlock.Text = $Title
    $titleBlock.FontSize = 15
    $titleBlock.FontWeight = 'Bold'
    $titleBlock.Foreground = 'White'
    $titleBlock.TextWrapping = 'Wrap'
    $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)

    $msgBlock = New-Object System.Windows.Controls.TextBlock
    $msgBlock.Text = $Message
    $msgBlock.FontSize = 13
    $msgBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromRgb(210, 210, 214))
    $msgBlock.TextWrapping = 'Wrap'

    [void]$stack.Children.Add($titleBlock)
    [void]$stack.Children.Add($msgBlock)
    $border.Child = $stack
    $window.Content = $border

    # Dismiss on click.
    $window.Add_MouseLeftButtonUp({ $window.Close() })

    # Auto-dismiss after 12 seconds.
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(12)
    $timer.Add_Tick({ $timer.Stop(); $window.Close() })
    $timer.Start()

    [void]$window.ShowDialog()
}

try {
    # Preferred path: native Windows toast via the WinRT API.
    [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]

    # PowerShell itself is a registered AppUserModelID, so toasts work without
    # registering a custom app shortcut.
    $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'

    # Escape XML-significant characters in the user-supplied strings.
    $safeTitle = [System.Security.SecurityElement]::Escape($Title)
    $safeMessage = [System.Security.SecurityElement]::Escape($Message)

    $template = @"
<toast scenario="reminder">
  <visual>
    <binding template="ToastGeneric">
      <text>$safeTitle</text>
      <text>$safeMessage</text>
    </binding>
  </visual>
</toast>
"@

    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xml.LoadXml($template)

    $toast = New-Object Windows.UI.Notifications.ToastNotification($xml)
    $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
    $notifier.Show($toast)
}
catch {
    # WinRT toast unavailable -- fall back to the WPF popup window.
    try {
        Show-WpfFallback -Title $Title -Message $Message
    }
    catch {
        Write-Error "Toast failed (WinRT and WPF fallback both errored): $_"
        exit 1
    }
}
