using namespace System.Windows.Forms
using namespace System.Drawing

function Initialize-Win32Functions {
    $SetForegroundWindowSignature = @'
[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
public static extern bool SetForegroundWindow(IntPtr hWnd);
'@
    $ShowWindowSignature = @'
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@

    $Win32Functions = Add-Type -MemberDefinition $SetForegroundWindowSignature, $ShowWindowSignature -Name "Win32Functions" -Namespace Win32Functions -PassThru

    $script:SetForegroundWindow = $Win32Functions::SetForegroundWindow
    $script:ShowWindow = $Win32Functions::ShowWindow
}
