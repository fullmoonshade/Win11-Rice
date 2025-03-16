<#
Run with admin privileges 

This powershell script adds the following hot keys to your Windows machine --> 

Win + C - opens a terminal on Desktop directory 
Win + Shift + C - opens a terminal with admin privileges on Desktop directory 
#>

Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

public class HotKeyManager : Form
{
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
    
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    
    private const int WM_HOTKEY = 0x0312;
    private const int MOD_WIN = 0x0008;
    private const int MOD_SHIFT = 0x0004;
    
    private const int ID_WIN_C = 1;
    private const int ID_WIN_SHIFT_C = 2;
    private readonly string _desktopPath;
    
    public HotKeyManager()
    {
        _desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        
        // Minimize form footprint
        this.WindowState = FormWindowState.Minimized;
        this.ShowInTaskbar = false;
        this.FormBorderStyle = FormBorderStyle.None;
        this.Opacity = 0;
        this.Size = new System.Drawing.Size(1, 1);
        
        // Register hotkeys
        if (!RegisterHotKey(this.Handle, ID_WIN_C, MOD_WIN, (int)Keys.C))
            Console.WriteLine("Failed to register Win+C");
        
        if (!RegisterHotKey(this.Handle, ID_WIN_SHIFT_C, MOD_WIN | MOD_SHIFT, (int)Keys.C))
            Console.WriteLine("Failed to register Win+Shift+C");
    }
    
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY)
        {
            bool isAdmin = m.WParam.ToInt32() == ID_WIN_SHIFT_C;
            LaunchTerminal(isAdmin);
        }
        base.WndProc(ref m);
    }
    
    private void LaunchTerminal(bool admin)
    {
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "wt.exe",
                Arguments = $"-d \"{_desktopPath}\"",
                UseShellExecute = true
            };
            
            if (admin) startInfo.Verb = "runas";
            
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error launching terminal: {ex.Message}");
        }
    }
    
    protected override void Dispose(bool disposing)
    {
        UnregisterHotKey(this.Handle, ID_WIN_C);
        UnregisterHotKey(this.Handle, ID_WIN_SHIFT_C);
        base.Dispose(disposing);
    }
}
"@ -ReferencedAssemblies "System.Windows.Forms"

# Create and run the form
$form = New-Object HotKeyManager
[System.Windows.Forms.Application]::Run($form)