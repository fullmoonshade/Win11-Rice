<#
Run with admin privileges 

Win + C opens a terminal on Desktop directory 
Win + Shift + C opens a terminal with admin privileges on Desktop directory 
#>
Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

public class HotKeyForm : Form
{
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
    
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    
    private const int WM_HOTKEY = 0x0312;
    private const int MOD_WIN = 0x0008;
    private const int MOD_SHIFT = 0x0004;
    
    private int id_win_c = 1;
    private int id_win_shift_c = 2;
    private string desktopPath;
    
    public HotKeyForm()
    {
        this.desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        this.RegisterHotKeys();
        this.InitializeForm();
    }
    
    private void InitializeForm()
    {
        this.WindowState = FormWindowState.Minimized;
        this.ShowInTaskbar = false;
        this.FormBorderStyle = FormBorderStyle.None;
        this.Size = new System.Drawing.Size(0, 0);
    }
    
    private void RegisterHotKeys()
    {
        if (!RegisterHotKey(this.Handle, id_win_c, MOD_WIN, (int)Keys.C))
            MessageBox.Show("Failed to register Win+C");
        if (!RegisterHotKey(this.Handle, id_win_shift_c, MOD_WIN | MOD_SHIFT, (int)Keys.C))
            MessageBox.Show("Failed to register Win+Shift+C");
    }
    
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY)
        {
            switch (m.WParam.ToInt32())
            {
                case 1:
                    LaunchTerminal(false);
                    break;
                case 2:
                    LaunchTerminal(true);
                    break;
            }
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
                Arguments = $"-d \"{desktopPath}\"",
                UseShellExecute = true
            };
            
            if (admin) startInfo.Verb = "runas";
            
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error launching terminal: {ex.Message}");
        }
    }
    
    protected override void Dispose(bool disposing)
    {
        UnregisterHotKey(this.Handle, id_win_c);
        UnregisterHotKey(this.Handle, id_win_shift_c);
        base.Dispose(disposing);
    }
}
"@ -ReferencedAssemblies "System.Windows.Forms"

# Create and run the form
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object HotKeyForm
[System.Windows.Forms.Application]::Run($form)