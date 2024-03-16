# x, y coordinates for different elements of UI
$coordMythic = @{x = 97; y = 227}
$coordScrollBarBottom = @{x = 1048; y = 664}
$coordOrangeEssence = @{x = 721; y = 583}
$coordForgeButton = @{x = 728; y = 511}
$coordAddToLoot = @{x = 723; y = 672}

# other settings
$extraDelay = 0.5

# getting coordinates of League of Legends window

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class DesktopWindows {
    public struct Rect {
       public int Left { get; set; }
       public int Top { get; set; }
       public int Right { get; set; }
       public int Bottom { get; set; }
    }

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);
}
'@

$processName = 'LeagueClientUx'
$processesByName = [System.Diagnostics.Process]::GetProcessesByName($processName)

if (-not $processesByName)
{
    throw "$processName process not running."
}

foreach($process in $processesByName) {
    if($process.MainWindowHandle -ne 0) {
        $windowRect = [DesktopWindows+Rect]::new()
        $return = [DesktopWindows]::GetWindowRect($process.MainWindowHandle,[ref]$windowRect)
            if($return) {
                $lolWindow = [PSCustomObject]@{ProcessName=$processName; ProcessID=$process.Id; MainWindowHandle=$process.MainWindowHandle; WindowTitle=$process.MainWindowTitle; Top=$windowRect.Top; Left=$windowRect.Left;}
                if ($lolWindow.Left -le 0 -or $lolWindow.Top -le 0)
                {throw 'Incorrect window position - maybe restart League of Legends or move it on main screen.'}
            }
            else
            {
                throw 'Failed to get window rect'
            }
    }
    else
    {
        throw "No MainWindowHandle for $processName"
    }
}

# preparing for clicks
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$signature=@'
[DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@

$SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru

$LMOUSE_DOWN = 0x00000002
$LMOUSE_UP = 0x00000004

function LeftClickOnXY {
    param (
        $coords,
        $postDelay
    )
    $mouseX = $lolWindow.Left + $coords.x
    $mouseY = $lolWindow.Top + $coords.y
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($mouseX, $mouseY)
    Start-Sleep -Seconds 01
    $SendMouseClick::mouse_event($LMOUSE_DOWN, 0, 0, 0, 0);
    $SendMouseClick::mouse_event($LMOUSE_UP, 0, 0, 0, 0); 
    Start-Sleep -Seconds $postDelay
}

[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$amount = [Microsoft.VisualBasic.Interaction]::InputBox("How many Mythic Essences you want to spend on OE?", "How Many")

try {$amount = [int]$amount} catch {throw "Invalid number: $amount"}
if ($amount -lt 1) {throw "Amount should be greater than 0, but is $amount"}

Start-Sleep -Seconds (5 + $extraDelay)

LeftClickOnXY $coordMythic (2 + $extraDelay)

foreach ($i in (1..$amount)) {
    LeftClickOnXY $coordScrollBarBottom (1 + $extraDelay)
    LeftClickOnXY $coordScrollBarBottom (1 + $extraDelay)

    LeftClickOnXY $coordOrangeEssence (1 + $extraDelay)

    LeftClickOnXY $coordForgeButton (2 + $extraDelay)

    LeftClickOnXY $coordAddToLoot (2 + $extraDelay)
}

