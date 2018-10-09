$keyboard_object = New-Object -ComObject WScript.Shell
Function Caps_Lock_Off
{
$caps_lock_key_status = [System.Windows.Forms.Control]::IsKeyLocked('CapsLock')
if ($caps_lock_key_status -eq $True)
    {
        $keyboard_object.SendKeys("{CAPSLOCK}")
        Start-Sleep -Seconds 1
    }
    else
    {
     $keyboard_object.SendKeys("{CAPSLOCK}")
        Start-Sleep -Seconds 1
    }
}
Caps_Lock_Off