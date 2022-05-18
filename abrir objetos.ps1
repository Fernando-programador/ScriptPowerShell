#comandos a objetos

$wshell = New-Object -COMObject WScript.Shell

$wshell | Get-Member
#ver os comandos possiveis

# $wshell.Popup("texto na tela")

$wshell.Run("Outlook")
$wshell.Run("Notepad")
$wshell.AppActivate("Outlook")
Start-Sleep 3
# sendKeys ("escreve um texto no programa")

