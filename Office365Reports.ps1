<#
.NOTES
	Name: Office365Reports.ps1
	Author: Agustin Nasillo
    Date: 01/09/2019 - Buenos Aires, Argentina
    Connect with me on Linkedin: https://www.linkedin.com/in/agustin-nasillo/
    Version: 1.00 - Armado del archivo
             1.01 - Conversor de TotalItemSize a GB

    
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.SYNOPSIS
    Automatically downloads an O365 user report
     
.EXAMPLE 
    .\Office365Reports.ps1

.COMPONENT
   STORE

.ROLE
   Support
#>

#COMENZANDO
$Host.UI.RawUI.WindowTitle = "Office365-ReportO365Mailboxes - Diseñado por Agustin Nasillo"
clear
Write-Host ""

$disclaimer = @"
#################################################################################"
# 
# El presente script fue generado por personal de INSSIDE S.R.L. con el
# objecto de automatizar ciertas tareas y funciones realizadas regularmente
# en el área de negocios de IDM. Este reporte contiene información confidencial
# y no debe ser divulgado y/o compartido con ninguna persona ajena al servicio. 
# El mismo posee herramientas de consulta hacía la BBDD de O365, por lo que 
# ninguna modificación es realizada mientras se ejecuta este script. 
# El autor se exime de toda responsabilidad. Usar bajo este conocimiento. 
#
#################################################################################"
"@

Write-Host $disclaimer
Write-Host ""

#PARAMETROS
$Report = @()
$Contador = 0
[validateset("Bytes", "KB", "MB", "GB", "TB")][string]$SizeIn = 'GB'


#OBTENIENDO USUARIOS DEL AZURE AD
Write-Host "Office365-ReportO365Mailboxes - [...] Obteniendo usuarios [...]" -ForegroundColor Red
Write-Host ""
$Users = Get-MsolUser -All | Where-Object {$_.IsLicensed -eq $true} | Select UserPrincipalName

#OBTENIENDO DATA
Write-Host "Office365-ReportO365Mailboxes - [...] Preparando $($Users.Count) usuarios [...] "  -ForegroundColor Red
Write-Host ""
$TotalUsuarios = $Users.Count
$ConversorHS = 3 / 3600

Write-Host (-join ("Office365-ReportO365Mailboxes - [...] Comenzando - Tiempo estimado: ", [math]::Round(($TotalUsuarios * $ConversorHS),2), " hs [...]" )) -ForegroundColor Red
Write-Host ""

ForEach ($User in $Users)
{
    $UPN = $User.UserPrincipalName
    $Azure = Get-MsolUser -UserPrincipalName $UPN
    $Mailboxes = Get-Mailbox -Identity $UPN
    $MailboxStatistics = Get-MailboxStatistics -Identity $UPN

    $Object = [PSCustomObject][ordered] @{
		UserPrincipalName = $UPN;
		DisplayName = $Azure.DisplayName;
        BlockCredential = $Azure.BlockCredential;
        UsageLocation = $Azure.UsageLocation;
        Sincronizando = $Mailboxes.IsDirSynced;
        PrimaryEmailAddress = $Mailboxes.PrimarySmtpAddress;
        AllEmailAddresses = Convert-ExchangeEmail -Emails $Mailboxes.EmailAddresses -Separator ', ' -RemoveDuplicates -RemovePrefix -AddSeparator;
        TotalItemSize = $MailboxStatistics.TotalItemSize.Value;
		"TotalItemSize (GB)" = Convert-ExchangeSize -Size $MailboxStatistics.TotalItemSize -To $SizeIn -Default '' -Precision "2";
		LastLogonTime = $MailboxStatistics.LastLogonTime;
        Licenses = Convert-Office365License -License $Azure.Licenses.AccountSkuID -Separator ', '}
    
        $Contador += 1
        $Actual = $TotalUsuarios - $Contador
        $Report += $Object

        $Host.UI.RawUI.WindowTitle = "Office365-ReportO365Mailboxes - Diseñado por Agustin Nasillo + Progreso: " + [math]::Round( ($Contador * 100 / $TotalUsuarios),2) + "%"
        Write-Host (-join ($Contador, ". $UPN - ", $Azure.DisplayName, " // Restan: $Actual - Estimado [hs]: ", [math]::Round( ($Actual * $ConversorHS),2) ))   
        
}

#ARMADO Y EXPORTACION DEL REPORTE

Write-Host ""
Write-Host "Office365-ReportO365Mailboxes - [...] Exportando reporte [...]" -ForegroundColor Yellow
$ReportExport = "$env:USERPROFILE\Desktop\ReportO365Mailboxes_$((Get-Date).ToString('MM-dd-yyyy_hhmmss')).csv"
Write-Host ""

$Report | Export-Csv -NoTypeInformation -Path $ReportExport -Encoding UTF8 | Format-Table -Wrap -AutoSize

Write-Host "Office365-ReportO365Mailboxes - [...] El reporte fue almacenado en $ReportExport [...]" -ForegroundColor Green

$Host.UI.RawUI.WindowTitle = "(Exchange Online + Compliance Center)"