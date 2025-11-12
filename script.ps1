<# OpenCars Setup v1.8
 - Auto-elevate
 - Winget apps (Acrobat, Chrome, WinRAR, TeamViewer+upgrade, AnyDesk con fallback, WireGuard)
 - Office 365 (Word/Excel/PowerPoint) vía ODT oficial
 - Fondo (convierte JFIF→JPG con WPF/WIC) + aplica con SystemParametersInfo
 - Lock screen (desactiva Spotlight y fija imagen)
 - Ricoh PCL6 (EXE oficial) + Quiter (EXE desde tu repo)
 - Logging en C:\ProgramData\OpenCars\setup-logs\
#>

# ===================== Auto-elevate =====================
$u=[Security.Principal.WindowsIdentity]::GetCurrent()
$p=New-Object Security.Principal.WindowsPrincipal($u)
if(-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
  Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
  exit
}

# ===================== CONFIG =====================
$WallpaperUrl   = "https://raw.githubusercontent.com/Fortecar/setup/main/Tumbnail.jfif"
$LockScreenUrl  = "https://raw.githubusercontent.com/Fortecar/setup/main/Tumbnail.jfif"
$QuiterUrl      = "https://raw.githubusercontent.com/Fortecar/setup/main/quiter.exe"
$RicohDriverUrl = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001344/0001344878/V44300/z05587L1f.exe"

# ===================== LOG =====================
$LogDir="C:\ProgramData\OpenCars\setup-logs"
New-Item -Force -ItemType Directory -Path $LogDir | Out-Null
$Log=Join-Path $LogDir ("setup-{0:yyyyMMdd-HHmmss}-{1}.log" -f (Get-Date),$env:COMPUTERNAME)
function Log([string]$m,[string]$lvl="INFO"){ "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date),$lvl,$m | Tee-Object -FilePath $Log -Append }

[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12

# ===================== Helpers =====================
function Get-File([string]$Url,[string]$Name){
  $dst=Join-Path $env:TEMP $Name
  for($i=1;$i -le 3;$i++){
    try{ Log "Descarga ($i/3): $Url"; Invoke-WebRequest -Uri $Url -OutFile $dst -UseBasicParsing -TimeoutSec 300; break }
    catch{ Start-Sleep 2; if($i -eq 3){ Log "Fallo descarga: $Url" "ERROR"; return $null } }
  }
  return $dst
}

function Winget-Ensure(){
  if(-not (Get-Command winget -EA SilentlyContinue)){ Log "No se encontró winget (App Installer)." "ERROR"; exit 1 }
  try{ winget source update | Out-Null } catch {}
}

function Install-App([string]$Id,[string]$Name){
  $exists = winget list -e --id $Id 2>$null | Select-String $Id
  if($exists){ Log "OK (ya instalado): $Name"; return $true }
  Log "Instalando: $Name"
  winget install -e --id $Id --silent --accept-source-agreements --accept-package-agreements --source winget
  Start-Sleep 2
  $ok = (winget list -e --id $Id 2>$null | Select-String $Id)
  if($ok){ Log "OK: $Name"; return $true } else { Log "No quedó instalado: $Name" "WARN"; return $false }
}

function Install-ExeSilent([string]$Exe,[string[]]$Tries=@('/s /v"/qn REBOOT=ReallySuppress"','/s /v"/qn /norestart"','/quiet /norestart','/S','/silent','/verysilent')){
  foreach($a in $Tries){
    try{
      Log "Instalando EXE: $Exe $a"
      Start-Process $Exe -ArgumentList $a -Wait -NoNewWindow
      if($LASTEXITCODE -eq 0){ Log "OK (EXE): $Exe"; return $true }
    } catch {}
  }
  Log "No se pudo instalar EXE en modo silencioso: $Exe" "WARN"
  return $false
}

# ===== Conversión JFIF/lo que sea → JPG (WPF/WIC) + Aplicar Wallpaper por SPI =====
Add-Type -AssemblyName PresentationCore,WindowsBase -ErrorAction SilentlyContinue
Add-Type @"
using System.Runtime.InteropServices;
public class Native {
  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@

function Convert-ToJpgIfNeeded([string]$SrcPath){
  try{
    $dst = Join-Path $env:TEMP (([IO.Path]::GetFileNameWithoutExtension($SrcPath)) + ".jpg")
    $stream = [System.IO.File]::OpenRead($SrcPath)
    $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
      $stream,
      [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
      [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    )
    $frame = $decoder.Frames[0]
    $encoder = New-Object System.Windows.Media.Imaging.JpegBitmapEncoder
    $encoder.QualityLevel = 95
    $encoder.Frames.Add($frame)
    $out = [System.IO.File]::Create($dst)
    $encoder.Save($out)
    $out.Close(); $stream.Close()
    return $dst
  } catch { return $SrcPath }
}

function Apply-Wallpaper([string]$Url){
  $src = Get-File $Url "wallpaper_src"; if(-not $src){ return }
  $img = Convert-ToJpgIfNeeded $src
  $dstDir="C:\ProgramData\OpenCars\wallpaper"; New-Item -Force -ItemType Directory -Path $dstDir | Out-Null
  $dst=Join-Path $dstDir "wallpaper.jpg"; Copy-Item $img $dst -Force
  # SPI_SETDESKWALLPAPER=20 ; SPIF_UPDATEINIFILE|SPIF_SENDWININICHANGE = 0x01|0x02 = 3
  [void][Native]::SystemParametersInfo(20,0,$dst,3)
  Log "Fondo aplicado: $dst"
}

function Apply-LockScreen([string]$Url){
  $src=Get-File $Url "lock_src"; if(-not $src){ return }
  $img=Convert-ToJpgIfNeeded $src
  $dstDir="C:\ProgramData\OpenCars\lock"; New-Item -Force -ItemType Directory -Path $dstDir | Out-Null
  $dst=Join-Path $dstDir "lock.jpg"; Copy-Item $img $dst -Force

  $pol="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
  if(-not (Test-Path $pol)){ New-Item $pol -Force | Out-Null }
  New-ItemProperty $pol -Name "LockScreenImage" -Value $dst -PropertyType String -Force | Out-Null

  $cloud="HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
  if(-not (Test-Path $cloud)){ New-Item $cloud -Force | Out-Null }
  New-ItemProperty $cloud -Name "DisableWindowsSpotlightOnSettings" -Value 1 -PropertyType DWord -Force | Out-Null
  New-ItemProperty $cloud -Name "DisableWindowsSpotlightFeatures"   -Value 1 -PropertyType DWord -Force | Out-Null

  Log "Lock screen fijado: $dst (bloqueá con Win+L o reiniciá sesión)"
}

function Set-ThemeDark{
  $key="HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
  if(-not (Test-Path $key)){ New-Item $key | Out-Null }
  New-ItemProperty $key -Name "AppsUseLightTheme" -Value 0 -PropertyType DWord -Force | Out-Null
  New-ItemProperty $key -Name "SystemUsesLightTheme" -Value 0 -PropertyType DWord -Force | Out-Null
  Log "Tema oscuro aplicado."
}

# ===================== Office (ODT directo oficial) =====================
function Install-OfficeODT{
  $odtExe = Get-File "https://go.microsoft.com/fwlink/?linkid=2157037" "officedeploymenttool.exe"
  if(-not $odtExe){ Log "No se pudo descargar ODT" "ERROR"; return $false }

  $odtDir = Join-Path $env:TEMP "ODT"; New-Item -ItemType Directory -Force -Path $odtDir | Out-Null
  Start-Process $odtExe -ArgumentList "/quiet /extract:`"$odtDir`"" -Wait
  $setup = Join-Path $odtDir "setup.exe"
  if(-not (Test-Path $setup)){ Log "ODT setup.exe no encontrado" "ERROR"; return $false }

  $cfg = Join-Path $odtDir "office-config.xml"
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="es-es"/>
      <ExcludeApp ID="Access"/>
      <ExcludeApp ID="Outlook"/>
      <ExcludeApp ID="OneNote"/>
      <ExcludeApp ID="Teams"/>
      <ExcludeApp ID="Lync"/>
      <ExcludeApp ID="Publisher"/>
      <ExcludeApp ID="OneDrive"/>
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
</Configuration>
"@ | Out-File -Encoding ascii $cfg

  Log "Instalando Microsoft 365 (Word/Excel/PowerPoint) vía ODT..."
  Start-Process $setup -ArgumentList "/configure `"$cfg`"" -Wait -NoNewWindow

  $apps=@("WINWORD.EXE","EXCEL.EXE","POWERPNT.EXE")
  $ok=$true; foreach($a in $apps){ if(-not (Get-Command $a -EA SilentlyContinue)){ $ok=$false } }
  if($ok){ Log "OK: Microsoft 365 (W/E/P) instalado" } else { Log "Office no verificado (puede requerir login/licencia)" "WARN" }
  return $ok
}

# ===================== Ricoh / Quiter =====================
function Install-Ricoh([string]$Url){
  $pkg=Get-File $Url "Ricoh_PCL6.exe"; if(-not $pkg){ return $false }
  $ok=Install-ExeSilent $pkg @('/s /v"/qn REBOOT=ReallySuppress"','/s /v"/qn /norestart"','/quiet /norestart')
  if($ok){ Log "OK: Ricoh PCL6" } else { Log "Ricoh no confirmada" "WARN" }
  return $ok
}
function Install-Quiter([string]$Url){
  $pkg=Get-File $Url "quiter.exe"; if(-not $pkg){ return $false }
  $ok=Install-ExeSilent $pkg
  if($ok){ Log "OK: Quiter" } else { Log "Quiter no confirmada (revisar switches/MSI)" "WARN" }
  return $ok
}

# ===================== RUN =====================
Log "=== OpenCars Setup v1.8 ==="
Winget-Ensure

# Apps winget (con upgrade para TeamViewer)
Install-App "Adobe.Acrobat.Reader.64-bit" "Adobe Acrobat Reader"
Install-App "Google.Chrome"               "Google Chrome"
Install-App "RARLab.WinRAR"               "WinRAR"
if(Install-App "TeamViewer.TeamViewer"    "TeamViewer"){ try{ winget upgrade -e --id TeamViewer.TeamViewer --silent | Out-Null }catch{} }
# AnyDesk: winget + fallback MSI oficial
$ad_ok = Install-App "AnyDeskSoftwareGmbH.AnyDesk" "AnyDesk"
if(-not $ad_ok){
  $msi = Get-File "https://download.anydesk.com/AnyDesk.msi" "AnyDesk.msi"
  if($msi){
    if(Install-ExeSilent $msi @('/qn /norestart')){ Log "OK: AnyDesk (MSI)" } else { Log "AnyDesk MSI no confirmó instalación" "WARN" }
  } else { Log "No se pudo descargar AnyDesk MSI" "WARN" }
}
# WireGuard (reintento si hace falta)
if(-not (Install-App "WireGuard.WireGuard" "WireGuard")){ Start-Sleep 2; Install-App "WireGuard.WireGuard" "WireGuard" }

# Office 365 (W/E/P) vía ODT
Install-OfficeODT

# Ricoh + Quiter
Install-Ricoh  $RicohDriverUrl
Install-Quiter $QuiterUrl

# Tema + Fondos
Set-ThemeDark
Apply-Wallpaper  $WallpaperUrl
Apply-LockScreen $LockScreenUrl

Log "=== Fin. Si algo no se ve aún (Office/Fondos), cerrá sesión o reiniciá. Logs: $Log ==="
