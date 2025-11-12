<# OpenCars Setup v1.5 - Instala apps + Quiter + Ricoh + Fondo/Lock + Tema #>

# ==== URLs (según tu repo y Ricoh) ====
$WallpaperUrl   = "https://raw.githubusercontent.com/Fortecar/setup/main/Tumbnail.jfif"
$LockScreenUrl  = "https://raw.githubusercontent.com/Fortecar/setup/main/Tumbnail.jfif"
$QuiterUrl      = "https://raw.githubusercontent.com/Fortecar/setup/main/quiter.exe"
$RicohDriverUrl = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001344/0001344878/V44300/z05587L1f.exe"

# ==== LOG ====
$LogDir="C:\ProgramData\OpenCars\setup-logs"; New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$LogFile=Join-Path $LogDir ("setup-{0:yyyyMMdd-HHmmss}-{1}.log" -f (Get-Date), $env:COMPUTERNAME)
function Log([string]$m,[string]$lvl="INFO"){ "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date),$lvl,$m | Tee-Object -FilePath $LogFile -Append }

# ==== PRERREQUISITOS ====
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
function Ensure-Admin{
  $isAdmin=([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if(-not $isAdmin){ Log "Ejecutar PowerShell como Administrador." "ERROR"; exit 1 }
}
function Ensure-Winget{
  if(-not (Get-Command winget -ErrorAction SilentlyContinue)){
    Log "No se encontró winget (App Installer). Aborto." "ERROR"; exit 1
  }
}
Ensure-Admin

# ==== UTILIDADES ====
function Get-File([string]$Url,[string]$Name){
  $dst = Join-Path $env:TEMP $Name
  for($i=1;$i -le 3;$i++){
    try{ Log "Descargando ($i/3): $Url"; Invoke-WebRequest -Uri $Url -OutFile $dst -UseBasicParsing -TimeoutSec 120; break }
    catch{ Start-Sleep -Seconds 2; if($i -eq 3){ Log "Fallo descarga: $Url" "ERROR"; return $null } }
  }
  return $dst
}
function Convert-ToJpgIfNeeded([string]$SrcPath){
  $ext = [IO.Path]::GetExtension($SrcPath).ToLower()
  if($ext -in @(".jpg",".jpeg",".png",".bmp")){ return $SrcPath }
  try{
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
    $img=[System.Drawing.Image]::FromFile($SrcPath)
    $dst=Join-Path $env:TEMP (([IO.Path]::GetFileNameWithoutExtension($SrcPath)) + ".jpg")
    $img.Save($dst,[System.Drawing.Imaging.ImageFormat]::Jpeg); $img.Dispose()
    Log "Convertido $ext -> JPG: $dst"; return $dst
  } catch { Log "No se pudo convertir $SrcPath a JPG, se usará tal cual." "WARN"; return $SrcPath }
}
function Install-App([string]$Id,[string]$Name){
  $exists = winget list --id $Id --accept-source-agreements 2>$null | Select-String $Id
  if($exists){ Log "OK (ya instalado): $Name"; return }
  Log "Instalando: $Name"
  winget install --id $Id --silent --accept-source-agreements --accept-package-agreements --source winget
  if($LASTEXITCODE -eq 0){ Log "OK: $Name" } else { Log "Revisar $Name (exit $LASTEXITCODE)" "WARN" }
}
function Install-ExeSilent([string]$ExePath,[string[]]$Tries=@('/s /v"/qn REBOOT=ReallySuppress"','/s /v"/qn /norestart"','/quiet /norestart','/S','/silent','/verysilent')){
  foreach($a in $Tries){
    try{ Log "Instalando: $ExePath $a"; Start-Process $ExePath -ArgumentList $a -Wait -NoNewWindow
      if($LASTEXITCODE -eq 0){ Log "OK (EXE): $ExePath"; return $true } } catch {}
  }
  Log "No se pudo instalar EXE en modo silencioso: $ExePath" "WARN"; return $false
}
function Install-MSI([string]$MsiPath){
  Start-Process "msiexec.exe" -ArgumentList "/i `"$MsiPath`" /qn /norestart" -Wait -NoNewWindow
  if($LASTEXITCODE -eq 0){ Log "OK (MSI): $MsiPath"; return $true } else { Log "Error MSI (exit $LASTEXITCODE): $MsiPath" "WARN"; return $false }
}
function Set-ThemeDark{
  $key="HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
  if(-not (Test-Path $key)){ New-Item $key | Out-Null }
  New-ItemProperty $key -Name "AppsUseLightTheme" -Value 0 -PropertyType DWord -Force | Out-Null
  New-ItemProperty $key -Name "SystemUsesLightTheme" -Value 0 -PropertyType DWord -Force | Out-Null
  Log "Tema oscuro aplicado."
}
function Set-Wallpaper([string]$Url){
  $src = Get-File $Url "wallpaper_src"; if(-not $src){ return }
  $img = Convert-ToJpgIfNeeded $src
  New-ItemProperty "HKCU:\Control Panel\Desktop" -Name Wallpaper      -Value $img -PropertyType String -Force | Out-Null
  New-ItemProperty "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10" -PropertyType String -Force | Out-Null
  New-ItemProperty "HKCU:\Control Panel\Desktop" -Name TileWallpaper  -Value "0"  -PropertyType String -Force | Out-Null
  RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters
  Log "Fondo de escritorio aplicado."
}
function Set-LockScreen([string]$Url){
  $src = Get-File $Url "lock_src"; if(-not $src){ return }
  $img = Convert-ToJpgIfNeeded $src
  $dstDir="C:\Windows\Web\Screen"; New-Item -ItemType Directory -Force -Path $dstDir | Out-Null
  $dst=Join-Path $dstDir "OpenCarsLock.jpg"; Copy-Item $img $dst -Force
  $polKey="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
  if(-not (Test-Path $polKey)){ New-Item $polKey -Force | Out-Null }
  New-ItemProperty $polKey -Name "LockScreenImage" -Value $dst -PropertyType String -Force | Out-Null
  Log "Pantalla de bloqueo configurada."
}

# ==== INSTALACIONES ====
Ensure-Winget
Log "Instalación de aplicaciones (winget) ..."
Install-App "Adobe.Acrobat.Reader.64-bit" "Adobe Acrobat Reader DC"
Install-App "Google.Chrome"               "Google Chrome"
Install-App "Microsoft.Office"            "Microsoft 365 Apps (Office)"
Install-App "WireGuard.WireGuard"         "WireGuard"
Install-App "RARLab.WinRAR"               "WinRAR"
Install-App "TeamViewer.TeamViewer"       "TeamViewer"
Install-App "AnyDeskSoftwareGmbH.AnyDesk" "AnyDesk"

# Quiter (EXE público desde GitHub)
$quiter = Get-File $QuiterUrl "quiter.exe"
if($quiter){ if(-not (Install-ExeSilent $quiter)){ Log "Instalación de Quiter no confirmada" "WARN" } }

# Ricoh PCL6 (EXE oficial)
$ricoh = Get-File $RicohDriverUrl "Ricoh_PCL6.exe"
if($ricoh){ if(-not (Install-ExeSilent $ricoh)){ Log "Instalación Ricoh no confirmada" "WARN" } }

# Tema + Fondos
Set-ThemeDark
Set-Wallpaper  $WallpaperUrl
Set-LockScreen $LockScreenUrl

Log "Finalizado. Si no ves el tema/fondos, cerrá sesión o reiniciá."
