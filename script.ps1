<# OpenCars Setup v1.7 - instalación robusta + fondo/lockscreen #>

# --- Auto-elevate ---
$u=[Security.Principal.WindowsIdentity]::GetCurrent()
$p=New-Object Security.Principal.WindowsPrincipal($u)
if(-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
  Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""; exit
}

# === URLs (ajustá si cambia algo) ===
$WallpaperUrl   = "https://raw.githubusercontent.com/Fortecar/setup/main/Tumbnail.jfif"
$LockScreenUrl  = "https://raw.githubusercontent.com/Fortecar/setup/main/Tumbnail.jfif"
$QuiterUrl      = "https://raw.githubusercontent.com/Fortecar/setup/main/quiter.exe"
$RicohDriverUrl = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001344/0001344878/V44300/z05587L1f.exe"

# === LOG ===
$LogDir="C:\ProgramData\OpenCars\setup-logs"; New-Item -Force -ItemType Directory -Path $LogDir | Out-Null
$Log=Join-Path $LogDir ("setup-{0:yyyyMMdd-HHmmss}-{1}.log" -f (Get-Date),$env:COMPUTERNAME)
function Log([string]$m,[string]$lvl="INFO"){ "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date),$lvl,$m | Tee-Object -FilePath $Log -Append }

[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12

# === Helpers ===
function Get-File([string]$Url,[string]$Name){
  $dst=Join-Path $env:TEMP $Name
  for($i=1;$i -le 3;$i++){ try{ Log "Descarga ($i/3): $Url"; Invoke-WebRequest -Uri $Url -OutFile $dst -UseBasicParsing -TimeoutSec 180; break } catch { Start-Sleep 2; if($i -eq 3){ Log "Fallo descarga: $Url" "ERROR"; return $null } } }
  if((Get-Item $dst -EA SilentlyContinue).Length -lt 10240){ Log "Descarga sospechosa (muy pequeña): $Url" "WARN" }
  return $dst
}
function Convert-ToJpgIfNeeded([string]$Src){
  $ext=[IO.Path]::GetExtension($Src).ToLower()
  if($ext -in ".jpg",".jpeg",".png",".bmp"){ return $Src }
  try{
    Add-Type -AssemblyName System.Drawing -EA SilentlyContinue
    $img=[System.Drawing.Image]::FromFile($Src)
    $dst=Join-Path $env:TEMP (([IO.Path]::GetFileNameWithoutExtension($Src))+".jpg")
    $img.Save($dst,[System.Drawing.Imaging.ImageFormat]::Jpeg); $img.Dispose()
    Log "Convertido $ext -> JPG: $dst"; return $dst
  } catch { Log "No se pudo convertir $Src a JPG" "WARN"; return $Src }
}
function Winget-Ensure(){ if(-not (Get-Command winget -EA SilentlyContinue)){ Log "No se encontró winget (App Installer). Aborto." "ERROR"; exit 1 }
  winget source update | Out-Null
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
  foreach($a in $Tries){ try{ Log "Instalando EXE: $Exe $a"; Start-Process $Exe -ArgumentList $a -Wait -NoNewWindow; if($LASTEXITCODE -eq 0){ return $true } } catch {} }
  return $false
}

# === Fondo y LockScreen ===
function Apply-Wallpaper([string]$Url){
  $src=Get-File $Url "wallpaper_src"; if(-not $src){ return }
  $img=Convert-ToJpgIfNeeded $src
  $dstDir="C:\ProgramData\OpenCars\wallpaper"; New-Item -Force -ItemType Directory -Path $dstDir | Out-Null
  $dst=Join-Path $dstDir "wallpaper.jpg"; Copy-Item $img $dst -Force
  New-ItemProperty "HKCU:\Control Panel\Desktop" -Name Wallpaper      -Value $dst -PropertyType String -Force | Out-Null
  New-ItemProperty "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10" -PropertyType String -Force | Out-Null
  New-ItemProperty "HKCU:\Control Panel\Desktop" -Name TileWallpaper  -Value "0"  -PropertyType String -Force | Out-Null
  RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters
  Log "Fondo aplicado: $dst"
}
function Apply-LockScreen([string]$Url){
  $src=Get-File $Url "lock_src"; if(-not $src){ return }
  $img=Convert-ToJpgIfNeeded $src
  $dstDir="C:\ProgramData\OpenCars\lock"; New-Item -Force -ItemType Directory -Path $dstDir | Out-Null
  $dst=Join-Path $dstDir "lock.jpg"; Copy-Item $img $dst -Force
  # Desactivar Spotlight y fijar imagen (Windows 10/11 Pro/Ent/Edu)
  $pol="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
  if(-not (Test-Path $pol)){ New-Item $pol -Force | Out-Null }
  New-ItemProperty $pol -Name "LockScreenImage" -Value $dst -PropertyType String -Force | Out-Null
  $cloud="HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
  if(-not (Test-Path $cloud)){ New-Item $cloud -Force | Out-Null }
  New-ItemProperty $cloud -Name "DisableWindowsSpotlightOnSettings" -Value 1 -PropertyType DWord -Force | Out-Null
  New-ItemProperty $cloud -Name "DisableWindowsSpotlightFeatures"   -Value 1 -PropertyType DWord -Force | Out-Null
  Log "Lock screen fijado: $dst (puede requerir bloqueo/reinicio)"
}

# === Office (ODT: Word/Excel/PowerPoint) ===
function Install-OfficeODT{
  # Instala ODT
  if(-not (Install-App "Microsoft.OfficeDeploymentTool" "Office Deployment Tool")){ return $false }
  $odtPath=(Get-ChildItem "$env:ProgramFiles\Microsoft Office\Office Deployment Tool\setup.exe" -EA SilentlyContinue).FullName
  if(-not $odtPath){ $odtPath = (Get-ChildItem "$env:ProgramFiles(x86)\Microsoft Office\Office Deployment Tool\setup.exe" -EA SilentlyContinue).FullName }
  if(-not $odtPath){ Log "No se encontró setup.exe de ODT" "ERROR"; return $false }

  $cfgDir=Join-Path $env:TEMP "ODT"; New-Item -Force -ItemType Directory -Path $cfgDir | Out-Null
  $xml = @"
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
"@
  $cfg = Join-Path $cfgDir "office-config.xml"; $xml | Out-File -Encoding ascii $cfg
  Log "Descargando e instalando Office (Word/Excel/PowerPoint)..."
  Start-Process $odtPath -ArgumentList "/configure `"$cfg`"" -Wait -NoNewWindow
  $apps=@("WINWORD.EXE","EXCEL.EXE","POWERPNT.EXE")
  $ok=$true; foreach($a in $apps){ if(-not (Get-Command $a -EA SilentlyContinue)){ $ok=$false } }
  if($ok){ Log "OK: Microsoft 365 (W/E/P) instalado" } else { Log "Office no verificado; revisar más tarde" "WARN" }
  return $ok
}

# === Ricoh / Quiter ===
function Install-Ricoh([string]$Url){
  $pkg=Get-File $Url "Ricoh_PCL6.exe"; if(-not $pkg){ return $false }
  $ok=Install-ExeSilent $pkg @('/s /v"/qn REBOOT=ReallySuppress"','/s /v"/qn /norestart"','/quiet /norestart')
  if($ok){ Log "OK: Ricoh PCL6" } else { Log "Ricoh no confirmada" "WARN" }
  return $ok
}
function Install-Quiter([string]$Url){
  $pkg=Get-File $Url "quiter.exe"; if(-not $pkg){ return $false }
  $ok=Install-ExeSilent $pkg
  if($ok){ Log "OK: Quiter" } else { Log "Quiter no confirmada (puede requerir instalador MSI/switches propios)" "WARN" }
  return $ok
}

# === RUN ===
Log "=== OpenCars Setup v1.7 ==="
Winget-Ensure

# Acrobat / Chrome / WinRAR / TeamViewer / AnyDesk / WireGuard
Install-App "Adobe.Acrobat.Reader.64-bit" "Adobe Acrobat Reader"        # :contentReference[oaicite:0]{index=0}
Install-App "Google.Chrome"                    "Google Chrome"            # :contentReference[oaicite:1]{index=1}
Install-App "RARLab.WinRAR"                    "WinRAR"                   # :contentReference[oaicite:2]{index=2}
if(Install-App "TeamViewer.TeamViewer"         "TeamViewer"){ winget upgrade -e --id TeamViewer.TeamViewer --silent | Out-Null }  # :contentReference[oaicite:3]{index=3}
# AnyDesk: winget id y fallback MSI oficial
if(-not (Install-App "AnyDeskSoftwareGmbH.AnyDesk" "AnyDesk")){  # :contentReference[oaicite:4]{index=4}
  $msi = Get-File "https://download.anydesk.com/AnyDesk.msi" "AnyDesk.msi"  # :contentReference[oaicite:5]{index=5}
  if($msi){ if(Install-ExeSilent $msi @('/qn /norestart')){ Log "OK: AnyDesk (MSI)" } }
}
# WireGuard: intento + verificación
if(-not (Install-App "WireGuard.WireGuard" "WireGuard")){ Start-Sleep 2; Install-App "WireGuard.WireGuard" "WireGuard" }  # :contentReference[oaicite:6]{index=6}

# Office (ODT: Word/Excel/PowerPoint)
Install-OfficeODT  # :contentReference[oaicite:7]{index=7}

# Ricoh + Quiter
Install-Ricoh  $RicohDriverUrl
Install-Quiter $QuiterUrl

# Tema + Fondos
# (Tema oscuro)
$key="HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
if(-not (Test-Path $key)){ New-Item $key | Out-Null }
New-ItemProperty $key -Name "AppsUseLightTheme" -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty $key -Name "SystemUsesLightTheme" -Value 0 -PropertyType DWord -Force | Out-Null
Log "Tema oscuro aplicado."

Apply-Wallpaper  $WallpaperUrl
Apply-LockScreen $LockScreenUrl  # :contentReference[oaicite:8]{index=8}

Log "=== Fin. Si Office/WireGuard no aparecen de inmediato, reiniciá sesión/equipo. Logs: $Log ==="
