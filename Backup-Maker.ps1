# =============================================================
# ========== Initialization ===================================
# =============================================================

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

# =============================================================
# ========== Constants & Variables ============================
# =============================================================

$FontName = "Verdana"
$FontSize = 9
$FontStyle = [System.Drawing.FontStyle]::Regular
$FormColor = [System.Drawing.Color]::LightSteelBlue
$TextBoxColor = [System.Drawing.Color]::Ivory
$ButtonSizeA = New-Object System.Drawing.Size(140,30)
$ButtonSizeB = New-Object System.Drawing.Size(26,26)
$ButtonSizeC = New-Object System.Drawing.Size(110,30)
$ButtonSizeD = New-Object System.Drawing.Size(160,30)
$ButtonColor = [System.Drawing.Color]::LightCyan
$Global:SourcePath = $env:USERPROFILE
$Global:DestinationPath = (Resolve-Path -Path ([System.Environment]::CurrentDirectory)).Drive.Root
$Options = @("Neue Dateien kopieren", "Fehlende Dateien löschen", "Ältere Dateien ersetzen", "Attribute übernehmen", "Unterordner einbeziehen", "Versteckte Dateien einbeziehen", "Versteckte Ordner einbeziehen", "Protokolldatei erstellen")
$TextNewFolder = "Bitte Ordnername eingeben. Bestätigen mit 'Enter'."
$SettingsFile = "$env:LOCALAPPDATA\PowerShellTools\Backup-Maker\Settings.ini"
$OutputFile = "$env:USERPROFILE\Desktop\FileProtocol.txt"
$NL = [System.Environment]::NewLine
$SyncHash = [System.Collections.Hashtable]::Synchronized(@{})
$RSList = @("CopyIncremental")
$L_Ptr = [System.IntPtr]::new(0)
$S_Ptr = [System.IntPtr]::new(0)

$Msg_List = @{
    Start        = "Backup-Maker gestartet."
    NoDir        = "Kein gültiger Verzeichnisname angegeben."
    NoIdent      = "Quelle und Ziel dürfen nicht identisch sein."
    NoParent     = "Quelle und Ziel dürfen keine Unterverzeichnisse von einander sein."
    NoEmpty      = "Quelle darf kein leeres Verzeichnis sein."
    NoAction     = "Keine Aktion ausgewählt."
    NewFolder    = "Neuer Ordner erfolgreich erstellt."
    FailFolder   = "Erstellen fehlgeschlagen. Ungültiger Verzeichnisname."
    Analyse      = "{0} Elemente werden analysiert..."
    Copy         = "{0} Datei(en) und {1} Ordner mit {2} MB ({3} Bytes) werden kopiert..."
    Remove       = "{0} Datei(en) und {1} Ordner mit {2} MB ({3} Bytes) werden gelöscht..."
    Replace      = "{0} Datei(en) und {1} Ordner mit {2} MB ({3} Bytes) werden ersetzt..."
    ReplaceA     = "{0} Datei(en) und {1} Ordner werden attribuiert..."
    Finished     = "Backup-Vorgang beendet."
    SkipCopy     = "Kopieren übersprungen."
    SkipReplace  = "Ersetzen übersprungen."
}

$Icons_List = @{
    NewFolder = "$env:windir\system32\shell32.dll|279"
    Copy      = "$env:windir\system32\shell32.dll|54"
    Remove    = "$env:windir\system32\shell32.dll|271"
    Replace   = "$env:windir\system32\shell32.dll|295"
}

$Synopsis = "Quellverzeichnis: {0}" + $NL +
            "Zielverzeichnis: {1}" + $NL + $NL +
            "Es werden {2} Dateien vom Quellverzeichnis ins Zielverzeichnis kopiert." + $NL +
            "Im Quellverzeichnis nicht vorhandene Dateien werden im Zielverzeichnis {3}." + $NL +
            "Dateien älteren Datums werden durch gleichnamige Dateien neueren Datums {4}." + $NL + $NL +
            "Im Zielverzeichnis abweichende Attribute werden {5} überschrieben (nur bei ERSETZEN)." + $NL + $NL +
            "Sämtliche Aktionen schließen Ordner und Unterordner sowie deren Dateien {6}." + $NL +
            "Versteckte Dateien sind {7} von den Aktionen betroffen." + $NL +
            "Versteckte Ordner sind {8} von den Aktionen betroffen." + $NL + $NL +
            "Im Anschluss wird {9} Protokolldatei ERSTELLT und GELADEN."

# =============================================================
# ========== Win32Functions ===================================
# =============================================================

$Member = @'
    [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, int nIcons);

    [DllImport("User32.dll", EntryPoint = "DestroyIcon")]
    public static extern bool DestroyIcon(IntPtr hIcon);
'@

Add-Type -MemberDefinition $Member -Name WinAPI -Namespace Win32Functions

# =============================================================
# ========== Functions ========================================
# =============================================================

function Initialize-Me ([string]$FilePath)
    {
        If (!(Test-Path -Path $FilePath))
            {
                [System.Windows.Forms.MessageBox]::Show("Konnte Datei `"$FilePath`" nicht finden.","Backup-Maker: Fehler",0)
                Exit
            }

        $Data = [array](Get-Content -Path $FilePath)

        ForEach ($i in $Data)
            {
                $ht_Result += @{$i.Split("=")[0].Trim(" ") = $i.Split("=")[-1].Trim(" ")}
            }

        return $ht_Result
    }

# -------------------------------------------------------------

function Create-Object ([string]$Name, [string]$Type, [HashTable]$Data, [array]$Events, [string]$Control)
    {
        New-Variable -Name $Name -Value (New-Object System.Windows.Forms.$Type) -Scope Global -Force

        ForEach ($k in $Data.Keys) {Invoke-Expression ("`$$Name.$k = " + {$Data.$k})}
        ForEach ($i in $Events)    {Invoke-Expression ("`$$Name.$i")}
        If ($Control)              {Invoke-Expression ("`$$Control.Controls.Add(`$$Name)")}
    }

# -------------------------------------------------------------

function Create-Icons ([string]$Name, [HashTable]$List, [string]$Path)
    {
        ForEach($Key in $List.Keys)
            {
                If (Test-Path -Path ($Path + "Icon_" + $Key.ToString() + ".ico"))
                    {
                        $ht_Icons += @{$Key = [System.Drawing.Image]::FromFile($Path + "Icon_" + $Key.ToString() + ".ico")}
                    }
                ElseIf (Test-Path -Path ($Path + "Icon_" + $Key.ToString() + ".png"))
                    {
                        $ht_Icons += @{$Key = [System.Drawing.Image]::FromFile($Path + "Icon_" + $Key.ToString() + ".png")}
                    }
                Else
                    {
                        $NewIcon = [System.Drawing.Bitmap]::new(20,20)
                        $Painter = [System.Drawing.Graphics]::FromImage($NewIcon)
                        [Win32Functions.WinAPI]::ExtractIconEx($List[$Key].ToString().Split("|")[0], $List[$Key].ToString().Split("|")[-1], [ref]$L_Ptr, [ref]$S_Ptr, 1) | Out-Null
                        $Painter.DrawIcon([System.Drawing.Icon]::FromHandle($S_Ptr),[System.Drawing.Rectangle]::new(0, 0, $NewIcon.Width, $NewIcon.Height))
                        $ht_Icons += @{$Key = $NewIcon}
                        [Win32Functions.WinApi]::DestroyIcon($L_Ptr) | Out-Null
                        [Win32Functions.WinApi]::DestroyIcon($S_Ptr) | Out-Null
                    }

                $NewIcon = $ht_Icons[$Key].Clone()

                For($x = 0; $x -lt $NewIcon.Width; $x++)
                    {
                        For($y = 0; $y -lt $NewIcon.Height; $y++)
                            {
                                $Pixel = $NewIcon.GetPixel($x,$y)
                                If ($Pixel.Name -ne 0)
                                    {
                                        $NewPixel = [System.Drawing.Color]::FromArgb(($Pixel.A / 3), ($Pixel.R / 2), ($Pixel.G / 2), ($Pixel.B / 2))
                                        $NewIcon.SetPixel($x,$y,$NewPixel)
                                    }
                            }
                    }

                $ht_Icons += @{($Key.ToString() + "_g") = $NewIcon}
            }

        New-Variable -Name $Name -Value $ht_Icons -Scope Global -Force
    }

# -------------------------------------------------------------

function Write-Msg ([object]$TextBox, [bool]$NL, [bool]$Time, [bool]$Count, [string]$Num1, [string]$Num2, [string]$Num3, [string]$Num4, [string]$Msg)
    {
        If ($NL)
            {
                $NLTime = [System.Environment]::NewLine
            }

        If ($Time)
            {
                $NLTime += [string](Get-Date -Format "HH:mm:ss") + " "
            }

        If ($Count)
            {
                $Msg = $Msg -f $Num1, $Num2, $Num3, $Num4
            }

        $TextBox.AppendText($NLTime + $Msg)
    }

# -------------------------------------------------------------

function Test-Root ([string]$Path)
    {
        return (Resolve-Path -Path $Path).Drive.Root -eq $Path
    }

# -------------------------------------------------------------

function Test-Ident ([string]$PathA, [string]$PathB)
    {
        return $PathA -eq $PathB
    }

# -------------------------------------------------------------

function Test-Parent ([string]$PathA, [string]$PathB)
    {
        While ((Split-Path -Path $PathA -Parent) -ne ((Split-Path -Path $PathA -Qualifier) + "\"))
            {
                $PathA = Split-Path -Path $PathA -Parent
            }
        While ((Split-Path -Path $PathB -Parent) -ne ((Split-Path -Path $PathB -Qualifier) + "\"))
            {
                $PathB = Split-Path -Path $PathB -Parent
            }

        return $PathA -eq $PathB
    }

# -------------------------------------------------------------

function Test-Empty ([string]$Path)
    {
        return [string]::Empty -eq (Get-Item -Path $Path).GetFiles('*',[System.IO.SearchOption]::AllDirectories)
    }

# -------------------------------------------------------------

function Clean-Up ([array]$List)
    {
        If (Get-Runspace -Name "CleanUp") {return}

        $Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Name = "CleanUp"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("List",$List)
        
        $PSInstance = [System.Management.Automation.PowerShell]::Create().AddScript(
            {
                While ($true)
                    {
                        $RS = Get-Runspace -Name $List
                        ForEach($i in $RS)
                            {
                                If ($i.RunspaceAvailability -eq 'Available')
                                    {
                                        $i.Dispose()
                                        [System.GC]::Collect()
                                    }
                            }
                        Start-Sleep -Milliseconds 100
                    }
            })

        $PSInstance.Runspace = $Runspace
        $PSInstance.BeginInvoke()
    }

# -------------------------------------------------------------

function Copy-Incremental ([HashTable]$SyncHash, [bool]$Copy, [bool]$Remove, [bool]$Replace, [bool]$Attributes, [bool]$Sub, [bool]$HiddenF, [bool]$HiddenD, [bool]$OutFile, [string]$SrcPath, [string]$DstPath, [string]$OutPath)
    {
        $Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Name = "CopyIncremental"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash)
        $Runspace.SessionStateProxy.SetVariable("Copy",$Copy)
        $Runspace.SessionStateProxy.SetVariable("Remove",$Remove)
        $Runspace.SessionStateProxy.SetVariable("Replace",$Replace)
        $Runspace.SessionStateProxy.SetVariable("Attributes",$Attributes)
        $Runspace.SessionStateProxy.SetVariable("Sub",$Sub)
        $Runspace.SessionStateProxy.SetVariable("HiddenF",$HiddenF)
        $Runspace.SessionStateProxy.SetVariable("HiddenD",$HiddenD)
        $Runspace.SessionStateProxy.SetVariable("OutFile",$OutFile)
        $Runspace.SessionStateProxy.SetVariable("SrcPath",$SrcPath)
        $Runspace.SessionStateProxy.SetVariable("DstPath",$DstPath)
        $Runspace.SessionStateProxy.SetVariable("OutPath",$OutPath)

        $PSInstance = [System.Management.Automation.PowerShell]::Create().AddScript(
            {
                $SyncHash.bt_Copy.Enabled = $false
                $SyncHash.bt_All.Enabled = $false
                $SyncHash.bt_Exit.Enabled = $false

                $NL = [System.Environment]::NewLine

                $Protocol_List = @{
                                    CopyD      = "$NL[COPIED ITEMS - DIRECTORIES]$NL"
                                    CopyF      = "$NL[COPIED ITEMS - FILES]$NL"
                                    RemoveD    = "$NL[REMOVED ITEMS - DIRECTORIES]$NL"
                                    RemoveF    = "$NL[REMOVED ITEMS - FILES]$NL"
                                    ReplaceD   = "$NL[REPLACED ITEMS - DIRECTORIES]$NL"
                                    ReplaceF   = "$NL[REPLACED ITEMS - FILES]$NL"
                                    ReplaceDA  = "$NL[REPLACED ITEMS - DIRECTORY ATTRIBUTES]$NL"
                                    ReplaceFA  = "$NL[REPLACED ITEMS - FILE ATTRIBUTES]$NL"
                                  }

                $Formatter = @{
                                A = @{
                                        RemoveD    = "{0}"
                                     }

                                B = @{
                                        CopyD      = "{0} => {1}"
                                        ReplaceD   = "{0} <= {1}"
                                     }

                                C = @{
                                        CopyF      = "{0} => {1} {2} Bytes"
                                        RemoveF    = "{0} == {1} {2} Bytes"
                                        ReplaceF   = "{0} <= {1} {2} Bytes"
                                     }

                                D = @{
                                        ReplaceA  = "{0} == {1} => {2}"
                                     }

                                E = @{
                                        DirTotal   = "$NL{0} Directories"
                                        FileTotal  = "$NL{0} Files / {1} MB / {2} Bytes"
                                        AttriTotal = "$NL{0} Files"
                                     }
                              }

                $DiskSpace_Msg = "Nicht genügend Speicherplatz für folgende Aktion vorhanden:" + $NL + $NL +
                                 "{0}" + $NL + $NL +
                                 "Es wird empfohlen, Speicher durch Löschen nicht benötigter Daten freizugeben oder die Aktion zu überspringen." + $NL + $NL +
                                 "Ungefähr benötigter Speicherplatz: {1:N2} MB"

                $Cmd_Src = 'Get-ChildItem -Path $SrcPath -File'
                $Cmd_Dst = 'Get-ChildItem -Path $DstPath -File'

                If ($Sub)
                    {
                        $Cmd_Src = $Cmd_Src.Replace('File','Recurse')
                        $Cmd_Dst = $Cmd_Dst.Replace('File','Recurse')
                    }

                If ($Sub -and $HiddenD)
                    {
                        $HiddenSrc = Get-ChildItem -Path $SrcPath -Recurse -Directory -Hidden -ErrorAction SilentlyContinue
                        $HiddenDst = Get-ChildItem -Path $DstPath -Recurse -Directory -Hidden -ErrorAction SilentlyContinue

                        ForEach($i in $HiddenSrc)
                            {
                                $i.Attributes -= [System.IO.FileAttributes]::Hidden
                            }

                        ForEach($i in $HiddenDst)
                            {
                                $i.Attributes -= [System.IO.FileAttributes]::Hidden
                            }
                    }
                
                $SrcItems = Invoke-Expression $Cmd_Src
                $DstItems = Invoke-Expression $Cmd_Dst

                If ($HiddenF)
                    {
                        $SrcItems += Invoke-Expression ($Cmd_Src + ' -Hidden')
                        $DstItems += Invoke-Expression ($Cmd_Dst + ' -Hidden')
                    }

                If ($Sub -and $HiddenD)
                    {
                        ForEach($i in $HiddenSrc)
                            {
                                $i.Attributes += [System.IO.FileAttributes]::Hidden
                                If ($i.FullName -in $SrcItems.FullName)
                                    {
                                        $SrcItems[$SrcItems.FullName.IndexOf($i.FullName)].Attributes += [System.IO.FileAttributes]::Hidden
                                    }
                            }

                        ForEach($i in $HiddenDst)
                            {
                                $i.Attributes += [System.IO.FileAttributes]::Hidden
                                If ($i.FullName -in $DstItems.FullName)
                                    {
                                        $DstItems[$DstItems.FullName.IndexOf($i.FullName)].Attributes += [System.IO.FileAttributes]::Hidden
                                    }
                            }
                    }

                $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $true, ($SrcItems.Count + $DstItems.Count), $null, $null, $null, $SyncHash.Msg_List.Analyse)

                ForEach($i in $SrcItems)
                    {
                        If ($i.PSIsContainer)
                            {
                                $LstSrcD += @([pscustomobject]@{Name = $i.FullName; Time = $i.LastWriteTime; Attributes = $i.Attributes})
                            }
                        Else
                            {
                                $LstSrcF += @([pscustomobject]@{Name = $i.FullName; Dir = $i.DirectoryName; Time = $i.LastWriteTime; Attributes = $i.Attributes; Bytes = $i.Length})
                            }
                    }

                ForEach($i in $DstItems)
                    {
                        If ($i.PSIsContainer)
                            {
                                $LstDstD += @([pscustomobject]@{Name = $i.FullName; Time = $i.LastWriteTime; Attributes = $i.Attributes})
                            }
                        Else
                            {
                                $LstDstF += @([pscustomobject]@{Name = $i.FullName; Dir = $i.DirectoryName; Time = $i.LastWriteTime; Attributes = $i.Attributes; Bytes = $i.Length})
                            }
                    }

                ForEach($i in $LstSrcD)
                    {
                        $Val = $i.Attributes.value__ -band [System.Convert]::ToInt32(10100111,2)
                        $i.Attributes = [System.IO.FileAttributes]$Val
                    }

                ForEach($i in $LstSrcF)
                    {
                        $Val = $i.Attributes.value__ -band [System.Convert]::ToInt32(10100111,2)
                        $i.Attributes = [System.IO.FileAttributes]$Val
                    }

                ForEach($i in $LstDstD)
                    {
                        $Val = $i.Attributes.value__ -band [System.Convert]::ToInt32(10100111,2)
                        $i.Attributes = [System.IO.FileAttributes]$Val
                    }

                ForEach($i in $LstDstF)
                    {
                        $Val = $i.Attributes.value__ -band [System.Convert]::ToInt32(10100111,2)
                        $i.Attributes = [System.IO.FileAttributes]$Val
                    }

                $LstSrcD = [array]($LstSrcD | Sort-Object -Descending -Property Name)
                $LstSrcF = [array]($LstSrcF | Sort-Object -Descending -Property Name)
                $LstDstD = [array]($LstDstD | Sort-Object -Descending -Property Name)
                $LstDstF = [array]($LstDstF | Sort-Object -Descending -Property Name)

                ForEach($i in $LstSrcD)
                    {
                        $CmpWDstD += @([pscustomobject]@{Name = $i.Name.Replace($SrcPath,$DstPath); Time = $i.Time; Attributes = $i.Attributes})
                    }

                ForEach($i in $LstSrcF)
                    {
                        $CmpWDstF += @([pscustomobject]@{Name = $i.Name.Replace($SrcPath,$DstPath); Dir = $i.Dir; Time = $i.Time; Attributes = $i.Attributes; Bytes = $i.Bytes})
                    }

                ForEach($i in $LstDstD)
                    {
                        $CmpWSrcD += @([pscustomobject]@{Name = $i.Name.Replace($DstPath,$SrcPath); Time = $i.Time; Attributes = $i.Attributes})
                    }

                ForEach($i in $LstDstF)
                    {
                        $CmpWSrcF += @([pscustomobject]@{Name = $i.Name.Replace($DstPath,$SrcPath); Dir = $i.Dir; Time = $i.Time; Attributes = $i.Attributes; Bytes = $i.Bytes})
                    }

                ForEach($i in $LstSrcD)
                    {
                        If ($i.Name -in $CmpWSrcD.Name)
                            {
                                $TimeCmpD += @([pscustomobject]@{Src = $i; Dst = $LstDstD[$CmpWSrcD.Name.IndexOf($i.Name)]})
                            }
                        Else
                            {
                                $ToCopyD += @([pscustomobject]@{Src = $i; Dst = $CmpWDstD[$LstSrcD.Name.IndexOf($i.Name)]})
                            }
                    }

                ForEach($i in $LstSrcF)
                    {
                        If ($i.Name -in $CmpWSrcF.Name)
                            {
                                $TimeCmpF += @([pscustomobject]@{Src = $i; Dst = $LstDstF[$CmpWSrcF.Name.IndexOf($i.Name)]})
                            }
                        Else
                            {
                                $ToCopyF += @([pscustomobject]@{Src = $i; Dst = $CmpWDstF[$LstSrcF.Name.IndexOf($i.Name)]})
                            }
                    }

                ForEach($i in $LstDstD)
                    {
                        If ($i.Name -notin $CmpWDstD.Name)
                            {
                                $ToRemoveD += @($i)
                            }
                    }

                ForEach($i in $LstDstF)
                    {
                        If ($i.Name -notin $CmpWDstF.Name)
                            {
                                $ToRemoveF += @($i)
                            }
                    }

                ForEach($i in $TimeCmpF)
                    {
                        If ($i.Src.Time -gt $i.Dst.Time)
                            {
                                $ToReplaceF += @($i)
                            }
                        ElseIf ($i.Src.Attributes -ne $i.Dst.Attributes)
                            {
                                $ToReplaceFA += @($i)
                            }
                    }

                ForEach($i in $TimeCmpD)
                    {
                        If ($i.Src.Time -gt $i.Dst.Time)
                            {
                                $ToReplaceD += @($i)
                            }
                        ElseIf ($i.Src.Attributes -ne $i.Dst.Attributes)
                            {
                                $ToReplaceDA += @($i)
                            }
                    }

                $SyncHash.pb_Copy.Value = 0
                $SyncHash.pb_Remove.Value = 0
                $SyncHash.pb_Replace.Value = 0

                $ToCopyD = [array]($ToCopyD | Sort-Object -Property {$_.Src.Name})
                $ToCopyF = [array]($ToCopyF | Sort-Object -Property {$_.Src.Dir, $_.Src.Name})

                $ToRemoveD = [array]($ToRemoveD | Sort-Object -Property {$_.Name})
                $ToRemoveF = [array]($ToRemoveF | Sort-Object -Property {$_.Dir, $_.Name})

                $ToReplaceD = [array]($ToReplaceD | Sort-Object -Property {$_.Src.Name})
                $ToReplaceF = [array]($ToReplaceF | Sort-Object -Property {$_.Src.Dir,$_.Src.Name})

                $ToReplaceDA = [array]($ToReplaceDA | Sort-Object -Property {$_.Src.Name})
                $ToReplaceFA = [array]($ToReplaceFA | Sort-Object -Property {$_.Src.Dir,$_.Src.Name})

                $FindLongest = @($null,$null,$null,$null,$null)
                $Longest     = @($null,$null,$null,$null,$null)
                $Data        = @($null,$null,$null,$null)

                $TransferSize = @{
                    Copy    = ($ToCopyF.Src.Bytes    | Measure-Object -Sum).Sum
                    Remove  = ($ToRemoveF.Bytes      | Measure-Object -Sum).Sum
                    Replace = ($ToReplaceF.Src.Bytes | Measure-Object -Sum).Sum
                }

                $ScriptMB = {
                    If     (!($null -eq $this.Src.Bytes)) {"{0:N2}" -f ($this.Src.Bytes / 1MB)}
                    ElseIf (!($null -eq $this.Bytes))     {"{0:N2}" -f ($this.Bytes / 1MB)}
                    ElseIf (!($null -eq $this))           {"{0:N2}" -f ($this / 1MB)}
                    Else                                  {"0"}
                }

                $ScriptB = {
                    If     (!($null -eq $this.Src.Bytes)) {If ($this.Src.Bytes) {"{0:#,#}" -f $this.Src.Bytes} Else {"0"}}
                    ElseIf (!($null -eq $this.Bytes))     {If ($this.Bytes)     {"{0:#,#}" -f $this.Bytes}     Else {"0"}}
                    ElseIf (!($null -eq $this))           {If ($this)           {"{0:#,#}" -f $this}           Else {"0"}}
                    Else                                  {"0"}
                }

                $TransferSize.Keys | ForEach-Object {$TransferSize.$_ | Add-Member -MemberType ScriptMethod -Name "ToMByte" -Value $ScriptMB -Force}
                $TransferSize.Keys | ForEach-Object {$TransferSize.$_ | Add-Member -MemberType ScriptMethod -Name "ToByte" -Value $ScriptB -Force}

                $ToCopyF | Add-Member -MemberType ScriptMethod -Name "ToMByte" -Value $ScriptMB -Force
                $ToCopyF | Add-Member -MemberType ScriptMethod -Name "ToByte" -Value $ScriptB -Force

                $ToReplaceF | Add-Member -MemberType ScriptMethod -Name "ToMByte" -Value $ScriptMB -Force
                $ToReplaceF | Add-Member -MemberType ScriptMethod -Name "ToByte" -Value $ScriptB -Force

                $ToRemoveF | Add-Member -MemberType ScriptMethod -Name "ToMByte" -Value $ScriptMB -Force
                $ToRemoveF | Add-Member -MemberType ScriptMethod -Name "ToByte" -Value $ScriptB -Force

                If ($Copy)
                    {
                        Do {
                            $Diskspace = (Resolve-Path -Path $DstPath).Drive.Free
                            $Go = $Diskspace -ge $TransferSize.Copy

                            If (-not $Go)
                                {
                                    $Space = [Math]::Abs(($Diskspace - $TransferSize.Copy) / 1MB)
                                    $SyncHash.lb_DiskSpace.Text = ($DiskSpace_Msg -f "Kopieren", $Space)
                                    $Choice = $SyncHash.DiskSpace_Form.ShowDialog()
                                }
                            }
                        Until ($Go -or ($Choice -eq [System.Windows.Forms.DialogResult]::Abort))

                        If ($Go)
                            {
                                $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $true, $ToCopyF.Count, $ToCopyD.Count, $TransferSize.Copy.ToMByte(), $TransferSize.Copy.ToByte(), $SyncHash.Msg_List.Copy)

                                ForEach($i in $ToCopyD)
                                    {
                                        Copy-Item -LiteralPath $i.Src.Name -Destination $i.Dst.Name -Force
                                        Set-ItemProperty -LiteralPath $i.Dst.Name -Name Attributes -Value $i.Src.Attributes
                                    }

                                ForEach($i in $ToCopyF)
                                    {
                                        Copy-Item -LiteralPath $i.Src.Name -Destination $i.Dst.Name -Force

                                        $Percent = ($ToCopyF.IndexOf($i) / $ToCopyF.Count) * 100
                                        $SyncHash.pb_Copy.Value = $Percent
                                    }

                                $SyncHash.pb_Copy.Value = 100

                                If ($OutFile)
                                    {
                                        If ($ToCopyD)
                                            {
                                                $Data[0] += @($Protocol_List.CopyD); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                ForEach($i in $ToCopyD)
                                                    {
                                                        $FindLongest[0] += @($i.Src.Name.Length); $FindLongest[1] += @($i.Dst.Name.Length)
                                                        $Data[0] += @($Formatter.B.CopyD); $Data[1] += @($i.Src.Name); $Data[2] += @($i.Dst.Name); $Data[3] += @([string]::Empty)
                                                    }
                                                $Data[0] += @($Formatter.E.DirTotal); $Data[1] += @($ToCopyD.Count); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                            }

                                        If ($ToCopyF)
                                            {
                                                $Data[0] += @($Protocol_List.CopyF); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                ForEach($i in $ToCopyF)
                                                    {
                                                        $FindLongest[0] += @($i.Src.Name.Length); $FindLongest[1] += @($i.Dst.Name.Length); $FindLongest[2] += @($i.ToByte().Length)
                                                        $Data[0] += @($Formatter.C.CopyF); $Data[1] += @($i.Src.Name); $Data[2] += @($i.Dst.Name); $Data[3] += @($i.ToByte())
                                                    }
                                                $Data[0] += @($Formatter.E.FileTotal); $Data[1] += @($ToCopyF.Count); $Data[2] += @($TransferSize.Copy.ToMByte()); $Data[3] += @($TransferSize.Copy.ToByte())
                                            }
                                    }
                            }
                        Else
                            {
                                $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $false, $null, $null, $null, $null, $SyncHash.Msg_List.SkipCopy)
                            }
                    }

                If ($Remove)
                    {
                        $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $true, $ToRemoveF.Count, $ToRemoveD.Count, $TransferSize.Remove.ToMByte(), $TransferSize.Remove.ToByte(), $SyncHash.Msg_List.Remove)

                        ForEach($i in $ToRemoveF)
                            {
                                Remove-Item -LiteralPath $i.Name -Force

                                $Percent = ($ToRemoveF.IndexOf($i) / $ToRemoveF.Count) * 100
                                $SyncHash.pb_Remove.Value = $Percent
                            }

                        ForEach($i in $ToRemoveD)
                            {
                                Remove-Item -LiteralPath $i.Name -Recurse -Force
                            }

                        $SyncHash.pb_Remove.Value = 100

                        If ($OutFile)
                            {
                                If ($ToRemoveD)
                                    {
                                        $Data[0] += @($Protocol_List.RemoveD); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                        ForEach($i in $ToRemoveD)
                                            {
                                                $FindLongest[0] += @($i.Name.Length)
                                                $Data[0] += @($Formatter.A.RemoveD); $Data[1] += @($i.Name); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                            }
                                        $Data[0] += @($Formatter.E.DirTotal); $Data[1] += @($ToRemoveD.Count); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                    }

                                If ($ToRemoveF)
                                    {
                                        $Data[0] += @($Protocol_List.RemoveF); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                        ForEach($i in $ToRemoveF)
                                            {
                                                $FindLongest[0] += @($i.Name.Length); $FindLongest[2] += @($i.ToByte().Length)
                                                $Data[0] += @($Formatter.C.RemoveF); $Data[1] += @($i.Name); $Data[2] += @([string]::Empty); $Data[3] += @($i.ToByte())
                                            }
                                        $Data[0] += @($Formatter.E.FileTotal); $Data[1] += @($ToRemoveF.Count); $Data[2] += @($TransferSize.Remove.ToMByte()); $Data[3] += @($TransferSize.Remove.ToByte())
                                    }
                            }
                    }

                If ($Replace)
                    {
                        Do {
                            $Diskspace = (Resolve-Path -Path $DstPath).Drive.Free
                            $Go = $Diskspace -ge $TransferSize.Replace

                            If (-not $Go)
                                {
                                    $Space = [Math]::Abs(($Diskspace - $TransferSize.Replace) / 1MB)
                                    $SyncHash.lb_DiskSpace.Text = ($DiskSpace_Msg -f "Ersetzen", $Space)
                                    $Choice = $SyncHash.DiskSpace_Form.ShowDialog()
                                }
                            }
                        Until ($Go -or ($Choice -eq [System.Windows.Forms.DialogResult]::Abort))

                        If ($Go)
                            {
                                $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $true, $ToReplaceF.Count, $ToReplaceD.Count, $TransferSize.Replace.ToMByte(), $TransferSize.Replace.ToByte(), $SyncHash.Msg_List.Replace)

                                ForEach($i in $ToReplaceF)
                                    {
                                        Copy-Item -LiteralPath $i.Src.Name -Destination $i.Dst.Name -Force

                                        $Percent = ($ToReplaceF.IndexOf($i) / $ToReplaceF.Count) * 100
                                        $SyncHash.pb_Replace.Value = $Percent
                                    }

                                ForEach($i in $ToReplaceD)
                                    {
                                        Set-ItemProperty -LiteralPath $i.Dst.Name -Name LastWriteTime -Value $i.Src.Time
                                        Set-ItemProperty -LiteralPath $i.Dst.Name -Name Attributes -Value $i.Src.Attributes
                                    }

                                $SyncHash.pb_Replace.Value = 100

                                If ($OutFile)
                                    {
                                        If ($ToReplaceD)
                                            {
                                                $Data[0] += @($Protocol_List.ReplaceD); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                ForEach($i in $ToReplaceD)
                                                    {
                                                        $FindLongest[0] += @($i.Dst.Name.Length); $FindLongest[1] += @($i.Src.Name.Length)
                                                        $Data[0] += @($Formatter.B.ReplaceD); $Data[1] += @($i.Dst.Name); $Data[2] += @($i.Src.Name); $Data[3] += @([string]::Empty)
                                                    }
                                                $Data[0] += @($Formatter.E.DirTotal); $Data[1] += @($ToReplaceD.Count); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                            }

                                        If ($ToReplaceF)
                                            {
                                                $Data[0] += @($Protocol_List.ReplaceF); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                ForEach($i in $ToReplaceF)
                                                    {
                                                        $FindLongest[0] += @($i.Dst.Name.Length); $FindLongest[1] += @($i.Src.Name.Length); $FindLongest[2] += @($i.ToByte().Length)
                                                        $Data[0] += @($Formatter.C.ReplaceF); $Data[1] += @($i.Dst.Name); $Data[2] += @($i.Src.Name); $Data[3] += @($i.ToByte())
                                                    }
                                                $Data[0] += @($Formatter.E.FileTotal); $Data[1] += @($ToReplaceF.Count); $Data[2] += @($TransferSize.Replace.ToMByte()); $Data[3] += @($TransferSize.Replace.ToByte())
                                            }
                                    }

                                If ($Attributes)
                                    {
                                        $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $true, $ToReplaceFA.Count, $ToReplaceDA.Count, $null, $null, $SyncHash.Msg_List.ReplaceA)

                                        ForEach($i in $ToReplaceFA)
                                            {
                                                Set-ItemProperty -LiteralPath $i.Dst.Name -Name Attributes -Value $i.Src.Attributes
                                            }

                                        ForEach($i in $ToReplaceDA)
                                            {
                                                Set-ItemProperty -LiteralPath $i.Dst.Name -Name Attributes -Value $i.Src.Attributes
                                            }

                                        If ($OutFile)
                                            {
                                                If ($ToReplaceDA)
                                                    {
                                                        $Data[0] += @($Protocol_List.ReplaceDA); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                        ForEach($i in $ToReplaceDA)
                                                            {
                                                                $Val = $i.Src.Attributes.value__ -bor [System.Convert]::ToInt32(10000,2)
                                                                $i.Src.Attributes = [System.IO.FileAttributes]$Val
                                                                $Val = $i.Dst.Attributes.value__ -bor [System.Convert]::ToInt32(10000,2)
                                                                $i.Dst.Attributes = [System.IO.FileAttributes]$Val
                                                                $FindLongest[0] += @($i.Dst.Name.Length); $FindLongest[3] += @($i.Dst.Attributes.ToString().Length); $FindLongest[4] += @($i.Src.Attributes.ToString().Length)
                                                                $Data[0] += @($Formatter.D.ReplaceA); $Data[1] += @($i.Dst.Name); $Data[2] += @($i.Dst.Attributes.ToString()); $Data[3] += @($i.Src.Attributes.ToString())
                                                            }
                                                        $Data[0] += @($Formatter.E.DirTotal); $Data[1] += @($ToReplaceDA.Count); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                    }

                                                If ($ToReplaceFA)
                                                    {
                                                        $Data[0] += @($Protocol_List.ReplaceFA); $Data[1] += @([string]::Empty); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                        ForEach($i in $ToReplaceFA)
                                                            {
                                                                $FindLongest[0] += @($i.Dst.Name.Length); $FindLongest[3] += @($i.Dst.Attributes.ToString().Length); $FindLongest[4] += @($i.Src.Attributes.ToString().Length)
                                                                $Data[0] += @($Formatter.D.ReplaceA); $Data[1] += @($i.Dst.Name); $Data[2] += @($i.Dst.Attributes.ToString()); $Data[3] += @($i.Src.Attributes.ToString())
                                                            }
                                                        $Data[0] += @($Formatter.E.AttriTotal); $Data[1] += @($ToReplaceFA.Count); $Data[2] += @([string]::Empty); $Data[3] += @([string]::Empty)
                                                    }
                                            }
                                    }
                            }
                        Else
                            {
                                $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $false, $null, $null, $null, $null, $SyncHash.Msg_List.SkipReplace)
                            }
                    }

                $SyncHash.WriteMsg.Invoke($SyncHash.tb_Events, $true, $true, $false, $null, $null, $null, $null, $SyncHash.Msg_List.Finished)

                $SyncHash.bt_Copy.Enabled = $true
                $SyncHash.bt_All.Enabled = $true
                $SyncHash.bt_Exit.Enabled = $true

                If ($OutFile)
                    {
                        For($i = 0; $i -lt $FindLongest.Count; $i++)
                            {
                                $Longest[$i] = ($FindLongest[$i] | Measure-Object -Maximum).Maximum
                            }

                        For($i = 0; $i -lt $Data[0].Count; $i++)
                            {
                                If ($Protocol_List.ContainsValue($Data[0][$i]))
                                    {
                                        $FileProtocol += @($Data[0][$i])
                                    }
                                ElseIf ($Formatter.A.ContainsValue($Data[0][$i]))
                                    {
                                        $FileProtocol += @($Data[0][$i] -f $Data[1][$i])
                                    }
                                ElseIf ($Formatter.B.ContainsValue($Data[0][$i]))
                                    {
                                        $FileProtocol += @($Data[0][$i] -f ($Data[1][$i]).PadRight($Longest[0],"."), $Data[2][$i])
                                    }
                                ElseIf ($Formatter.C.ContainsValue($Data[0][$i]))
                                    {
                                        $FileProtocol += @($Data[0][$i] -f ($Data[1][$i]).PadRight($Longest[0],"."), ($Data[2][$i]).PadRight($Longest[1],"."), ($Data[3][$i]).PadLeft($Longest[2]," "))
                                    }
                                ElseIf ($Formatter.D.ContainsValue($Data[0][$i]))
                                    {
                                        $FileProtocol += @($Data[0][$i] -f ($Data[1][$i]).PadRight($Longest[0],"."), ($Data[2][$i]).PadLeft($Longest[3],"."), ($Data[3][$i]).PadLeft($Longest[4],"."))
                                    }
                                ElseIf ($Formatter.E.ContainsValue($Data[0][$i]))
                                    {
                                        $FileProtocol += @($Data[0][$i] -f $Data[1][$i], $Data[2][$i], $Data[3][$i])
                                    }
                            }

                        Set-Content -Value ("Created by Backup-Maker - " + [string](Get-Date -Format "dd.MM.yyyy HH:mm:ss")) -Path $OutPath -Force
                        Add-Content -Value $FileProtocol -Path $OutPath
                        Invoke-Item -Path $OutPath
                    }
            })

        $PSInstance.Runspace = $Runspace
        $PSInstance.BeginInvoke()
    }

# -------------------------------------------------------------

function Munition-ComboBox ([object]$Name, [string]$Path)
    {
        $Name.Items.Clear()

        If (Split-Path -Path $Path -Parent)
            {
                [void]$Name.Items.Add([PSCustomObject]@{Name = Split-Path -Path $Path -Parent; DisplayName = '[...]'})
            }
        
        ForEach($i in Get-Childitem -Path $Path -Directory)
            {
                [void]$Name.Items.Add([PSCustomObject]@{Name = $i.FullName; DisplayName = $i.FullName})
            }

        ForEach($i in Get-PSDrive | Where-Object {$_.Free})
            {
                If ($i.Root -ne (Resolve-Path -Path $Path).Drive.Root)
                    {
                        If ($i.Description)
                            {
                                [void]$Name.Items.Add([PSCustomObject]@{Name = $i.Root; DisplayName = $i.Description + " ($i`:)"})
                            }
                        Else
                            {
                                [void]$Name.Items.Add([PSCustomObject]@{Name = $i.Root; DisplayName = $i.Root})
                            }
                    }

            }

        $Name.DisplayMember = 'DisplayName'
    }

# -------------------------------------------------------------

function Reverse-Me ([object]$Button, [object]$TextBox, [string]$Path)
    {
        $Form.Controls | ForEach-Object {If (($_ -ne $Button) -and ($_ -ne $TextBox)) {$_.Enabled = !($_.Enabled)}}
        $TextBox.ReadOnly = (!($TextBox.ReadOnly))

        If ($TextBox.ReadOnly)
            {
                $TextBox.Text = $Path
                $Form.ActiveControl = $bt_Exit
            }
        Else
            {
                $TextBox.Text = $TextNewFolder
                $Form.ActiveControl = $TextBox
            }
    }

# -------------------------------------------------------------

function Fill-Synopsis ([string]$Synopsis, [string]$SrcPath, [string]$DstPath, [bool]$Copy, [bool]$Remove, [bool]$Replace, [bool]$Attributes, [bool]$Sub, [bool]$HiddenF, [bool]$HiddenD, [bool]$OutFile)
    {
        $Var01 = $SrcPath
        $Var02 = $DstPath
        If ($Copy)       {$Var03 = "ALLE"}             Else {$Var03 = "KEINE"}
        If ($Remove)     {$Var04 = "GELÖSCHT"}         Else {$Var04 = "NICHT GELÖSCHT"}
        If ($Replace)    {$Var05 = "ERSETZT"}          Else {$Var05 = "NICHT ERSETZT"}
        If ($Attributes) {$Var06 = "MIT DEM ORIGINAL"} Else {$Var06 = "NICHT"}
        If ($Sub)        {$Var07 = "EIN"}              Else {$Var07 = "AUS"}
        If ($HiddenF)    {$Var08 = "EBENFALLS"}        Else {$Var08 = "NICHT"}
        If ($HiddenD)    {$Var09 = "EBENFALLS"}        Else {$Var09 = "NICHT"}
        If ($OutFile)    {$Var10 = "eine"}             Else {$Var10 = "KEINE"}

        $Synopsis = $Synopsis -f $Var01, $Var02, $Var03 ,$Var04 ,$Var05, $Var06, $Var07, $Var08, $Var09, $Var10

        return $Synopsis
    }

# =============================================================
# ========== Code =============================================
# =============================================================

$Paths = Initialize-Me -FilePath $SettingsFile

# -------------------------------------------------------------

Create-Object -Name Tooltip -Type Tooltip

# -------------------------------------------------------------

Create-Icons -Name Icons -List $Icons_List -Path $Paths.IconFolder

$IconMax = New-Object -TypeName System.Drawing.Size(($Icons.Values.Width | Measure-Object -Maximum).Maximum,($Icons.Values.Height | Measure-Object -Maximum).Maximum)

# =============================================================
# ========== Form =============================================
# =============================================================

$ht_Data = @{
            ClientSize = New-Object System.Drawing.Size(600,560)
            StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
            Icon = $Paths.IconFolder + "Backup-Maker.ico"
            Text = "Backup-Maker"
            BackColor = $FormColor
            FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
            MaximizeBox = $false
            KeyPreview = $true
            }

$ar_Events = @(
                {Add_Load(
                    {
                        $this.TopMost = $true
                        $this.ActiveControl = $bt_Exit
                        Clean-Up -List $RSList
                        Write-Msg -TextBox $tb_Events -NL $false -Time $true -Count $false -Msg $Msg_List.Start
                    })}
                {Add_FormClosing(
                    {
                        If (Get-Runspace -Name $RSList)
                            {
                                If ([System.Windows.Forms.MessageBox]::Show("Der Backup-Prozess wurde noch nicht abgeschlossen. Soll dieser wirklich beendet werden (nicht empfohlen)?","Achtung!",4) -eq [System.Windows.Forms.DialogResult]::No)
                                    {
                                        $_.Cancel = $true
                                    }
                            }

                        If (!($_.Cancel))
                            {
                                $RS = Get-Runspace -Name ($RSList + "CleanUp")
                                ForEach($i in $RS)
                                    {
                                        $i.Dispose()
                                        [System.GC]::Collect()
                                    }
                            }
                    })}
              )

Create-Object -Name Form -Type Form -Data $ht_Data -Events $ar_Events

# -------------------------------------------------------------

$SyncHash.Form = $Form

# =============================================================
# ========== Form: Labels =====================================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = 10
            Width = 200
            Text = "Quelle:"
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            }

Create-Object -Name lb_Source -Type Label -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top = 110
$ht_Data.Text = "Ziel:"

Create-Object -Name lb_Destination -Type Label -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top = 210
$ht_Data.Text = "Fortschritt:"

Create-Object -Name lb_Progress -Type Label -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Left = $Form.ClientSize.Width / 2 + 10
$ht_Data.Text = "Optionen:"

Create-Object -Name lb_Options -Type Label -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top = 410
$ht_Data.Left = 10
$ht_Data.Text = "Ereignisse:"

Create-Object -Name lb_Events -Type Label -Data $ht_Data -Control Form

# =============================================================
# ========== Form: TextBoxes ==================================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = $lb_Source.Bounds.Bottom
            Width = $Form.ClientSize.Width - 50
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            Text = $Global:SourcePath
            BackColor = $TextBoxColor
            Cursor = [System.Windows.Forms.Cursors]::IBeam
            ReadOnly = $true
            }

$ar_Events = @(
                {Add_KeyDown(
                    {
                        If (!($this.Readonly))
                            {
                                If ($_.KeyCode -eq "Enter")
                                    {
                                        If (!(Test-Path -Path ($Global:SourcePath + "\" + $this.Text) -PathType Container) -and ($this.Text -ne $TextNewFolder))
                                            {
                                                New-Item -Path $Global:SourcePath -Name $this.Text -ItemType Directory
                                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.NewFolder
                                                Munition-ComboBox -Name $cb_Source -Path $Global:SourcePath
                                            }
                                        Else
                                            {
                                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.FailFolder
                                            }

                                        Reverse-Me -Button $bt_SourceNewFolder -TextBox $this -Path $Global:SourcePath
                                    }
                                ElseIf ($_.KeyCode -eq "Escape")
                                    {
                                        Reverse-Me -Button $bt_SourceNewFolder -TextBox $this -Path $Global:SourcePath
                                    }
                            }
                    })}
                {Add_Click(
                    {
                        If ($this.Readonly)
                            {
                                $this.SelectAll()
                            }
                        Else
                            {
                                $this.Clear()
                            }
                    })}
              )

Create-Object -Name tb_Source -Type TextBox -Data $ht_Data -Events $ar_Events -Control Form

# -------------------------------------------------------------

$ht_Data.Top = $lb_Destination.Bounds.Bottom
$ht_Data.Text = $Global:DestinationPath

$ar_Events[0] = $ar_Events[0].ToString().Replace("Source","Destination")

Create-Object -Name tb_Destination -Type TextBox -Data $ht_Data -Events $ar_Events -Control Form

# -------------------------------------------------------------

$ht_Data.Top = $lb_Events.Bounds.Bottom
$ht_Data.Height = $Form.ClientSize.Height - $lb_Events.Bounds.Bottom - 10
$ht_Data.Width = $Form.ClientSize.Width - 20
$ht_Data.Font = New-Object System.Drawing.Font($FontName, ($FontSize - 1), $FontStyle)
$ht_Data.Text = $null
$ht_Data += @{
             Multiline = $true
             ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
             WordWrap = $true
             }

Create-Object -Name tb_Events -Type TextBox -Data $ht_Data -Control Form

# -------------------------------------------------------------

$SyncHash.tb_Events = $tb_Events

# =============================================================
# ========== Form: Buttons ====================================
# =============================================================

$ht_Data = @{
            Left = $tb_Source.Bounds.Right + 4
            Top = $tb_Source.Bounds.Top
            Size = $ButtonSizeB
            FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
            BackColor = $ButtonColor
            Image = $Icons.NewFolder
            ImageAlign = [System.Drawing.ContentAlignment]::BottomCenter
            Cursor = [System.Windows.Forms.Cursors]::Hand
            }

$ar_Events = @(
                {Add_Click(
                    {
                        Reverse-Me -Button $this -TextBox $tb_Source -Path $Global:SourcePath
                    })}
                {Add_MouseHover(
                    {
                        $Tooltip.SetToolTip($this,"Klicken um neuen Ordner zu erstellen.")
                    })}
              )

Create-Object -Name bt_SourceNewFolder -Type Button -Data $ht_Data -Events $ar_Events -Control Form

# -------------------------------------------------------------

$ht_Data.Left = $tb_Destination.Bounds.Right + 4
$ht_Data.top = $tb_Destination.Bounds.Top

$ar_Events[0] = $ar_Events[0].ToString().Replace("Source","Destination")

Create-Object -Name bt_DestinationNewFolder -Type Button -Data $ht_Data -Events $ar_Events -Control Form

# =============================================================
# ========== Form: ComboBoxes =================================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = $tb_Source.Bounds.Bottom + 4
            Width = $Form.ClientSize.Width - 20
            BackColor = $TextBoxColor
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            Cursor = [System.Windows.Forms.Cursors]::Hand
            }

$ar_Events = @(
                {Add_SelectedIndexChanged(
                    {
                        $Global:SourcePath = $this.SelectedItem.Name

                        Munition-ComboBox -Name $this -Path $Global:SourcePath

                        $tb_Source.Text = $Global:SourcePath
                    })}
              )

Create-Object -Name cb_Source -Type ComboBox -Data $ht_Data -Events $ar_Events -Control Form

Munition-ComboBox -Name $cb_Source -Path $Global:SourcePath

# -------------------------------------------------------------

$ht_Data.Top = $tb_Destination.Bounds.Bottom + 4

$ar_Events[0] = $ar_Events[0].ToString().Replace("Source","Destination")

Create-Object -Name cb_Destination -Type ComboBox -Data $ht_Data -Events $ar_Events -Control Form

Munition-ComboBox -Name $cb_Destination -Path $Global:DestinationPath

# =============================================================
# ========== Form: Panels =====================================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = $lb_Options.Bounds.Bottom
            Width = $Form.ClientSize.Width / 2 - 20
            Height = 97
            BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            }

Create-Object -Name pn_Panel -Type Panel -Data $ht_Data -Control Form

# =============================================================
# ========== Form: Labels =====================================
# =============================================================

$ht_Data = @{
            Left = 14
            Top = [Math]::Round($pn_Panel.Bounds.Top + $pn_Panel.ClientSize.Height / 3)
            Width = $pn_Panel.Width - 8
            Height = 2
            BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
            }

Create-Object -Name lb_Line1 -Type Label -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top = [Math]::Round($pn_Panel.Bounds.Top + $pn_Panel.ClientSize.Height / 3 * 2)

Create-Object -Name lb_Line2 -Type Label -Data $ht_Data -Control Form

# =============================================================
# ========== Form: PictureBoxes ===============================
# =============================================================

$ht_Data = @{
            Left = 20 - (($IconMax.Width - 16) / 2 )
            Top = $pn_Panel.Bounds.Top + (10 - (($IconMax.Height - 16) / 2 ))
            Width = 20
            Height = 20
            Image = $Icons.Copy_g
            BorderStyle = [System.Windows.Forms.BorderStyle]::None
            }

Create-Object -Name pb_IconCopy -Type PictureBox -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top += $pn_Panel.ClientSize.Height / 3
$ht_Data.Image = $Icons.Remove_g

Create-Object -Name pb_IconRemove -Type PictureBox -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top += $pn_Panel.ClientSize.Height / 3
$ht_Data.Image = $Icons.Replace_g

Create-Object -Name pb_IconReplace -Type PictureBox -Data $ht_Data -Control Form

# =============================================================
# ========== Form: ProgressBars ===============================
# =============================================================

$ht_Data = @{
            Left = 43
            Top = $pn_Panel.Bounds.Top + 8
            Width = $pn_Panel.Width - 40
            Height = 20
            Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            Visible = $false
            }

Create-Object -Name pb_Copy -Type ProgressBar -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top += $pn_Panel.ClientSize.Height / 3

Create-Object -Name pb_Remove -Type ProgressBar -Data $ht_Data -Control Form

# -------------------------------------------------------------

$ht_Data.Top += $pn_Panel.ClientSize.Height / 3

Create-Object -Name pb_Replace -Type ProgressBar -Data $ht_Data -Control Form

# -------------------------------------------------------------

$SyncHash.pb_Copy = $pb_Copy
$SyncHash.pb_Remove = $pb_Remove
$SyncHash.pb_Replace = $pb_Replace

# -------------------------------------------------------------

$pn_Panel.SendToBack()

# =============================================================
# ========== Form: CheckedListBoxes ===========================
# =============================================================

$ht_Data = @{
            Left = $Form.ClientSize.Width / 2 + 10
            Top = $lb_Options.Bounds.Bottom
            Width = $Form.ClientSize.Width / 2 - 20
            Height = 100
            BackColor = [System.Drawing.Color]::Ivory
            Font = New-Object System.Drawing.Font($FontName, ($FontSize - 1), $FontStyle)
            BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            CheckOnClick = $true
            }

$ar_Events = @(
                {Items.AddRange($Options)}
                {Add_SelectedIndexChanged(
                    {
                        $pb_Copy.Visible = $this.GetItemChecked(0)
                        $pb_Remove.Visible = $this.GetItemChecked(1)
                        $pb_Replace.Visible = $this.GetItemChecked(2)

                        If ($this.GetItemChecked(0))
                            {$pb_IconCopy.Image = $Icons.Copy}
                        Else
                            {$pb_IconCopy.Image = $Icons.Copy_g}

                        If ($this.GetItemChecked(1))
                            {$pb_IconRemove.Image = $Icons.Remove}
                        Else
                            {$pb_IconRemove.Image = $Icons.Remove_g}

                        If ($this.GetItemChecked(2))
                            {$pb_IconReplace.Image = $Icons.Replace}
                        Else
                            {$pb_IconReplace.Image = $Icons.Replace_g}
                    })}
              )

Create-Object -Name clb_Box -Type CheckedListBox -Data $ht_Data -Events $ar_Events -Control Form

# =============================================================
# ========== Form: Buttons ====================================
# =============================================================

$ht_Data = @{
            Left = 30
            Top = $clb_Box.Bounds.Bottom + 26
            Size = $ButtonSizeA
            FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
            BackColor = $ButtonColor
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            Text = "Backup"
            TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            Cursor = [System.Windows.Forms.Cursors]::Hand
            }

$ar_Events = @(
                {Add_Click(
                    {
                        If ((Test-Root -Path $Global:SourcePath) -or (Test-Root -Path $Global:DestinationPath))
                            {
                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.NoDir
                            }
                        ElseIf (Test-Ident -PathA $Global:SourcePath -PathB $Global:DestinationPath)
                            {
                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.NoIdent
                            }
                        ElseIf (Test-Parent -PathA $Global:SourcePath -PathB $Global:DestinationPath)
                            {
                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.NoParent
                            }
                        ElseIf (Test-Empty -Path $Global:SourcePath)
                            {
                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.NoEmpty
                            }
                        ElseIf (!($clb_Box.GetItemChecked(0)) -and !($clb_Box.GetItemChecked(1)) -and !($clb_Box.GetItemChecked(2)))
                            {
                                Write-Msg -TextBox $tb_Events -NL $true -Time $true -Count $false -Msg $Msg_List.NoAction
                            }
                        ElseIf ($Copy_Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
                            {
                                $SyncHash.WriteMsg = (Get-Item "function:Write-Msg").ScriptBlock
                                $SyncHash.Msg_List = $Msg_List
                                Copy-Incremental -SyncHash $SyncHash -Copy $clb_Box.GetItemChecked(0) -Remove $clb_Box.GetItemChecked(1) -Replace $clb_Box.GetItemChecked(2) -Attributes $clb_Box.GetItemChecked(3) -Sub $clb_Box.GetItemChecked(4) -HiddenF $clb_Box.GetItemChecked(5) -HiddenD $clb_Box.GetItemChecked(6) -OutFile $clb_Box.GetItemChecked(7) -SrcPath $Global:SourcePath -DstPath $Global:DestinationPath -OutPath $OutputFile
                            }
                    })}
              )

Create-Object -Name bt_Copy -Type Button -Data $ht_Data -Events $ar_Events -Control Form

# -------------------------------------------------------------

$ht_Data.Left = $Form.ClientSize.Width / 2 - $ButtonSizeA.Width / 2
$ht_Data.Text = "Alles markieren"

$ar_Events = @(
                {Add_Click(
                    {
                        For($i = 0; $i -lt $clb_Box.Items.Count; $i++)
                            {
                                $clb_Box.SetItemChecked($i,$true)
                            }
                        $pb_Copy.Show()
                        $pb_Remove.Show()
                        $pb_Replace.Show()
                        $pb_IconCopy.Image = $Icons.Copy
                        $pb_IconRemove.Image = $Icons.Remove
                        $pb_IconReplace.Image = $Icons.Replace
                    })}
              )

Create-Object -Name bt_All -Type Button -Data $ht_Data -Events $ar_Events -Control Form

# -------------------------------------------------------------

$ht_Data.Left = $Form.ClientSize.Width - $ButtonSizeA.Width - 30
$ht_Data.Text = "Exit"

$ar_Events = @(
                {Add_Click({$Form.Close()})}
              )

Create-Object -Name bt_Exit -Type Button -Data $ht_Data -Events $ar_Events -Control Form

# -------------------------------------------------------------

$SyncHash.bt_Copy = $bt_Copy
$SyncHash.bt_All = $bt_All
$SyncHash.bt_Exit = $bt_Exit

# =============================================================
# ========== Copy_Form ========================================
# =============================================================

$ht_Data = @{
            ClientSize = New-Object System.Drawing.Size(600,300)
            StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
            Icon = $Paths.IconFolder + "Backup-Maker.ico"
            Text = "Aufgabenformular"
            BackColor = $FormColor
            FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
            MaximizeBox = $false
            MinimizeBox = $false
            ShowInTaskbar = $false
            }

$ar_Events = @(
                {Add_Load(
                    {
                        $Form.TopMost = $false
                        $this.TopMost = $true
                        $this.ActiveControl = $bt_Cancel
                        $tb_Synopsis.Text = Fill-Synopsis -Synopsis $Synopsis -SrcPath $Global:SourcePath -DstPath $Global:DestinationPath -Copy $clb_Box.GetItemChecked(0) -Remove $clb_Box.GetItemChecked(1) -Replace $clb_Box.GetItemChecked(2) -Attributes $clb_Box.GetItemChecked(3) -Sub $clb_Box.GetItemChecked(4) -HiddenF $clb_Box.GetItemChecked(5) -HiddenD $clb_Box.GetItemChecked(6) -OutFile $clb_Box.GetItemChecked(7)
                    })}
                {Add_FormClosed({$Form.TopMost = $true})}
              )

Create-Object -Name Copy_Form -Type Form -Data $ht_Data -Events $ar_Events

# =============================================================
# ========== Copy_Form: Labels ================================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = 10
            Width = $Copy_Form.ClientSize.Width - 20
            Text = "Synopsis"
            Font = New-Object System.Drawing.Font($FontName, ($FontSize + 1), $FontStyle)
            TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            }

Create-Object -Name lb_Synopsis -Type Label -Data $ht_Data -Control Copy_Form

# =============================================================
# ========== Copy_Form: TextBoxes =============================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = $lb_Synopsis.Bounds.Bottom + 4
            Width = $Copy_Form.ClientSize.Width - 20
            Height = $Copy_Form.ClientSize.Height - $lb_Synopsis.Bounds.Bottom - 70
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            BackColor = $TextBoxColor
            Cursor = [System.Windows.Forms.Cursors]::IBeam
            ReadOnly = $true
            Multiline = $true
            ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
            WordWrap = $true
            }

Create-Object -Name tb_Synopsis -Type TextBox -Data $ht_Data -Control Copy_Form

# =============================================================
# ========== Copy_Form: Buttons ===============================
# =============================================================

$ht_Data = @{
            Left = 60
            Top = $tb_Synopsis.Bounds.Bottom + 20
            Size = $ButtonSizeC
            FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
            BackColor = $ButtonColor
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            Text = "Start"
            TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            Cursor = [System.Windows.Forms.Cursors]::Hand
            DialogResult = [System.Windows.Forms.DialogResult]::OK
            }

Create-Object -Name bt_OK -Type Button -Data $ht_Data -Control Copy_Form

# -------------------------------------------------------------

$ht_Data.Left = $Copy_Form.ClientSize.Width - $ButtonSizeC.Width - 60
$ht_Data.Text = "Abbruch"
$ht_Data.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

Create-Object -Name bt_Cancel -Type Button -Data $ht_Data -Control Copy_Form

# -------------------------------------------------------------

$Copy_Form.AcceptButton = $bt_OK
$Copy_Form.CancelButton = $bt_Cancel

# =============================================================
# ========== DiskSpace_Form ===================================
# =============================================================

$ht_Data = @{
            ClientSize = New-Object System.Drawing.Size(600,240)
            StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
            Icon = $Paths.IconFolder + "Backup-Maker.ico"
            Text = "Benachrichtigungsformular"
            BackColor = $FormColor
            FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
            MaximizeBox = $false
            MinimizeBox = $false
            ShowInTaskbar = $false
            }

$ar_Events = @(
                {Add_Load(
                    {
                        $Form.TopMost = $false
                        $this.TopMost = $true
                        $this.ActiveControl = $bt_Skip
                    })}
                {Add_FormClosed({$Form.TopMost = $true})}
              )

Create-Object -Name DiskSpace_Form -Type Form -Data $ht_Data -Events $ar_Events

# -------------------------------------------------------------

$SyncHash.DiskSpace_Form = $DiskSpace_Form

# =============================================================
# ========== DiskSpace_Form: Labels ===========================
# =============================================================

$ht_Data = @{
            Left = 10
            Top = 10
            Width = $DiskSpace_Form.ClientSize.Width - 20
            Height = 160
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            }

Create-Object -Name lb_DiskSpace -Type Label -Data $ht_Data -Control DiskSpace_Form

# -------------------------------------------------------------

$SyncHash.lb_DiskSpace = $lb_DiskSpace

# =============================================================
# ========== DiskSpace_Form: Buttons ==========================
# =============================================================

$ht_Data = @{
            Left = 60
            Top = $lb_DiskSpace.Bounds.Bottom + 20
            Size = $ButtonSizeD
            FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
            BackColor = $ButtonColor
            Font = New-Object System.Drawing.Font($FontName, $FontSize, $FontStyle)
            Text = "Erneut versuchen"
            TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            Cursor = [System.Windows.Forms.Cursors]::Hand
            DialogResult = [System.Windows.Forms.DialogResult]::Retry
            }

Create-Object -Name bt_Retry -Type Button -Data $ht_Data -Control DiskSpace_Form

# -------------------------------------------------------------

$ht_Data.Left = $DiskSpace_Form.ClientSize.Width - $ButtonSizeD.Width - 60
$ht_Data.Text = "Überspringen"
$ht_Data.DialogResult = [System.Windows.Forms.DialogResult]::Abort

Create-Object -Name bt_Abort -Type Button -Data $ht_Data -Control DiskSpace_Form

# -------------------------------------------------------------

$DiskSpace_Form.AcceptButton = $bt_Retry
$DiskSpace_Form.CancelButton = $bt_Abort

# =============================================================
# ========== Start ============================================
# =============================================================

$SyncHash.Form.ShowDialog()