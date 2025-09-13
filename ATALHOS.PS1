Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Drawing.Printing

# ---------------------- FORMULÁRIO ----------------------
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Atualizador de Atalhos"
$form.ClientSize      = New-Object System.Drawing.Size(580,320)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $true

# ---------------------- TABS ----------------------
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock      = [System.Windows.Forms.DockStyle]::Fill
$tabPage1             = New-Object System.Windows.Forms.TabPage("Atualizador")
$tabPageCopiaExe      = New-Object System.Windows.Forms.TabPage("Copiar exe")
$tabPageSobre         = New-Object System.Windows.Forms.TabPage("Sobre")
$tabControl.TabPages.AddRange(@($tabPage1,$tabPageCopiaExe,$tabPageSobre))
$form.Controls.Add($tabControl)

# ---------------------- ABA “Atualizador” ----------------------

# Linha 1: Seleção de pasta SACS
$lblSacs = New-Object System.Windows.Forms.Label
$lblSacs.Text     = "Selecione a pasta SACS:"
$lblSacs.AutoSize = $true
$lblSacs.Location = New-Object System.Drawing.Point(10,20)
$tabPage1.Controls.Add($lblSacs)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Location      = New-Object System.Drawing.Point(150,18)
$comboBox.Size          = New-Object System.Drawing.Size(350,20)
$comboBox.SelectedIndex = -1
$tabPage1.Controls.Add($comboBox)

# Linha 2: Atalhos, BootStrap, Registrar OCX, Instalar Fontes
$chkAtalhos = New-Object System.Windows.Forms.CheckBox
$chkAtalhos.Location = New-Object System.Drawing.Point(10,60)
$chkAtalhos.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkAtalhos)
$lblAtalhos = New-Object System.Windows.Forms.Label
$lblAtalhos.Text     = "Atalhos"
$lblAtalhos.AutoSize = $true
$lblAtalhos.Location = New-Object System.Drawing.Point(30,62)
$tabPage1.Controls.Add($lblAtalhos)

$chkBootstrap = New-Object System.Windows.Forms.CheckBox
$chkBootstrap.Location = New-Object System.Drawing.Point(110,60)
$chkBootstrap.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkBootstrap)
$lblBootstrap = New-Object System.Windows.Forms.Label
$lblBootstrap.Text     = "BootStrap"
$lblBootstrap.AutoSize = $true
$lblBootstrap.Location = New-Object System.Drawing.Point(130,62)
$tabPage1.Controls.Add($lblBootstrap)

$chkRegistrarOCX = New-Object System.Windows.Forms.CheckBox
$chkRegistrarOCX.Location = New-Object System.Drawing.Point(210,60)
$chkRegistrarOCX.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkRegistrarOCX)
$lblRegistrarOCX = New-Object System.Windows.Forms.Label
$lblRegistrarOCX.Text     = "Registrar OCX"
$lblRegistrarOCX.AutoSize = $true
$lblRegistrarOCX.Location = New-Object System.Drawing.Point(230,62)
$tabPage1.Controls.Add($lblRegistrarOCX)

$chkFontes = New-Object System.Windows.Forms.CheckBox
$chkFontes.Location = New-Object System.Drawing.Point(310,60)
$chkFontes.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkFontes)
$lblFontes = New-Object System.Windows.Forms.Label
$lblFontes.Text     = "Instalar Fontes"
$lblFontes.AutoSize = $true
$lblFontes.Location = New-Object System.Drawing.Point(330,62)
$tabPage1.Controls.Add($lblFontes)

# Linha 3: COD_HEMO, INSTITUIÇÃO
$chkCodHemo = New-Object System.Windows.Forms.CheckBox
$chkCodHemo.Location = New-Object System.Drawing.Point(10,100)
$chkCodHemo.Size     = New-Object System.Drawing.Size(20,20)
$chkCodHemo.Add_CheckedChanged({
    $txtCodHemo.Enabled = $chkCodHemo.Checked
    if (-not $chkCodHemo.Checked) { $txtCodHemo.Clear() }
})
$tabPage1.Controls.Add($chkCodHemo)

$lblCodHemo = New-Object System.Windows.Forms.Label
$lblCodHemo.Text     = "COD_HEMO"
$lblCodHemo.AutoSize = $true
$lblCodHemo.Location = New-Object System.Drawing.Point(30,102)
$tabPage1.Controls.Add($lblCodHemo)

$txtCodHemo = New-Object System.Windows.Forms.TextBox
$txtCodHemo.Location  = New-Object System.Drawing.Point(110,100)
$txtCodHemo.Size      = New-Object System.Drawing.Size(100,20)
$txtCodHemo.MaxLength = 15
$txtCodHemo.Enabled   = $false
$tabPage1.Controls.Add($txtCodHemo)

$chkInstituicao = New-Object System.Windows.Forms.CheckBox
$chkInstituicao.Location = New-Object System.Drawing.Point(230,100)
$chkInstituicao.Size     = New-Object System.Drawing.Size(20,20)
$chkInstituicao.Add_CheckedChanged({
    $txtInstituicao.Enabled = $chkInstituicao.Checked
    if (-not $chkInstituicao.Checked) { $txtInstituicao.Clear() }
})
$tabPage1.Controls.Add($chkInstituicao)

$lblInstituicao = New-Object System.Windows.Forms.Label
$lblInstituicao.Text     = "INSTITUIÇÃO"
$lblInstituicao.AutoSize = $true
$lblInstituicao.Location = New-Object System.Drawing.Point(250,102)
$tabPage1.Controls.Add($lblInstituicao)

$txtInstituicao = New-Object System.Windows.Forms.TextBox
$txtInstituicao.Location  = New-Object System.Drawing.Point(350,100)
$txtInstituicao.Size      = New-Object System.Drawing.Size(100,20)
$txtInstituicao.MaxLength = 10
$txtInstituicao.Enabled   = $false
$tabPage1.Controls.Add($txtInstituicao)

# Linha 4: Impressora Etiqueta / Tipo Impressora
$chkPrinterEtiqueta = New-Object System.Windows.Forms.CheckBox
$chkPrinterEtiqueta.Location = New-Object System.Drawing.Point(10,140)
$chkPrinterEtiqueta.Size     = New-Object System.Drawing.Size(20,20)
$chkPrinterEtiqueta.Add_CheckedChanged({
    $comboPrinterEtiqueta.Enabled = $chkPrinterEtiqueta.Checked
    if (-not $chkPrinterEtiqueta.Checked) { $comboPrinterEtiqueta.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkPrinterEtiqueta)

$lblPrinterEtiqueta = New-Object System.Windows.Forms.Label
$lblPrinterEtiqueta.Text     = "Impressora Etiqueta"
$lblPrinterEtiqueta.AutoSize = $true
$lblPrinterEtiqueta.Location = New-Object System.Drawing.Point(30,142)
$tabPage1.Controls.Add($lblPrinterEtiqueta)

$comboPrinterEtiqueta = New-Object System.Windows.Forms.ComboBox
$comboPrinterEtiqueta.DropDownStyle = 'DropDownList'
$comboPrinterEtiqueta.Location      = New-Object System.Drawing.Point(150,140)
$comboPrinterEtiqueta.Size          = New-Object System.Drawing.Size(185,20)
$comboPrinterEtiqueta.Enabled       = $false
$tabPage1.Controls.Add($comboPrinterEtiqueta)

$chkTipoImp = New-Object System.Windows.Forms.CheckBox
$chkTipoImp.Location = New-Object System.Drawing.Point(345,140)
$chkTipoImp.Size     = New-Object System.Drawing.Size(20,20)
$chkTipoImp.Add_CheckedChanged({
    $comboTipoImp.Enabled = $chkTipoImp.Checked
    if (-not $chkTipoImp.Checked) { $comboTipoImp.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkTipoImp)

$lblTipoImp = New-Object System.Windows.Forms.Label
$lblTipoImp.Text     = "Tipo Impressora"
$lblTipoImp.AutoSize = $true
$lblTipoImp.Location = New-Object System.Drawing.Point(365,142)
$tabPage1.Controls.Add($lblTipoImp)

$comboTipoImp = New-Object System.Windows.Forms.ComboBox
$comboTipoImp.DropDownStyle = 'DropDownList'
$comboTipoImp.Location      = New-Object System.Drawing.Point(460,140)
$comboTipoImp.Size          = New-Object System.Drawing.Size(100,20)
$comboTipoImp.Items.AddRange(@("1 - ALEGRO","2 - ZEBRA"))
$comboTipoImp.SelectedIndex = -1
$comboTipoImp.Enabled       = $false
$tabPage1.Controls.Add($comboTipoImp)

# Linha 5: Impressora Gráfica / DPI
$chkPrinterGraphics = New-Object System.Windows.Forms.CheckBox
$chkPrinterGraphics.Location = New-Object System.Drawing.Point(10,180)
$chkPrinterGraphics.Size     = New-Object System.Drawing.Size(20,20)
$chkPrinterGraphics.Add_CheckedChanged({
    $comboPrinterGraphics.Enabled = $chkPrinterGraphics.Checked
    if (-not $chkPrinterGraphics.Checked) { $comboPrinterGraphics.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkPrinterGraphics)

$lblPrinterGraphics = New-Object System.Windows.Forms.Label
$lblPrinterGraphics.Text     = "Impressora Gráfica"
$lblPrinterGraphics.AutoSize = $true
$lblPrinterGraphics.Location = New-Object System.Drawing.Point(30,182)
$tabPage1.Controls.Add($lblPrinterGraphics)

$comboPrinterGraphics = New-Object System.Windows.Forms.ComboBox
$comboPrinterGraphics.DropDownStyle = 'DropDownList'
$comboPrinterGraphics.Location      = New-Object System.Drawing.Point(150,180)
$comboPrinterGraphics.Size          = New-Object System.Drawing.Size(185,20)
$comboPrinterGraphics.Enabled       = $false
$tabPage1.Controls.Add($comboPrinterGraphics)

$chkDPI2 = New-Object System.Windows.Forms.CheckBox
$chkDPI2.Location = New-Object System.Drawing.Point(345,180)
$chkDPI2.Size     = New-Object System.Drawing.Size(20,20)
$chkDPI2.Add_CheckedChanged({
    $txtDPI2.Enabled = $chkDPI2.Checked
    if (-not $chkDPI2.Checked) { $txtDPI2.Clear() }
})
$tabPage1.Controls.Add($chkDPI2)

$lblDPI2 = New-Object System.Windows.Forms.Label
$lblDPI2.Text     = "DPI"
$lblDPI2.AutoSize = $true
$lblDPI2.Location = New-Object System.Drawing.Point(365,182)
$tabPage1.Controls.Add($lblDPI2)

$txtDPI2 = New-Object System.Windows.Forms.TextBox
$txtDPI2.Location  = New-Object System.Drawing.Point(390,180)
$txtDPI2.Size      = New-Object System.Drawing.Size(30,20)
$txtDPI2.MaxLength = 3
$txtDPI2.Enabled   = $false
$tabPage1.Controls.Add($txtDPI2)

# Botão Atualizar (rodapé, ao lado do DPI)
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text     = "Atualizar"
$updateButton.Size     = New-Object System.Drawing.Size(100,30)
$updateButton.Location = New-Object System.Drawing.Point(464,250)
$tabPage1.Controls.Add($updateButton)

# Progress bar e label (para OCX)
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text     = "Aguarde, registrando as OCX."
$lblProgress.AutoSize = $true
$lblProgress.Location = New-Object System.Drawing.Point(10,300)
$lblProgress.Visible  = $false
$tabPage1.Controls.Add($lblProgress)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,320)
$progressBar.Size     = New-Object System.Drawing.Size(530,10)
$progressBar.Minimum  = 0
$progressBar.Value    = 0
$progressBar.Visible  = $false
$tabPage1.Controls.Add($progressBar)

# ---------------------- FUNÇÕES ----------------------
function Populate-SacsFolders {
    $comboBox.Items.Clear()
    try {
        Get-ChildItem -Path 'C:\' -Directory |
          Where-Object Name -like 'SACS*' |
          Sort-Object Name |
          ForEach-Object { $comboBox.Items.Add($_.Name) }
        $comboBox.SelectedIndex = -1
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erro ao listar pastas:`n$_",
            "Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Populate-Printers {
    $comboPrinterGraphics.Items.Clear()
    $comboPrinterEtiqueta.Items.Clear()
    foreach ($p in [System.Drawing.Printing.PrinterSettings]::InstalledPrinters) {
        $comboPrinterGraphics.Items.Add($p)
        $comboPrinterEtiqueta.Items.Add($p)
    }
    $comboPrinterGraphics.SelectedIndex = -1
    $comboPrinterEtiqueta.SelectedIndex = -1
}

$comboBox.Add_SelectedIndexChanged({
    if ($comboBox.SelectedIndex -ge 0) {
        $sel      = $comboBox.SelectedItem.ToString().Trim()
        $pathBoot = Join-Path (Join-Path 'C:\' $sel) 'bootstrap'
        if (Test-Path $pathBoot) {
            $resp = [System.Windows.Forms.MessageBox]::Show(
                "Bootstrap encontrado em '$sel'. Deseja pré-marcá-lo?",
                "Bootstrap",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            $chkBootstrap.Checked = ($resp -eq [System.Windows.Forms.DialogResult]::Yes)
        } else {
            $chkBootstrap.Checked = $false
        }
    }
})

$form.Add_Shown({
    Populate-SacsFolders
    Populate-Printers
})

# ---------------------- CLIQUE “ATUALIZAR” ----------------------
$updateButton.Add_Click({
    $anyChanged     = $false
    $fontsInstalled = $false

    $needsFolder = $chkCodHemo.Checked -or $chkBootstrap.Checked -or $chkPrinterGraphics.Checked `
                 -or $chkPrinterEtiqueta.Checked -or $chkInstituicao.Checked -or $chkDPI2.Checked -or $chkTipoImp.Checked

    if ($needsFolder -and $comboBox.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Selecione uma pasta SACS na lista.",
            "Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $nomeSubpasta = if ($comboBox.SelectedIndex -ge 0) {
        $comboBox.SelectedItem.ToString().Trim().ToLower()
    } else { "" }
    $caminhoNovo = if ($nomeSubpasta) { Join-Path 'C:\' $nomeSubpasta } else { $null }

    # 1) COD_HEMO
    if ($chkCodHemo.Checked -and $txtCodHemo.Text.Trim()) {
        $iniDA = Join-Path $caminhoNovo '_data_access.ini'
        if (Test-Path $iniDA) {
            $lines = Get-Content $iniDA
            for ($j=0; $j -lt $lines.Count; $j++) {
                if ($lines[$j] -match '^N =\[_cod_hem\]=') {
                    $lines[$j]  = "N =[_cod_hem]= $($txtCodHemo.Text.Trim())"
                    $anyChanged = $true
                }
            }
            Set-Content -Path $iniDA -Value $lines
        }
    }

    # 2) BootStrap
    if ($chkBootstrap.Checked) {
        $iniWU = Join-Path (Join-Path $caminhoNovo 'bootstrap') 'WebUpdate.ini'
        if (Test-Path $iniWU) {
            $lines = Get-Content $iniWU
            for ($i=0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^PathLocalRoot=C:\\[^\\]+\\update\\?$') {
                    $lines[$i]  = "PathLocalRoot=C:\$nomeSubpasta\update"
                    $anyChanged = $true
                }
                if ($lines[$i] -match '^FinalApp="C:\\[^\\]+\\atalhos\\') {
                    $pattern     = '^FinalApp="C:\\[^\\]+\\atalhos\\'
                    $replacement = 'FinalApp="C:\' + $nomeSubpasta + '\atalhos\'
                    $lines[$i]   = $lines[$i] -replace $pattern, $replacement
                    $anyChanged = $true
                }
            }
            Set-Content -Path $iniWU -Value $lines
        }
    }

    # 3) Registrar OCX
    if ($chkRegistrarOCX.Checked) {
        $ocxNames  = @('ezVidC60.ocx','ezVidCap.ocx','MSCOMCTL.OCX','mci32.ocx','mschrt20.ocx','mscomct2.ocx','mscal.ocx')
        $sacsRoots = Get-ChildItem -Path 'C:\' -Directory -Filter 'SACS*' -ErrorAction SilentlyContinue
        $lblProgress.Visible = $true; $progressBar.Visible = $true
        $progressBar.Minimum = 0; $progressBar.Maximum = $ocxNames.Count; $progressBar.Value = 0
        foreach ($ocx in $ocxNames) {
            foreach ($root in $sacsRoots) {
                $src = Get-ChildItem -Path $root.FullName -Filter $ocx -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($src) {
                    Copy-Item -Path $src.FullName -Destination (Join-Path $env:SystemRoot\System32 $ocx) -Force
                    Copy-Item -Path $src.FullName -Destination (Join-Path $env:SystemRoot\SysWOW64 $ocx) -Force
                    Start-Process -FilePath 'regsvr32.exe' -ArgumentList '/s', "`"$(Join-Path $env:SystemRoot\System32 $ocx)`"" -Wait
                    Start-Process -FilePath (Join-Path $env:SystemRoot 'SysWOW64\regsvr32.exe') -ArgumentList '/s', "`"$(Join-Path $env:SystemRoot\SysWOW64 $ocx)`"" -Wait
                    $anyChanged = $true
                    break
                }
            }
            $progressBar.Value++
            [System.Windows.Forms.Application]::DoEvents()
        }
        $lblProgress.Visible = $false; $progressBar.Visible = $false
    }

    # 4) Instalar Fontes
    if ($chkFontes.Checked) {
        $fontsAppData = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
        if (-not (Test-Path $fontsAppData)) { New-Item -ItemType Directory -Path $fontsAppData | Out-Null }
        $regPath    = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
        $sacsRoots  = Get-ChildItem -Path 'C:\' -Directory -Filter 'SACS*' -ErrorAction SilentlyContinue
        foreach ($root in $sacsRoots) {
            Get-ChildItem -Path $root.FullName -Filter '*.ttf' -Recurse -File | ForEach-Object {
                $dest        = Join-Path $fontsAppData $_.Name
                Copy-Item -Path $_.FullName -Destination $dest -Force
                $fontName    = $_.BaseName + ' (TrueType)'
                Set-ItemProperty -Path $regPath -Name $fontName -Value $_.Name -Force
                $fontsInstalled = $true
                $anyChanged     = $true
            }
        }
    }

    # 5) Impressoras / INSTITUIÇÃO / DPI / Tipo Impressora
    if ($needsFolder) {
        $configIni = Join-Path $caminhoNovo 'configuracao.ini'
        if (Test-Path $configIni) {
            $cfgLines = Get-Content $configIni
            if ($chkPrinterGraphics.Checked) {
                $printerG = $comboPrinterGraphics.SelectedItem
                if ([string]::IsNullOrEmpty($printerG)) {
                    $resp = [System.Windows.Forms.MessageBox]::Show(
                        "Nenhuma impressora gráfica selecionada. Deixar o parâmetro sem valor no arquivo de configuração?",
                        "Confirmação",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                    $printerG = ''
                }
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[FICHA_DOADOR\]\s*=') { $cfgLines[$i] = "[FICHA_DOADOR]   = $printerG"; $anyChanged = $true }
                    if ($cfgLines[$i] -match '^\[FICHA_REDOME\]\s*=') { $cfgLines[$i] = "[FICHA_REDOME]   = $printerG"; $anyChanged = $true }
                }
            }
            if ($chkPrinterEtiqueta.Checked) {
                $printerE = $comboPrinterEtiqueta.SelectedItem
                if ([string]::IsNullOrEmpty($printerE)) {
                    $resp = [System.Windows.Forms.MessageBox]::Show(
                        "Nenhuma impressora de etiqueta selecionada. Deixar o parâmetro sem valor no arquivo de configuração?",
                        "Confirmação",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                    $printerE = ''
                }
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[BARCODE_DOADOR\]\s*=') { $cfgLines[$i] = "[BARCODE_DOADOR] = $printerE"; $anyChanged = $true }
                    if ($cfgLines[$i] -match '^\[BARCODE_GERAL\]\s*=/')   { $cfgLines[$i] = "[BARCODE_GERAL]  = $printerE"; $anyChanged = $true }
                }
            }
            if ($chkInstituicao.Checked -and $txtInstituicao.Text.Trim()) {
                $instVal = $txtInstituicao.Text.Trim()
                if ($cfgLines -match '^\[INSTITUICAO\]\s*=') {
                    for ($i=0; $i -lt $cfgLines.Count; $i++) {
                        if ($cfgLines[$i] -match '^\[INSTITUICAO\]\s*=') {
                            $cfgLines[$i] = "[INSTITUICAO]    = $instVal"; $anyChanged = $true
                        }
                    }
                } else {
                    $cfgLines += "[INSTITUICAO]    = $instVal"; $anyChanged = $true
                }
            }
            if ($chkDPI2.Checked -and $txtDPI2.Text.Trim()) {
                $dpiVal = $txtDPI2.Text.Trim()
                if ($cfgLines -match '^\[DPI\]\s*=') {
                    for ($i=0; $i -lt $cfgLines.Count; $i++) {
                        if ($cfgLines[$i] -match '^\[DPI\]\s*=') {
                            $cfgLines[$i] = "[DPI]            = $dpiVal"; $anyChanged = $true
                        }
                    }
                } else {
                    $cfgLines += "[DPI]            = $dpiVal"; $anyChanged = $true
                }
            }
            if ($chkTipoImp.Checked -and $comboTipoImp.SelectedIndex -ge 0) {
                $tipoVal = if ($comboTipoImp.SelectedItem -like "1*") { "1" } else { "2" }
                if ($cfgLines -match '^\[TIPO_IMPRESSORA\]\s*=') {
                    for ($i=0; $i -lt $cfgLines.Count; $i++) {
                        if ($cfgLines[$i] -match '^\[TIPO_IMPRESSORA\]\s*=') {
                            $cfgLines[$i] = "[TIPO_IMPRESSORA] = $tipoVal"; $anyChanged = $true
                        }
                    }
                } else {
                    $cfgLines += "[TIPO_IMPRESSORA] = $tipoVal"; $anyChanged = $true
                }
            }
            if ($anyChanged) { Set-Content -Path $configIni -Value $cfgLines }
        }
    }

    # 6) Atalhos
    if ($chkAtalhos.Checked) {
        $pathAt = Join-Path $caminhoNovo 'atalhos'
        if (Test-Path $pathAt) {
            $atalhos = Get-ChildItem -Path $pathAt -Recurse -Filter '*.lnk' -File
            foreach ($a in $atalhos) {
                try {
                    $shell = New-Object -ComObject WScript.Shell
                    $s     = $shell.CreateShortcut($a.FullName)
                    $mudou = $false
                    foreach ($campo in 'TargetPath','WorkingDirectory') {
                        $orig = $s.$campo
                        if ($orig -like 'C:\*') {
                            $parts = $orig -split '\\',4
                            if ($parts.Count -ge 3) {
                                $parts[1] = $nomeSubpasta
                                $novo     = $parts -join '\'
                                if ($s.$campo -ne $novo) {
                                    $s.$campo = $novo
                                    $mudou    = $true
                                    $anyChanged = $true
                                }
                            }
                        }
                    }
                    if ($mudou) { $s.Save() }
                } catch {}
            }
        }
    }

    # exibe alertas finais
    if ($fontsInstalled) {
        [System.Windows.Forms.MessageBox]::Show("Fontes instaladas com sucesso!","Fontes",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
    elseif ($anyChanged) {
        [System.Windows.Forms.MessageBox]::Show("Os parâmetros selecionados foram atualizados com sucesso!","Configuração",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
})



#----------------------------------- --- ABA "Copiar exe" ----------------------

# Label Origem
$lblOrigem = New-Object System.Windows.Forms.Label
$lblOrigem.Text     = "Pasta de origem:"
$lblOrigem.AutoSize = $true
$lblOrigem.Location = New-Object System.Drawing.Point(10, 20)
$tabPageCopiaExe.Controls.Add($lblOrigem)

# TextBox Origem
$txtOrigem = New-Object System.Windows.Forms.TextBox
$txtOrigem.Size     = New-Object System.Drawing.Size(300, 20)
$txtOrigem.Location = New-Object System.Drawing.Point(110, 18)
$tabPageCopiaExe.Controls.Add($txtOrigem)

# Botão Browse Origem
$btnBrowseOrigem = New-Object System.Windows.Forms.Button
$btnBrowseOrigem.Text     = "…"
$btnBrowseOrigem.Size     = New-Object System.Drawing.Size(30, 20)
$btnBrowseOrigem.Location = New-Object System.Drawing.Point(420, 18)
$btnBrowseOrigem.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOrigem.Text = $dlg.SelectedPath
    }
})
$tabPageCopiaExe.Controls.Add($btnBrowseOrigem)

# Label Destino
$lblDestino = New-Object System.Windows.Forms.Label
$lblDestino.Text     = "Pasta de destino:"
$lblDestino.AutoSize = $true
$lblDestino.Location = New-Object System.Drawing.Point(10, 60)
$tabPageCopiaExe.Controls.Add($lblDestino)

# TextBox Destino
$txtDestino = New-Object System.Windows.Forms.TextBox
$txtDestino.Size     = New-Object System.Drawing.Size(300, 20)
$txtDestino.Location = New-Object System.Drawing.Point(110, 58)
$tabPageCopiaExe.Controls.Add($txtDestino)

# Botão Browse Destino
$btnBrowseDestino = New-Object System.Windows.Forms.Button
$btnBrowseDestino.Text     = "…"
$btnBrowseDestino.Size     = New-Object System.Drawing.Size(30, 20)
$btnBrowseDestino.Location = New-Object System.Drawing.Point(420, 58)
$btnBrowseDestino.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDestino.Text = $dlg.SelectedPath
    }
})
$tabPageCopiaExe.Controls.Add($btnBrowseDestino)

# Botão “Copiar EXE” – agora usando o gerenciador do Windows
$btnCopiar = New-Object System.Windows.Forms.Button
$btnCopiar.Text     = "Copiar EXE"
$btnCopiar.Size     = New-Object System.Drawing.Size(440, 30)
$btnCopiar.Location = New-Object System.Drawing.Point(10, 100)
$btnCopiar.Add_Click({
    # Valida campos
    if ([string]::IsNullOrWhiteSpace($txtOrigem.Text) -or
        [string]::IsNullOrWhiteSpace($txtDestino.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Informe os caminhos de Origem e Destino.","Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Normaliza e verifica se são iguais
    $src  = $txtOrigem.Text.TrimEnd('\')
    $dest = $txtDestino.Text.TrimEnd('\')
    if ($src -ieq $dest) {
        [System.Windows.Forms.MessageBox]::Show(
            "Origem e Destino não podem ser o mesmo caminho.","Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Verifica existência de pasta
    if (-not (Test-Path $src) -or -not (Test-Path $dest)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Origem ou Destino inválido(s).","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Copia todos os .exe de uma vez, exibindo progresso e sobrescrevendo
    $shell      = New-Object -ComObject Shell.Application
    $destFolder = $shell.NameSpace($dest)

    # 0x10 = FOF_NOCONFIRMATION (substituir sem perguntar)
    $flags = 0x10

    try {
        # usa wildcard para agrupar tudo num único diálogo
        $destFolder.CopyHere("$src\*.exe", $flags)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erro ao copiar arquivos:`n$($_.Exception.Message)","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})
$tabPageCopiaExe.Controls.Add($btnCopiar)









# ---------------------- ABA SOBRE ----------------------
$aboutText = New-Object System.Windows.Forms.TextBox
$aboutText.Multiline        = $true
$aboutText.ReadOnly         = $true
$aboutText.ScrollBars       = 'Vertical'
$aboutText.WordWrap         = $true
$aboutText.Size             = New-Object System.Drawing.Size(460,240)
$aboutText.Location         = New-Object System.Drawing.Point(10,10)
$aboutText.Text     = @"
Instruções de uso:

- Aba Atualizador: informe o nome da pasta que contém os atalhos do Hemote Plus. O script corrige o caminho de destino e o diretório de trabalho. Em seguida, pergunta se deseja ajustar o WebUpdate.ini na pasta Bootstrap; se confirmado, faz o ajuste automático.

Permite alterar o COD_HEMO, no _data_access da pasta informada, ao ativar o parâmetro e digitando no campo de texto o valor desejado.

- Aba Bootstrap: permite, de forma independente, ajustar os parâmetros PathLocalRoot e FinalApp no arquivo WebUpdate.ini presente na pasta Bootstrap.

Versão: 1.6
Autor: Felipe Almeida
"@
$tabPage2.Controls.Add($aboutText)

# Exibe o formulário
[void]$form.ShowDialog()