Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Drawing.Printing

# ---------------------- FORMULÁRIO ----------------------
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Atualizador de Atalhos"
$form.ClientSize      = New-Object System.Drawing.Size(550,400)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $true

# ---------------------- TABS ----------------------
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabPage1        = New-Object System.Windows.Forms.TabPage("Atualizador")
$tabPageBootstrap= New-Object System.Windows.Forms.TabPage("Bootstrap")
$tabPageCopiaExe = New-Object System.Windows.Forms.TabPage("Copiar exe")
$tabPageSobre    = New-Object System.Windows.Forms.TabPage("Sobre")
$tabControl.TabPages.AddRange(@($tabPage1,$tabPageBootstrap,$tabPageCopiaExe,$tabPageSobre))
$form.Controls.Add($tabControl)

# ---------------------- ABA “Atualizador” ----------------------

# instrução
$label = New-Object System.Windows.Forms.Label
$label.Text     = "Selecione a pasta SACS:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,20)
$tabPage1.Controls.Add($label)

# ComboBox pastas SACS
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Location      = New-Object System.Drawing.Point(10,50)
$comboBox.Size          = New-Object System.Drawing.Size(350,20)
$comboBox.SelectedIndex = -1
$tabPage1.Controls.Add($comboBox)

# botão Atualizar
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text     = "Atualizar"
$updateButton.Location = New-Object System.Drawing.Point(370,48)
$tabPage1.Controls.Add($updateButton)

# COD_HEMO
$chkCodHemo = New-Object System.Windows.Forms.CheckBox
$chkCodHemo.Location = New-Object System.Drawing.Point(10,90)
$chkCodHemo.Size     = New-Object System.Drawing.Size(20,20)
$chkCodHemo.Add_CheckedChanged({
    $txtCodHemo.Enabled = $chkCodHemo.Checked
    if (-not $chkCodHemo.Checked) { $txtCodHemo.Clear() }
})
$tabPage1.Controls.Add($chkCodHemo)

$lblCodHemo = New-Object System.Windows.Forms.Label
$lblCodHemo.Text     = "COD_HEMO"
$lblCodHemo.AutoSize = $true
$lblCodHemo.Location = New-Object System.Drawing.Point(30,92)
$tabPage1.Controls.Add($lblCodHemo)

$txtCodHemo = New-Object System.Windows.Forms.TextBox
$txtCodHemo.Location  = New-Object System.Drawing.Point(110,90)
$txtCodHemo.Size      = New-Object System.Drawing.Size(100,20)
$txtCodHemo.MaxLength = 15
$txtCodHemo.Enabled   = $false
$tabPage1.Controls.Add($txtCodHemo)

# Atalhos
$chkAtalhos = New-Object System.Windows.Forms.CheckBox
$chkAtalhos.Location = New-Object System.Drawing.Point(220,90)
$chkAtalhos.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkAtalhos)

$lblAtalhos = New-Object System.Windows.Forms.Label
$lblAtalhos.Text     = "Atalhos"
$lblAtalhos.AutoSize = $true
$lblAtalhos.Location = New-Object System.Drawing.Point(240,92)
$tabPage1.Controls.Add($lblAtalhos)

# BootStrap
$chkBootstrap = New-Object System.Windows.Forms.CheckBox
$chkBootstrap.Location = New-Object System.Drawing.Point(320,90)
$chkBootstrap.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkBootstrap)

$lblBootstrap = New-Object System.Windows.Forms.Label
$lblBootstrap.Text     = "BootStrap"
$lblBootstrap.AutoSize = $true
$lblBootstrap.Location = New-Object System.Drawing.Point(340,92)
$tabPage1.Controls.Add($lblBootstrap)

# Registrar OCX
$chkRegistrarOCX = New-Object System.Windows.Forms.CheckBox
$chkRegistrarOCX.Location = New-Object System.Drawing.Point(420,90)
$chkRegistrarOCX.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkRegistrarOCX)

$lblRegistrarOCX = New-Object System.Windows.Forms.Label
$lblRegistrarOCX.Text     = "Registrar OCX"
$lblRegistrarOCX.AutoSize = $true
$lblRegistrarOCX.Location = New-Object System.Drawing.Point(440,92)
$tabPage1.Controls.Add($lblRegistrarOCX)

# Impressora Gráfica (linha própria)
$chkPrinterGraphics = New-Object System.Windows.Forms.CheckBox
$chkPrinterGraphics.Location = New-Object System.Drawing.Point(10,130)
$chkPrinterGraphics.Size     = New-Object System.Drawing.Size(20,20)
$chkPrinterGraphics.Add_CheckedChanged({
    $comboPrinterGraphics.Enabled = $chkPrinterGraphics.Checked
    if (-not $chkPrinterGraphics.Checked) { $comboPrinterGraphics.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkPrinterGraphics)

$lblPrinterGraphics = New-Object System.Windows.Forms.Label
$lblPrinterGraphics.Text     = "Impressora Gráfica"
$lblPrinterGraphics.AutoSize = $true
$lblPrinterGraphics.Location = New-Object System.Drawing.Point(30,132)
$tabPage1.Controls.Add($lblPrinterGraphics)

$comboPrinterGraphics = New-Object System.Windows.Forms.ComboBox
$comboPrinterGraphics.DropDownStyle = 'DropDownList'
$comboPrinterGraphics.Location      = New-Object System.Drawing.Point(150,130)
$comboPrinterGraphics.Size          = New-Object System.Drawing.Size(370,20)
$comboPrinterGraphics.Enabled       = $false
$tabPage1.Controls.Add($comboPrinterGraphics)

# Impressora Etiqueta (linha própria abaixo)
$chkPrinterEtiqueta = New-Object System.Windows.Forms.CheckBox
$chkPrinterEtiqueta.Location = New-Object System.Drawing.Point(10,160)
$chkPrinterEtiqueta.Size     = New-Object System.Drawing.Size(20,20)
$chkPrinterEtiqueta.Add_CheckedChanged({
    $comboPrinterEtiqueta.Enabled = $chkPrinterEtiqueta.Checked
    if (-not $chkPrinterEtiqueta.Checked) { $comboPrinterEtiqueta.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkPrinterEtiqueta)

$lblPrinterEtiqueta = New-Object System.Windows.Forms.Label
$lblPrinterEtiqueta.Text     = "Impressora Etiqueta"
$lblPrinterEtiqueta.AutoSize = $true
$lblPrinterEtiqueta.Location = New-Object System.Drawing.Point(30,162)
$tabPage1.Controls.Add($lblPrinterEtiqueta)

$comboPrinterEtiqueta = New-Object System.Windows.Forms.ComboBox
$comboPrinterEtiqueta.DropDownStyle = 'DropDownList'
$comboPrinterEtiqueta.Location      = New-Object System.Drawing.Point(150,160)
$comboPrinterEtiqueta.Size          = New-Object System.Drawing.Size(370,20)
$comboPrinterEtiqueta.Enabled       = $false
$tabPage1.Controls.Add($comboPrinterEtiqueta)


# Progress bar and label (rodapé elevado)
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text     = "Aguarde, registrando as OCX."
$lblProgress.AutoSize = $true
$lblProgress.Location = New-Object System.Drawing.Point(10,280)
$lblProgress.Visible  = $false
$tabPage1.Controls.Add($lblProgress)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,300)
$progressBar.Size     = New-Object System.Drawing.Size(530,20)
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

# exibe form e popula listas
$form.Add_Shown({
    Populate-SacsFolders
    Populate-Printers
})

# ---------------------- CLIQUE “Atualizar” ----------------------
$updateButton.Add_Click({
   # 0) valida pasta SACS para todos os itens que precisam dela
    $needsFolder = $chkCodHemo.Checked `
                 -or $chkAtalhos.Checked `
                 -or $chkBootstrap.Checked `
                 -or $chkPrinterGraphics.Checked `
                 -or $chkPrinterEtiqueta.Checked

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
    if ($chkCodHemo.Checked) {
        $iniDA = Join-Path $caminhoNovo '_data_access.ini'
        if (-not (Test-Path $iniDA)) {
            [System.Windows.Forms.MessageBox]::Show(
                "_data_access.ini não encontrado em '$caminhoNovo'.",
                "Aviso",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        $lines = Get-Content $iniDA
        $found = $lines | Where-Object { $_ -match '^N =\[_cod_hem\]=' }
        if (-not $found) {
            [System.Windows.Forms.MessageBox]::Show(
                "Parâmetro 'N =[_cod_hem]=' não encontrado ou incorreto.",
                "Erro",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
        for ($j=0; $j -lt $lines.Count; $j++) {
            if ($lines[$j] -match '^N =\[_cod_hem\]=') {
                $lines[$j] = "N =[_cod_hem]= $($txtCodHemo.Text.Trim())"
            }
        }
        Set-Content -Path $iniDA -Value $lines
    }

    # 2) BootStrap
    if ($chkBootstrap.Checked) {
        $iniWU = Join-Path (Join-Path $caminhoNovo 'bootstrap') 'WebUpdate.ini'
        if (-not (Test-Path $iniWU)) {
            [System.Windows.Forms.MessageBox]::Show(
                "WebUpdate.ini não encontrado em '$nomeSubpasta\bootstrap'.",
                "Aviso",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        $lines = Get-Content $iniWU
        for ($i=0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^PathLocalRoot=C:\\[^\\]+\\update\\?$') {
                $lines[$i] = "PathLocalRoot=C:\$nomeSubpasta\update"
            }
            if ($lines[$i] -match '^FinalApp="C:\\[^\\]+\\atalhos\\') {
                $pattern     = '^FinalApp="C:\\[^\\]+\\atalhos\\'
                $replacement = 'FinalApp="C:\' + $nomeSubpasta + '\atalhos\'
                $lines[$i]   = $lines[$i] -replace $pattern, $replacement
            }
        }
        Set-Content -Path $iniWU -Value $lines
    }

    # 3) Registrar OCX
    if ($chkRegistrarOCX.Checked) {
        $ocxNames   = @(
            'ezVidC60.ocx','ezVidCap.ocx','MSCOMCTL.OCX',
            'mci32.ocx','mschrt20.ocx','mscomct2.ocx','mscal.ocx'
        )
        $sacsRoots  = Get-ChildItem -Path 'C:\' -Directory -Filter 'SACS*' -ErrorAction SilentlyContinue
        $notFound   = @()
        $registered = @()

        # exibe progress
        $lblProgress.Visible = $true
        $progressBar.Visible = $true
        $progressBar.Minimum = 0
        $progressBar.Maximum = $ocxNames.Count
        $progressBar.Value   = 0

        foreach ($ocx in $ocxNames) {
            $foundOne = $false
            foreach ($root in $sacsRoots) {
                $src = Get-ChildItem -Path $root.FullName -Filter $ocx -Recurse -File -ErrorAction SilentlyContinue |
                       Select-Object -First 1
                if ($src) {
                    $dest32  = Join-Path $env:SystemRoot\System32  $ocx
                    $destWow = Join-Path $env:SystemRoot\SysWOW64 $ocx
                    Copy-Item -Path $src.FullName -Destination $dest32  -Force
                    Copy-Item -Path $src.FullName -Destination $destWow -Force
                    Start-Process -FilePath 'regsvr32.exe' `
                                  -ArgumentList '/s', "`"$dest32`"" -Wait
                    Start-Process -FilePath (Join-Path $env:SystemRoot 'SysWOW64\regsvr32.exe') `
                                  -ArgumentList '/s', "`"$destWow`"" -Wait
                    $foundOne   = $true
                    $registered += $ocx
                    break
                }
            }
            if (-not $foundOne) { $notFound += $ocx }
            $progressBar.Value += 1
            [System.Windows.Forms.Application]::DoEvents()
        }

        # oculta progress
        $lblProgress.Visible = $false
        $progressBar.Visible = $false

        if ($registered.Count) {
            [System.Windows.Forms.MessageBox]::Show(
                "OCX registradas:`n`n$($registered -join "`n")",
                "Registrar OCX",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        if ($notFound.Count) {
            [System.Windows.Forms.MessageBox]::Show(
                "OCX não registradas:`n`n$($notFound -join "`n")",
                "Registrar OCX",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    }

   # 4) Configurar impressoras em configuracao.ini
    if ($chkPrinterGraphics.Checked -or $chkPrinterEtiqueta.Checked) {
        $configIni = Join-Path $caminhoNovo 'configuracao.ini'
        if (-not (Test-Path $configIni)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Arquivo 'configuracao.ini' não encontrado em '$caminhoNovo'.",
                "Erro",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
        $cfgLines = Get-Content $configIni

        # Impressora Gráfica
        if ($chkPrinterGraphics.Checked) {
            # valida existência e formato dos parâmetros
            $missing = @()
            if (-not ($cfgLines -match '^\[FICHA_DOADOR\]\s*=')) { $missing += '[FICHA_DOADOR]' }
            if (-not ($cfgLines -match '^\[FICHA_REDOME\]\s*='))  { $missing += '[FICHA_REDOME]' }
            if ($missing.Count) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Parâmetros incorretos ou ausentes em configuracao.ini: $($missing -join ', ')",
                    "Erro",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            $printerG = $comboPrinterGraphics.SelectedItem
            if ([string]::IsNullOrEmpty($printerG)) {
                $resp = [System.Windows.Forms.MessageBox]::Show(
                    "Nenhuma impressora gráfica selecionada. Os parâmetros serão definidos em branco. Deseja continuar?",
                    "Confirmação",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                $printerG = ''
            }
            for ($i = 0; $i -lt $cfgLines.Count; $i++) {
                if ($cfgLines[$i] -match '^\[FICHA_DOADOR\]\s*=') {
                    $cfgLines[$i] = "[FICHA_DOADOR]   = $printerG"
                }
                if ($cfgLines[$i] -match '^\[FICHA_REDOME\]\s*=') {
                    $cfgLines[$i] = "[FICHA_REDOME]   = $printerG"
                }
            }
        }

        # Impressora Etiqueta
        if ($chkPrinterEtiqueta.Checked) {
            $missing = @()
            if (-not ($cfgLines -match '^\[BARCODE_DOADOR\]\s*=')) { $missing += '[BARCODE_DOADOR]' }
            if (-not ($cfgLines -match '^\[BARCODE_GERAL\]\s*='))  { $missing += '[BARCODE_GERAL]' }
            if ($missing.Count) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Parâmetros incorretos ou ausentes em configuracao.ini: $($missing -join ', ')",
                    "Erro",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            $printerE = $comboPrinterEtiqueta.SelectedItem
            if ([string]::IsNullOrEmpty($printerE)) {
                $resp = [System.Windows.Forms.MessageBox]::Show(
                    "Nenhuma impressora de etiqueta selecionada. Os parâmetros serão definidos em branco. Deseja continuar?",
                    "Confirmação",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($resp -ne [System.Windows.Forms.DialogResult]::Yes) { return }
                $printerE = ''
            }
            for ($i = 0; $i -lt $cfgLines.Count; $i++) {
                if ($cfgLines[$i] -match '^\[BARCODE_DOADOR\]\s*=') {
                    $cfgLines[$i] = "[BARCODE_DOADOR] = $printerE"
                }
                if ($cfgLines[$i] -match '^\[BARCODE_GERAL\]\s*=') {
                    $cfgLines[$i] = "[BARCODE_GERAL]  = $printerE"
                }
            }
        }

        # grava alterações e exibe confirmação
        Set-Content -Path $configIni -Value $cfgLines
        [System.Windows.Forms.MessageBox]::Show(
            "Os parâmetros selecionados foram atualizados com sucesso!",
            "Configuração",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }




    # 4) Atalhos
    if ($chkAtalhos.Checked) {
        $pathAt = Join-Path $caminhoNovo 'atalhos'
        if (-not (Test-Path $pathAt)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Pasta 'atalhos' não encontrada em '$caminhoNovo'.",
                "Atalhos",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        $atalhos = Get-ChildItem -Path $pathAt -Recurse -Filter '*.lnk' -File -ErrorAction SilentlyContinue
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
                            }
                        }
                    }
                }
                if ($mudou) { $s.Save() }
            } catch {}
        }
    }
})




# ---------------------- ABA “Sobre” ----------------------
$aboutText = New-Object System.Windows.Forms.TextBox
$aboutText.Multiline  = $true
$aboutText.ReadOnly   = $true
$aboutText.Dock       = [System.Windows.Forms.DockStyle]::Fill
$aboutText.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$aboutText.Text       = @"
Ajuste de Atalhos v1.1
Desenvolvido por HMG.
"@
$tabPageSobre.Controls.Add($aboutText)

# ---------------------- ABA BOOTSTRAP (manual) ----------------------
$bootstrapLabel = New-Object System.Windows.Forms.Label
$bootstrapLabel.Text     = "Selecione o arquivo WebUpdate.ini:"
$bootstrapLabel.AutoSize = $true
$bootstrapLabel.Location = New-Object System.Drawing.Point(10,20)
$tabPageBootstrap.Controls.Add($bootstrapLabel)

$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Filter = "INI files (*.ini)|*.ini"
$fileDialog.Title  = "Selecione o arquivo WebUpdate.ini"

$selectFileButton = New-Object System.Windows.Forms.Button
$selectFileButton.Text     = "Buscar"
$selectFileButton.Location = New-Object System.Drawing.Point(10,50)
$tabPageBootstrap.Controls.Add($selectFileButton)

$selectedFileLabel = New-Object System.Windows.Forms.Label
$selectedFileLabel.AutoSize = $true
$selectedFileLabel.Location = New-Object System.Drawing.Point(10,80)
$tabPageBootstrap.Controls.Add($selectedFileLabel)

$nomeLabel = New-Object System.Windows.Forms.Label
$nomeLabel.Text     = "Nome da pasta (ex: sacs_teste):"
$nomeLabel.AutoSize = $true
$nomeLabel.Location = New-Object System.Drawing.Point(10,110)
$tabPageBootstrap.Controls.Add($nomeLabel)

$nomeTextBox = New-Object System.Windows.Forms.TextBox
$nomeTextBox.Size     = New-Object System.Drawing.Size(300,20)
$nomeTextBox.Location = New-Object System.Drawing.Point(10,135)
$tabPageBootstrap.Controls.Add($nomeTextBox)

$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text     = "Aplicar"
$applyButton.Location = New-Object System.Drawing.Point(320,133)
$tabPageBootstrap.Controls.Add($applyButton)

$bootstrapStatus = New-Object System.Windows.Forms.TextBox
$bootstrapStatus.Multiline   = $true
$bootstrapStatus.ScrollBars  = 'Vertical'
$bootstrapStatus.ReadOnly    = $true
$bootstrapStatus.Size        = New-Object System.Drawing.Size(460,50)
$bootstrapStatus.Location    = New-Object System.Drawing.Point(10,170)
$tabPageBootstrap.Controls.Add($bootstrapStatus)

$selectFileButton.Add_Click({
    if ($fileDialog.ShowDialog() -eq 'OK') {
        $selectedFileLabel.Text = "Arquivo selecionado: $($fileDialog.FileName)"
    }
})

$applyButton.Add_Click({
    $arquivoIni = $fileDialog.FileName
    $novoNome   = $nomeTextBox.Text.Trim()
    if (-not (Test-Path $arquivoIni)) {
        [System.Windows.Forms.MessageBox]::Show("Selecione um arquivo válido.","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if ([string]::IsNullOrWhiteSpace($novoNome)) {
        [System.Windows.Forms.MessageBox]::Show("Informe o nome da pasta.","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $linhas   = Get-Content $arquivoIni
    $alterIni = 0
    for ($i = 0; $i -lt $linhas.Count; $i++) {
        if ($linhas[$i] -match '^PathLocalRoot=C:\\[^\\]+\\update\\?$') {
            $linhas[$i] = "PathLocalRoot=C:\$novoNome\update"
            $alterIni++
        }
        if ($linhas[$i] -match '^FinalApp="C:\\[^\\]+\\atalhos\\.*"$') {
            $linhas[$i] = $linhas[$i] -replace '^FinalApp="C:\\[^\\]+\\atalhos\\', "FinalApp=`"C:\$novoNome\atalhos\"
            $alterIni++
        }
    }
    if ($alterIni -gt 0) {
        Set-Content -Path $arquivoIni -Value $linhas
        $bootstrapStatus.Text = "✅ Alterações aplicadas: $alterIni linha(s)."
    } else {
        $bootstrapStatus.Text = "ℹ️ Nenhuma linha foi alterada."
    }
})

# ---------------------- ABA COPIAR EXE ----------------------

# Origem
$lblOrigem = New-Object System.Windows.Forms.Label
$lblOrigem.Text     = "Pasta de origem:"
$lblOrigem.AutoSize = $true
$lblOrigem.Location = New-Object System.Drawing.Point(10,20)
$tabPageCopiaExe.Controls.Add($lblOrigem)

$txtOrigem = New-Object System.Windows.Forms.TextBox
$txtOrigem.Size     = New-Object System.Drawing.Size(300,20)
$txtOrigem.Location = New-Object System.Drawing.Point(110,18)
$tabPageCopiaExe.Controls.Add($txtOrigem)

$btnBrowseOrigem = New-Object System.Windows.Forms.Button
$btnBrowseOrigem.Text     = "…"
$btnBrowseOrigem.Size     = New-Object System.Drawing.Size(30,20)
$btnBrowseOrigem.Location = New-Object System.Drawing.Point(420,18)
$btnBrowseOrigem.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOrigem.Text = $dlg.SelectedPath
    }
})
$tabPageCopiaExe.Controls.Add($btnBrowseOrigem)

# Destino
$lblDestino = New-Object System.Windows.Forms.Label
$lblDestino.Text     = "Pasta de destino:"
$lblDestino.AutoSize = $true
$lblDestino.Location = New-Object System.Drawing.Point(10,60)
$tabPageCopiaExe.Controls.Add($lblDestino)

$txtDestino = New-Object System.Windows.Forms.TextBox
$txtDestino.Size     = New-Object System.Drawing.Size(300,20)
$txtDestino.Location = New-Object System.Drawing.Point(110,58)
$tabPageCopiaExe.Controls.Add($txtDestino)

$btnBrowseDestino = New-Object System.Windows.Forms.Button
$btnBrowseDestino.Text     = "…"
$btnBrowseDestino.Size     = New-Object System.Drawing.Size(30,20)
$btnBrowseDestino.Location = New-Object System.Drawing.Point(420,58)
$btnBrowseDestino.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDestino.Text = $dlg.SelectedPath
    }
})
$tabPageCopiaExe.Controls.Add($btnBrowseDestino)

# Botão Copiar EXE
$btnCopiar = New-Object System.Windows.Forms.Button
$btnCopiar.Text     = "Copiar EXE"
$btnCopiar.Size     = New-Object System.Drawing.Size(440,30)
$btnCopiar.Location = New-Object System.Drawing.Point(10,100)
$btnCopiar.Add_Click({
    # Verifica campos preenchidos
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

    # Valida existência de pastas
    if (-not (Test-Path $src) -or -not (Test-Path $dest)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Origem ou Destino inválido(s).","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Coleta arquivos .exe na raiz
    $files = Get-ChildItem -Path $src -Filter *.exe -File
    if ($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Nenhum .exe encontrado em:`n$src","Aviso",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    # Cria janela de progresso modal centralizada
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Size            = New-Object System.Drawing.Size(300,100)
    $progressForm.FormBorderStyle = 'FixedToolWindow'
    $progressForm.StartPosition   = 'CenterScreen'
    $progressForm.TopMost         = $true
    $progressForm.ControlBox      = $false

    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text     = "Copiando executáveis..."
    $lblMsg.AutoSize = $true
    $lblMsg.Location = New-Object System.Drawing.Point(10,10)
    $progressForm.Controls.Add($lblMsg)

    $prg = New-Object System.Windows.Forms.ProgressBar
    $prg.Location  = New-Object System.Drawing.Point(10,40)
    $prg.Size      = New-Object System.Drawing.Size(260,20)
    $prg.Minimum   = 0
    $prg.Maximum   = $files.Count
    $prg.Value     = 0
    $progressForm.Controls.Add($prg)

    # Exibe form de progresso modeless para permitir DoEvents()
    $progressForm.Show($form)
    [System.Windows.Forms.Application]::DoEvents()

    # Loop de cópia e atualização da barra
    $copiados = 0
    foreach ($file in $files) {
        try {
            Copy-Item -Path $file.FullName `
                      -Destination (Join-Path $dest $file.Name) `
                      -Force -ErrorAction Stop
            $copiados++
        } catch { }
        $prg.Value = $copiados
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Fecha a janela de progresso
    $progressForm.Close()

    # Mensagem final apenas com total de .exe copiados
    [System.Windows.Forms.MessageBox]::Show(
        "$copiados arquivo(s) copiado(s) com sucesso.","Sucesso",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})
$tabPageCopiaExe.Controls.Add($btnCopiar)


# Anexa o TabControl ao formulário e mostra
$form.Controls.Add($tabControl)




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