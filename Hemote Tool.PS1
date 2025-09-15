try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing")      | Out-Null
}
[System.Windows.Forms.Application]::EnableVisualStyles()


$windowsIdentity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($windowsIdentity)
$isAdmin          = $windowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
$defaultSacs      = 'SACS'

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Hemote Tool"
$form.ClientSize      = New-Object System.Drawing.Size(580,350)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $true

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

$tabPage1     = New-Object System.Windows.Forms.TabPage("Configurações")
$tabPageSobre = New-Object System.Windows.Forms.TabPage("Sobre")

$tabControl.TabPages.AddRange(@($tabPage1, $tabPageSobre))
$form.Controls.Add($tabControl)

# Aba Configurações – SACS e Ações
$lblSacs = New-Object System.Windows.Forms.Label
$lblSacs.Text     = "Selecione a pasta:"
$lblSacs.AutoSize = $true
$lblSacs.Location = New-Object System.Drawing.Point(10,20)
$tabPage1.Controls.Add($lblSacs)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Location      = New-Object System.Drawing.Point(150,18)
$comboBox.Size          = New-Object System.Drawing.Size(350,20)
$tabPage1.Controls.Add($comboBox)

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
$chkRegistrarOCX.Location = New-Object System.Drawing.Point(230,60)
$chkRegistrarOCX.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkRegistrarOCX)
$lblRegistrarOCX = New-Object System.Windows.Forms.Label
$lblRegistrarOCX.Text     = "Registrar OCX"
$lblRegistrarOCX.AutoSize = $true
$lblRegistrarOCX.Location = New-Object System.Drawing.Point(250,62)
$tabPage1.Controls.Add($lblRegistrarOCX)

$chkFontes = New-Object System.Windows.Forms.CheckBox
$chkFontes.Location = New-Object System.Drawing.Point(350,60)
$chkFontes.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkFontes)
$lblFontes = New-Object System.Windows.Forms.Label
$lblFontes.Text     = "Instalar Fontes"
$lblFontes.AutoSize = $true
$lblFontes.Location = New-Object System.Drawing.Point(370,62)
$tabPage1.Controls.Add($lblFontes)

# Aba Configurações – COD_HEMO e INSTITUIÇÃO
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

# Aba Configurações – Impressoras, DPI e atualização
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
$comboTipoImp.Items.AddRange(@("1 - ALEGRO","2 - ZEBRA"))
$comboTipoImp.Location      = New-Object System.Drawing.Point(460,140)
$comboTipoImp.Size          = New-Object System.Drawing.Size(100,20)
$comboTipoImp.Enabled       = $false
$tabPage1.Controls.Add($comboTipoImp)

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

$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text     = "Atualizar"
$updateButton.Size     = New-Object System.Drawing.Size(100,30)
$updateButton.Location = New-Object System.Drawing.Point(464,275)
$tabPage1.Controls.Add($updateButton)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text     = "Aguarde, registrando as OCX."
$lblProgress.AutoSize = $true
$lblProgress.Location = New-Object System.Drawing.Point(10,215)
$lblProgress.Visible  = $false
$tabPage1.Controls.Add($lblProgress)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,235)
$progressBar.Size     = New-Object System.Drawing.Size(550,20)
$progressBar.Minimum  = 0
$progressBar.Value    = 0
$progressBar.Visible  = $false
$tabPage1.Controls.Add($progressBar)

# Aba Sobre – descrição da ferramenta
$txtSobre = New-Object System.Windows.Forms.TextBox
$txtSobre.Multiline   = $true
$txtSobre.ReadOnly    = $true
$txtSobre.ScrollBars  = 'Vertical'
$txtSobre.Dock        = [System.Windows.Forms.DockStyle]::Fill
$txtSobre.Text = @"


Bem-vindo ao Hemote Tool

Desenvolvido para acelerar tarefas de rotina, reduzir erros, retrabalhos, otimizar o tempo de resposta e elevar a qualidade das entregas. o Hemote Tool automatiza:
  
  • Atualização de atalhos;
  • Ajustes de parâmetros (COD_HEMO, Instituição, DPI, Tipo de Impressora);
  • Configuração de impressoras (etiqueta e gráfica);
  • Registro de componentes OCX;
  • Instalação de fontes  







Versão: 2.0
Autor: Felipe Almeida

"@
$tabPageSobre.Controls.Clear()
$tabPageSobre.Controls.Add($txtSobre)

# ToolTips para esclarecimento das opções
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay  = 5000
$toolTip.InitialDelay  = 500
$toolTip.ReshowDelay   = 500
$toolTip.ShowAlways    = $true

$toolTip.SetToolTip($lblSacs,            "Escolha a pasta SACS onde deseja realizar as configurações selecionadas abaixo.")
$toolTip.SetToolTip($lblAtalhos,         "Marque para atualizar os atalhos do Hemote Plus.")
$toolTip.SetToolTip($lblBootstrap,       "Marque para atualizar o caminho do WebUpdate.ini, no Bootstrap.")
$toolTip.SetToolTip($lblRegistrarOCX,    "Marque para registrar as OCX no Windows.")
$toolTip.SetToolTip($lblFontes,          "Marque para instalar fontes necessárias ao sistema.")
$toolTip.SetToolTip($lblCodHemo,         "Informe o COD_HEMO que será definido no arquivo _data_access.ini.")
$toolTip.SetToolTip($lblInstituicao,     "Informe o nome da instituição que será definido no arquivo configuracao.ini.")
$toolTip.SetToolTip($lblPrinterEtiqueta, "Define a impressora de etiqueta selecionada, no arquivo de configuração.")
$toolTip.SetToolTip($lblTipoImp,         "Escolha o tipo de etiqueta: 1 - ALEGRO ou 2 - ZEBRA.")
$toolTip.SetToolTip($lblPrinterGraphics, "Define a impressora gráfica selecionada, no arquivo de configuração.")
$toolTip.SetToolTip($lblDPI2,            "Informe a resolução (DPI) para impressões de etiqueta.")

# Preenche combos ao exibir o form
$form.Add_Shown({
    $comboBox.Items.Clear()
    Get-ChildItem -Path 'C:\' -Directory -Filter 'SACS*' |
      Sort-Object Name |
      ForEach-Object { $comboBox.Items.Add($_.Name) }

    if ($comboBox.Items.Contains($defaultSacs)) {
        $comboBox.SelectedItem = $defaultSacs
    } elseif ($comboBox.Items.Count -gt 0) {
        $comboBox.SelectedIndex = 0
    }

    $comboPrinterGraphics.Items.Clear()
    $comboPrinterEtiqueta.Items.Clear()
    foreach ($p in [System.Drawing.Printing.PrinterSettings]::InstalledPrinters) {
        $comboPrinterGraphics.Items.Add($p)
        $comboPrinterEtiqueta.Items.Add($p)
    }
})

$updateButton.Add_Click({
    $anyChanged     = $false
    $ocxRegistered  = $false
    $fontsInstalled = $false
    $paramsChanged  = $false

    $needsFolder = $chkAtalhos.Checked -or $chkBootstrap.Checked -or $chkFontes.Checked `
                 -or $chkCodHemo.Checked -or $chkInstituicao.Checked `
                 -or $chkPrinterGraphics.Checked -or $chkPrinterEtiqueta.Checked `
                 -or $chkDPI2.Checked -or $chkTipoImp.Checked

    if ($needsFolder -and $comboBox.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
          "Selecione uma pasta SACS na lista.","Erro",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $nomeSubpasta = if ($comboBox.SelectedIndex -ge 0) {
        $comboBox.SelectedItem.ToString().Trim().ToLower()
    } else { "" }
    $caminhoNovo  = if ($nomeSubpasta) { Join-Path 'C:\' $nomeSubpasta } else { $null }

    # 1) COD_HEMO
    if ($chkCodHemo.Checked -and $txtCodHemo.Text.Trim()) {
        $iniDA = Join-Path $caminhoNovo '_data_access.ini'
        if (Test-Path $iniDA) {
            $lines = Get-Content $iniDA
            for ($i=0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^N =\[_cod_hem\]=') {
                    $lines[$i]     = "N =[_cod_hem]= $($txtCodHemo.Text.Trim())"
                    $anyChanged    = $true; $paramsChanged = $true
                }
            }
            Set-Content -Path $iniDA -Value $lines
        }
    }

    # 2) BootStrap
    if ($chkBootstrap.Checked -and $caminhoNovo) {
        $iniWU = Join-Path (Join-Path $caminhoNovo 'bootstrap') 'WebUpdate.ini'
        if (Test-Path $iniWU) {
            $lines = Get-Content $iniWU
            for ($i=0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^PathLocalRoot=C:\\[^\\]+\\update\\?$') {
                    $lines[$i]  = "PathLocalRoot=C:\$nomeSubpasta\update"
                    $anyChanged    = $true; $paramsChanged = $true
                }
                if ($lines[$i] -match '^FinalApp="C:\\[^\\]+\\atalhos\\') {
                    $pattern     = '^FinalApp="C:\\[^\\]+\\atalhos\\'
                    $replacement = 'FinalApp="C:\' + $nomeSubpasta + '\atalhos\'
                    $lines[$i]   = $lines[$i] -replace $pattern, $replacement
                    $anyChanged  = $true; $paramsChanged = $true
                }
            }
            Set-Content -Path $iniWU -Value $lines
        }
    }

    # 3) Registrar OCX
    if ($chkRegistrarOCX.Checked) {
        if (-not $isAdmin) {
            [System.Windows.Forms.MessageBox]::Show(
              "Para instalação de Fontes e/ou OCX, execute a ferramenta como Administrador.",
              "Permissão",
              [System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $ocxNames  = @('ezVidC60.ocx','ezVidCap.ocx','MSCOMCTL.OCX',
                       'mci32.ocx','mschrt20.ocx','mscomct2.ocx','mscal.ocx')
        $sacsRoots = Get-ChildItem -Path 'C:\' -Directory -Filter 'SACS*' -ErrorAction SilentlyContinue
        $lblProgress.Visible = $true; $progressBar.Visible = $true
        $progressBar.Minimum = 0; $progressBar.Maximum = $ocxNames.Count; $progressBar.Value = 0

        foreach ($ocx in $ocxNames) {
            foreach ($root in $sacsRoots) {
                $src = Get-ChildItem -Path $root.FullName -Filter $ocx -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($src) {
                    Copy-Item -Path $src.FullName -Destination (Join-Path $env:SystemRoot\System32 $ocx) -Force -ErrorAction SilentlyContinue
                    Copy-Item -Path $src.FullName -Destination (Join-Path $env:SystemRoot\SysWOW64 $ocx) -Force -ErrorAction SilentlyContinue
                    Start-Process -FilePath 'regsvr32.exe' -ArgumentList '/s',(Join-Path $env:SystemRoot\System32 $ocx) -Wait -ErrorAction SilentlyContinue
                    Start-Process -FilePath (Join-Path $env:SystemRoot\SysWOW64 'regsvr32.exe') -ArgumentList '/s',(Join-Path $env:SystemRoot\SysWOW64 $ocx) -Wait -ErrorAction SilentlyContinue
                    $anyChanged    = $true; $ocxRegistered = $true
                    break
                }
            }
            $progressBar.Value++; [System.Windows.Forms.Application]::DoEvents()
        }
        $lblProgress.Visible = $false; $progressBar.Visible = $false
            }

    # 4) Instalar Fontes
    if ($chkFontes.Checked) {
        if (-not $isAdmin) {
            [System.Windows.Forms.MessageBox]::Show(
              "Para instalação de Fontes e/ou OCX, execute a ferramenta como Administrador.",
              "Permissão",
              [System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $fontsAppData = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
        if (-not (Test-Path $fontsAppData)) { New-Item -ItemType Directory -Path $fontsAppData | Out-Null }
        $regPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
        $sacsRoots = Get-ChildItem -Path 'C:\' -Directory -Filter 'SACS*' -ErrorAction SilentlyContinue
        $anyFontsInstalled = $false

        foreach ($root in $sacsRoots) {
            Get-ChildItem -Path $root.FullName -Filter '*.ttf' -Recurse -File | ForEach-Object {
                $dest = Join-Path $fontsAppData $_.Name
                if (-not (Test-Path $dest)) {
                    Copy-Item -Path $_.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
                    $fontName = $_.BaseName + ' (TrueType)'
                    if (-not (Get-ItemProperty -Path $regPath -Name $fontName -ErrorAction SilentlyContinue)) {
                        Set-ItemProperty -Path $regPath -Name $fontName -Value $_.Name -Force
                    }
                    $anyChanged      = $true
                    $fontsInstalled  = $true
                    $anyFontsInstalled = $true
                }
            }
        }

        if ($anyFontsInstalled) {
            [System.Windows.Forms.MessageBox]::Show(
              "Fontes instaladas com sucesso!",
              "Fontes",
              [System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
              "Não existem novas fontes a serem instaladas.",
              "Fontes",
              [System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }

    # 5) Impressoras / INSTITUIÇÃO / DPI / Tipo Impressora
    if ($needsFolder) {
        $configIni = Join-Path $caminhoNovo 'configuracao.ini'
        if (Test-Path $configIni) {
            $cfgLines = Get-Content $configIni

            if ($chkPrinterGraphics.Checked) {
                if ($comboPrinterGraphics.SelectedIndex -lt 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                      "Selecione uma impressora ou desmarque a opção.",
                      "Atenção",
                      [System.Windows.Forms.MessageBoxButtons]::OK,
                      [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
                $printerG = $comboPrinterGraphics.SelectedItem
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[FICHA_DOADOR\]\s*=') {
                        $cfgLines[$i] = "[FICHA_DOADOR]   = $printerG"; $anyChanged=$true; $paramsChanged=$true
                    }
                    if ($cfgLines[$i] -match '^\[FICHA_REDOME\]\s*=') {
                        $cfgLines[$i] = "[FICHA_REDOME]   = $printerG"; $anyChanged=$true; $paramsChanged=$true
                    }
                }
            }

            if ($chkPrinterEtiqueta.Checked) {
                if ($comboPrinterEtiqueta.SelectedIndex -lt 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                      "Selecione uma impressora ou desmarque a opção.",
                      "Atenção",
                      [System.Windows.Forms.MessageBoxButtons]::OK,
                      [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
                $printerE = $comboPrinterEtiqueta.SelectedItem
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[BARCODE_DOADOR\]\s*=') {
                        $cfgLines[$i] = "[BARCODE_DOADOR] = $printerE"; $anyChanged=$true; $paramsChanged=$true
                    }
                    if ($cfgLines[$i] -match '^\[BARCODE_GERAL\]\s*=') {
                        $cfgLines[$i] = "[BARCODE_GERAL]  = $printerE"; $anyChanged=$true; $paramsChanged=$true
                    }
                }
            }

            if ($chkInstituicao.Checked -and $txtInstituicao.Text.Trim()) {
                $instVal = $txtInstituicao.Text.Trim(); $found = $false
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[INSTITUICAO\]\s*=') {
                        $cfgLines[$i] = "[INSTITUICAO]    = $instVal"; $anyChanged=$true; $paramsChanged=$true; $found = $true
                    }
                }
                if (-not $found) {
                    $cfgLines += "[INSTITUICAO]    = $instVal"; $anyChanged=$true; $paramsChanged=$true
                }
            }

            if ($chkDPI2.Checked -and $txtDPI2.Text.Trim()) {
                $dpiVal = $txtDPI2.Text.Trim(); $found = $false
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[DPI\]\s*=') {
                        $cfgLines[$i] = "[DPI]            = $dpiVal"; $anyChanged=$true; $paramsChanged=$true; $found = $true
                    }
                }
                if (-not $found) {
                    $cfgLines += "[DPI]            = $dpiVal"; $anyChanged=$true; $paramsChanged=$true
                }
            }

            if ($chkTipoImp.Checked -and $comboTipoImp.SelectedIndex -ge 0) {
                $tipoVal = if ($comboTipoImp.SelectedItem -like "1*") { "1" } else { "2" }; $found = $false
                for ($i=0; $i -lt $cfgLines.Count; $i++) {
                    if ($cfgLines[$i] -match '^\[TIPO_IMPRESSORA\]\s*=') {
                        $cfgLines[$i] = "[TIPO_IMPRESSORA] = $tipoVal"; $anyChanged=$true; $paramsChanged=$true; $found = $true
                    }
                }
                if (-not $found) {
                    $cfgLines += "[TIPO_IMPRESSORA] = $tipoVal"; $anyChanged=$true; $paramsChanged=$true
                }
            }

            if ($anyChanged) { Set-Content -Path $configIni -Value $cfgLines }
        }
    }

    # 6) Atualiza atalhos (.lnk)
    if ($chkAtalhos.Checked) {
        $pathAt = Join-Path $caminhoNovo 'atalhos'
        if (Test-Path $pathAt) {
            Get-ChildItem -Path $pathAt -Recurse -Filter '*.lnk' -File | ForEach-Object {
                try {
                    $shell = New-Object -ComObject WScript.Shell
                    $s     = $shell.CreateShortcut($_.FullName)
                    $mudou = $false
                    foreach ($campo in 'TargetPath','WorkingDirectory') {
                        $orig = $s.$campo
                        if ($orig -and $orig -like 'C:\*') {
                            $parts = $orig -split '\\',4
                            if ($parts.Count -ge 3) {
                                $parts[1] = $nomeSubpasta
                                $novo     = $parts -join '\'
                                if ($s.$campo -ne $novo) {
                                    $s.$campo = $novo; $mudou=$true; $paramsChanged=$true
                                }
                            }
                        }
                    }
                    if ($mudou) { $s.Save(); $anyChanged=$true }
                } catch {}
            }
        }
    }

    # Mensagens finais
    if ($ocxRegistered) {
        [System.Windows.Forms.MessageBox]::Show(
          "OCX registradas com sucesso.",
          "Sucesso",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    if ($fontsInstalled) {
        [System.Windows.Forms.MessageBox]::Show(
          "Fontes instaladas com sucesso!",
          "Fontes",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    if ($paramsChanged) {
        [System.Windows.Forms.MessageBox]::Show(
          "Parâmetros atualizados com sucesso!",
          "Configuração",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

[System.Windows.Forms.Application]::Run($form)