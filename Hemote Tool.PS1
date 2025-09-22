try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
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

# Selecione a pasta
$lblSacs = New-Object System.Windows.Forms.Label
$lblSacs.Text     = "Selecione a Pasta:"
$lblSacs.AutoSize = $true
$lblSacs.Location = New-Object System.Drawing.Point(10,20)
$tabPage1.Controls.Add($lblSacs)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Location      = New-Object System.Drawing.Point(10,45)
$comboBox.Size          = New-Object System.Drawing.Size(300,20)
$tabPage1.Controls.Add($comboBox)

# Atualiza a lista de pastas SACS toda vez que o usuário abre a combo
$comboBox.Add_DropDown({
    $comboBox.Items.Clear()

    # Inclui a pasta SACS se existir
    if (Test-Path "C:\SACS") {
        $comboBox.Items.Add("SACS")
    }

    # Adiciona as demais pastas que seguem o padrão sacs_*
    $pastas = Get-ChildItem "C:\" -Directory | Where-Object { $_.Name -like "sacs_*" }
    foreach ($pasta in $pastas) {
        $comboBox.Items.Add($pasta.Name)
    }

    # Seleciona automaticamente a primeira pasta, se houver
    if ($comboBox.Items.Count -gt 0) {
        $comboBox.SelectedIndex = 0
    }
})


# Checkbox "Criar cópia"
$chkCriarCopia = New-Object System.Windows.Forms.CheckBox
$chkCriarCopia.Location = New-Object System.Drawing.Point(320,45)
$chkCriarCopia.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkCriarCopia)

# Label "Nova Instância:"
$lblNovaInstancia = New-Object System.Windows.Forms.Label
$lblNovaInstancia.Text     = "Nova Instância:"
$lblNovaInstancia.AutoSize = $true
$lblNovaInstancia.Location = New-Object System.Drawing.Point(345,47)
$tabPage1.Controls.Add($lblNovaInstancia)

$txtNovaInstancia = New-Object System.Windows.Forms.TextBox
$txtNovaInstancia.Location  = New-Object System.Drawing.Point(450,45)
$txtNovaInstancia.Size      = New-Object System.Drawing.Size(100,20)
$txtNovaInstancia.MaxLength = 30
$txtNovaInstancia.Enabled   = $false
$tabPage1.Controls.Add($txtNovaInstancia)

$chkCriarCopia.Add_CheckedChanged({
    $txtNovaInstancia.Enabled = $chkCriarCopia.Checked
    if (-not $chkCriarCopia.Checked) { $txtNovaInstancia.Clear() }
})

# Atalhos
$chkAtalhos = New-Object System.Windows.Forms.CheckBox
$chkAtalhos.Location = New-Object System.Drawing.Point(10,80)
$chkAtalhos.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkAtalhos)

$lblAtalhos = New-Object System.Windows.Forms.Label
$lblAtalhos.Text     = "Atalhos"
$lblAtalhos.AutoSize = $true
$lblAtalhos.Location = New-Object System.Drawing.Point(30,82)
$tabPage1.Controls.Add($lblAtalhos)

# Bootstrap
$chkBootstrap = New-Object System.Windows.Forms.CheckBox
$chkBootstrap.Location = New-Object System.Drawing.Point(110,80)
$chkBootstrap.Size     = New-Object System.Drawing.Size(20,20)
$tabPage1.Controls.Add($chkBootstrap)

$lblBootstrap = New-Object System.Windows.Forms.Label
$lblBootstrap.Text     = "BootStrap"
$lblBootstrap.AutoSize = $true
$lblBootstrap.Location = New-Object System.Drawing.Point(130,82)
$tabPage1.Controls.Add($lblBootstrap)

# LOC_ARM
$chkLocArm = New-Object System.Windows.Forms.CheckBox
$chkLocArm.Location = New-Object System.Drawing.Point(230,80)
$chkLocArm.Size     = New-Object System.Drawing.Size(20,20) 
$chkLocArm.Add_CheckedChanged({
    $txtLocArm.Enabled = $chkLocArm.Checked
    if (-not $chkLocArm.Checked) { $txtLocArm.Clear() }
})
$tabPage1.Controls.Add($chkLocArm)

$lblLocArm = New-Object System.Windows.Forms.Label
$lblLocArm.Text     = "LOC_ARM"
$lblLocArm.AutoSize = $true
$lblLocArm.Location = New-Object System.Drawing.Point(250,82)
$tabPage1.Controls.Add($lblLocArm)

$txtLocArm = New-Object System.Windows.Forms.TextBox
$txtLocArm.Location  = New-Object System.Drawing.Point(350,80)
$txtLocArm.Size      = New-Object System.Drawing.Size(100,20)
$txtLocArm.MaxLength = 20
$txtLocArm.Enabled   = $false
$tabPage1.Controls.Add($txtLocArm)

# COD_HEMO
$chkCodHemo = New-Object System.Windows.Forms.CheckBox
$chkCodHemo.Location = New-Object System.Drawing.Point(10,120)
$chkCodHemo.Size     = New-Object System.Drawing.Size(20,20)
$chkCodHemo.Add_CheckedChanged({
    $txtCodHemo.Enabled = $chkCodHemo.Checked
    if (-not $chkCodHemo.Checked) { $txtCodHemo.Clear() }
})
$tabPage1.Controls.Add($chkCodHemo)

$lblCodHemo = New-Object System.Windows.Forms.Label
$lblCodHemo.Text     = "COD_HEMO"
$lblCodHemo.AutoSize = $true
$lblCodHemo.Location = New-Object System.Drawing.Point(30,122)
$tabPage1.Controls.Add($lblCodHemo)

$txtCodHemo = New-Object System.Windows.Forms.TextBox
$txtCodHemo.Location  = New-Object System.Drawing.Point(110,120)
$txtCodHemo.Size      = New-Object System.Drawing.Size(100,20)
$txtCodHemo.MaxLength = 15
$txtCodHemo.Enabled   = $false
$tabPage1.Controls.Add($txtCodHemo)

# INSTITUIÇÃO
$chkInstituicao = New-Object System.Windows.Forms.CheckBox
$chkInstituicao.Location = New-Object System.Drawing.Point(230,120)
$chkInstituicao.Size     = New-Object System.Drawing.Size(20,20)
$chkInstituicao.Add_CheckedChanged({
    $txtInstituicao.Enabled = $chkInstituicao.Checked
    if (-not $chkInstituicao.Checked) { $txtInstituicao.Clear() }
})
$tabPage1.Controls.Add($chkInstituicao)

$lblInstituicao = New-Object System.Windows.Forms.Label
$lblInstituicao.Text     = "INSTITUIÇÃO"
$lblInstituicao.AutoSize = $true
$lblInstituicao.Location = New-Object System.Drawing.Point(250,122)
$tabPage1.Controls.Add($lblInstituicao)

$txtInstituicao = New-Object System.Windows.Forms.TextBox
$txtInstituicao.Location  = New-Object System.Drawing.Point(350,120)
$txtInstituicao.Size      = New-Object System.Drawing.Size(100,20)
$txtInstituicao.MaxLength = 10
$txtInstituicao.Enabled   = $false
$tabPage1.Controls.Add($txtInstituicao)

# Botão Confirmar
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text     = "Confirmar"
$updateButton.Size     = New-Object System.Drawing.Size(100,30)
$updateButton.Location = New-Object System.Drawing.Point(464,275)
$tabPage1.Controls.Add($updateButton)

$updateButton.Add_Click({
    $anyChanged    = $false
    $paramsChanged = $false

    $usarCopia = $chkCriarCopia.Checked -and $txtNovaInstancia.Text.Trim()
    $nomeSubpasta = if ($usarCopia) {
        "sacs_" + $txtNovaInstancia.Text.Trim().ToLower()
    } elseif ($comboBox.SelectedIndex -ge 0) {
        $comboBox.SelectedItem.ToString().Trim().ToLower()
    } else { "" }

    $caminhoNovo = if ($nomeSubpasta) { Join-Path 'C:\' $nomeSubpasta } else { $null }

    if ($usarCopia) {
        $origem      = $comboBox.SelectedItem.ToString().Trim()
        $caminhoOrig = Join-Path 'C:\' $origem
        if (-not (Test-Path $caminhoOrig)) {
            [System.Windows.Forms.MessageBox]::Show("A pasta de origem não existe.","Erro",
              [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if (Test-Path $caminhoNovo) {
            [System.Windows.Forms.MessageBox]::Show("A nova pasta já existe.","Erro",
              [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        try {
            Copy-Item -Path $caminhoOrig -Destination $caminhoNovo -Recurse -Force
            [System.Windows.Forms.MessageBox]::Show("Cópia realizada com sucesso!",
              "Sucesso",[System.Windows.Forms.MessageBoxButtons]::OK,
              [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao copiar a pasta.","Erro",
              [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    $needsFolder = $chkAtalhos.Checked -or $chkBootstrap.Checked -or $chkCodHemo.Checked -or $chkInstituicao.Checked `
                -or $chkPrinterGraphics.Checked -or $chkPrinterEtiqueta.Checked -or $chkDPI2.Checked -or $chkTipoImp.Checked `
                -or $chkLocArm.Checked

    if ($needsFolder -and -not (Test-Path $caminhoNovo)) {
        [System.Windows.Forms.MessageBox]::Show("A pasta selecionada não existe.","Erro",
          [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # COD_HEMO
    if ($chkCodHemo.Checked -and $txtCodHemo.Text.Trim()) {
        $iniDA = Join-Path $caminhoNovo '_data_access.ini'
        if (Test-Path $iniDA) {
            $lines = Get-Content $iniDA
            for ($i=0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match '^N =\[_cod_hem\]=') {
                    $lines[$i] = "N =[_cod_hem]= $($txtCodHemo.Text.Trim())"
                    $anyChanged = $true; $paramsChanged = $true
                }
            }
            Set-Content -Path $iniDA -Value $lines
        }
    }

# BootStrap
if ($chkBootstrap.Checked -and $caminhoNovo) {
    $bootstrapPath = Join-Path $caminhoNovo 'bootstrap'
    $iniWU = Join-Path $bootstrapPath 'WebUpdate.ini'

    if (Test-Path $iniWU) {
        $lines = Get-Content $iniWU
        $changed = $false

        for ($i = 0; $i -lt $lines.Count; $i++) {
            # Atualiza PathLocalRoot (mais flexível)
            $matchLocal = [regex]::Match($lines[$i], '^PathLocalRoot=C:\\[^\\]+\\update')
            if ($matchLocal.Success) {
                $lines[$i] = "PathLocalRoot=C:\$nomeSubpasta\update"
                $changed = $true
            }

            # Atualiza FinalApp preservando o caminho do atalho
            $matchFinal = [regex]::Match($lines[$i], '^FinalApp="C:\\[^\\]+\\atalhos\\(.+)"')
            if ($matchFinal.Success) {
                $atalhoRelativo = $matchFinal.Groups[1].Value
                $lines[$i] = "FinalApp=`"C:\$nomeSubpasta\atalhos\$atalhoRelativo`""
                $changed = $true
            }
        }

        if ($changed) {
            Set-Content -Path $iniWU -Value $lines
            $anyChanged = $true
            $paramsChanged = $true
        }
    }
}
    # Configuração geral
    $configIni = Join-Path $caminhoNovo 'configuracao.ini'
    if (Test-Path $configIni) {
        $cfgLines = Get-Content $configIni

        if ($chkPrinterGraphics.Checked -and $comboPrinterGraphics.SelectedIndex -ge 0) {
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

        if ($chkPrinterEtiqueta.Checked -and $comboPrinterEtiqueta.SelectedIndex -ge 0) {
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

        if ($chkLocArm.Checked -and $txtLocArm.Text.Trim()) {
            $locVal = $txtLocArm.Text.Trim(); $found = $false
            for ($i=0; $i -lt $cfgLines.Count; $i++) {
                if ($cfgLines[$i] -match '^\[LOC_ARM\]\s*=') {
                    $cfgLines[$i] = "[LOC_ARM]         = $locVal"; $anyChanged=$true; $paramsChanged=$true; $found = $true
                }
            }
            if (-not $found) {
                $cfgLines += "[LOC_ARM]         = $locVal"; $anyChanged=$true; $paramsChanged=$true
            }
        }

        if ($anyChanged) { Set-Content -Path $configIni -Value $cfgLines }
    }

    # Atualiza atalhos (.lnk)
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

    # Mensagem final
    if ($paramsChanged) {
        [System.Windows.Forms.MessageBox]::Show(
            "Parâmetros atualizados com sucesso!",
            "Configuração",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

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

# Impressora Etiqueta
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
$comboPrinterEtiqueta.Size          = New-Object System.Drawing.Size(185,20)
$comboPrinterEtiqueta.Enabled       = $false
$tabPage1.Controls.Add($comboPrinterEtiqueta)

# Tipo Impressora
$chkTipoImp = New-Object System.Windows.Forms.CheckBox
$chkTipoImp.Location = New-Object System.Drawing.Point(345,160)
$chkTipoImp.Size     = New-Object System.Drawing.Size(20,20)
$chkTipoImp.Add_CheckedChanged({
    $comboTipoImp.Enabled = $chkTipoImp.Checked
    if (-not $chkTipoImp.Checked) { $comboTipoImp.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkTipoImp)

$lblTipoImp = New-Object System.Windows.Forms.Label
$lblTipoImp.Text     = "Tipo Impressora"
$lblTipoImp.AutoSize = $true
$lblTipoImp.Location = New-Object System.Drawing.Point(365,162)
$tabPage1.Controls.Add($lblTipoImp)

$comboTipoImp = New-Object System.Windows.Forms.ComboBox
$comboTipoImp.DropDownStyle = 'DropDownList'
$comboTipoImp.Items.AddRange(@("1 - ALEGRO","2 - ZEBRA"))
$comboTipoImp.Location      = New-Object System.Drawing.Point(460,160)
$comboTipoImp.Size          = New-Object System.Drawing.Size(100,20)
$comboTipoImp.Enabled       = $false
$tabPage1.Controls.Add($comboTipoImp)

# Impressora Gráfica
$chkPrinterGraphics = New-Object System.Windows.Forms.CheckBox
$chkPrinterGraphics.Location = New-Object System.Drawing.Point(10,200)
$chkPrinterGraphics.Size     = New-Object System.Drawing.Size(20,20)
$chkPrinterGraphics.Add_CheckedChanged({
    $comboPrinterGraphics.Enabled = $chkPrinterGraphics.Checked
    if (-not $chkPrinterGraphics.Checked) { $comboPrinterGraphics.SelectedIndex = -1 }
})
$tabPage1.Controls.Add($chkPrinterGraphics)

$lblPrinterGraphics = New-Object System.Windows.Forms.Label
$lblPrinterGraphics.Text     = "Impressora Gráfica"
$lblPrinterGraphics.AutoSize = $true
$lblPrinterGraphics.Location = New-Object System.Drawing.Point(30,202)
$tabPage1.Controls.Add($lblPrinterGraphics)

$comboPrinterGraphics = New-Object System.Windows.Forms.ComboBox
$comboPrinterGraphics.DropDownStyle = 'DropDownList'
$comboPrinterGraphics.Location      = New-Object System.Drawing.Point(150,200)
$comboPrinterGraphics.Size          = New-Object System.Drawing.Size(185,20)
$comboPrinterGraphics.Enabled       = $false
$tabPage1.Controls.Add($comboPrinterGraphics)




# DPI
$chkDPI2 = New-Object System.Windows.Forms.CheckBox
$chkDPI2.Location = New-Object System.Drawing.Point(345,200)
$chkDPI2.Size     = New-Object System.Drawing.Size(20,20)
$chkDPI2.Add_CheckedChanged({
    $txtDPI2.Enabled = $chkDPI2.Checked
    if (-not $chkDPI2.Checked) { $txtDPI2.Clear() }
})
$tabPage1.Controls.Add($chkDPI2)

$lblDPI2 = New-Object System.Windows.Forms.Label
$lblDPI2.Text     = "DPI"
$lblDPI2.AutoSize = $true
$lblDPI2.Location = New-Object System.Drawing.Point(365,202)
$tabPage1.Controls.Add($lblDPI2)

$txtDPI2 = New-Object System.Windows.Forms.TextBox
$txtDPI2.Location  = New-Object System.Drawing.Point(390,200)
$txtDPI2.Size      = New-Object System.Drawing.Size(30,20)
$txtDPI2.MaxLength = 3
$txtDPI2.Enabled   = $false
$tabPage1.Controls.Add($txtDPI2)

# Aba Sobre – descrição da ferramenta
$txtSobre = New-Object System.Windows.Forms.TextBox
$txtSobre.Multiline   = $true
$txtSobre.ReadOnly    = $true
$txtSobre.ScrollBars  = 'Vertical'
$txtSobre.Dock        = [System.Windows.Forms.DockStyle]::Fill
$txtSobre.Text = @"
Bem-vindo ao Hemote Tool

Desenvolvido para acelerar tarefas de rotina, reduzir erros, retrabalhos, otimizar o tempo de resposta e elevar a qualidade das entregas. O Hemote Tool automatiza:

 • Atualização de atalhos
 • Ajustes de parâmetros (COD_HEMO, Instituição, DPI, Tipo de Impressora, LOC_ARM)
 • Configuração de impressoras (etiqueta e gráfica)

Versão: 2.0
Autor: Felipe Almeida
"@
$tabPageSobre.Controls.Clear()
$tabPageSobre.Controls.Add($txtSobre)
$toolTip = New-Object System.Windows.Forms.ToolTip

$toolTip.SetToolTip($lblSacs, "Selecione a pasta base do sistema SACS.")
$toolTip.SetToolTip($lblAtalhos, "Atualiza os atalhos (.lnk) para apontarem para a nova pasta.")
$toolTip.SetToolTip($lblBootstrap, "Atualiza o caminho de bootstrap para refletir a nova instância.")
$toolTip.SetToolTip($lblLocArm, "Define o valor da variável LOC_ARM no arquivo de configuração.")
$toolTip.SetToolTip($lblCodHemo, "Define o código do hemocentro no arquivo _data_access.ini.")
$toolTip.SetToolTip($lblInstituicao, "Define o nome da instituição no arquivo de configuração.")
$toolTip.SetToolTip($lblPrinterEtiqueta, "Seleciona a impressora usada para imprimir etiquetas.")
$toolTip.SetToolTip($lblPrinterGraphics, "Seleciona a impressora usada para imprimir fichas gráficas.")
$toolTip.SetToolTip($lblTipoImp, "Define o tipo de impressora: 1 para ALEGRO, 2 para ZEBRA.")
$toolTip.SetToolTip($lblDPI2, "Define a resolução de impressão (DPI) para etiquetas.")
$toolTip.SetToolTip($chkCriarCopia, "Ative para criar uma nova instância com o nome digitado ao lado.")
$toolTip.SetToolTip($lblNovaInstancia, "Exemplo: EXPEDICAO, a pasta criada será 'sacs_expedicao'")

# Executa o formulário
[System.Windows.Forms.Application]::Run($form)