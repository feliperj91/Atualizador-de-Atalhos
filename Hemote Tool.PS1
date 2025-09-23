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
$form.ClientSize      = New-Object System.Drawing.Size(580,430)
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

### GROUPBOX 1 – DEFINIÇÃO DE PASTAS
$groupPastas = New-Object System.Windows.Forms.GroupBox
$groupPastas.Text = "Definição de Pastas"
$groupPastas.Location = New-Object System.Drawing.Point(10, 10)
$groupPastas.Size = New-Object System.Drawing.Size(560, 80)
$tabPage1.Controls.Add($groupPastas)

$lblSacs = New-Object System.Windows.Forms.Label
$lblSacs.Text = "Selecione a pasta sacs a ser configurada:"
$lblSacs.AutoSize = $true
$lblSacs.Location = New-Object System.Drawing.Point(10,20)
$lblSacs.Parent = $groupPastas

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.DropDownStyle = 'DropDownList'
$comboBox.Location = New-Object System.Drawing.Point(10,45)
$comboBox.Size = New-Object System.Drawing.Size(300,20)
$comboBox.Parent = $groupPastas

$comboBox.Add_DropDown({
    $comboBox.Items.Clear()
    if (Test-Path "C:\SACS") {
        $comboBox.Items.Add("SACS")
    }
    $pastas = Get-ChildItem "C:\" -Directory | Where-Object { $_.Name -like "sacs_*" }
    foreach ($pasta in $pastas) {
        $comboBox.Items.Add($pasta.Name)
    }
    if ($comboBox.Items.Count -gt 0) {
        $comboBox.SelectedIndex = 0
    }
})

$chkCriarCopia = New-Object System.Windows.Forms.CheckBox
$chkCriarCopia.Location = New-Object System.Drawing.Point(320,45)
$chkCriarCopia.Size = New-Object System.Drawing.Size(20,20)
$chkCriarCopia.Parent = $groupPastas

$lblNovaInstancia = New-Object System.Windows.Forms.Label
$lblNovaInstancia.Text = "Nova pasta SACS:"
$lblNovaInstancia.AutoSize = $true
$lblNovaInstancia.Location = New-Object System.Drawing.Point(345,47)
$lblNovaInstancia.Parent = $groupPastas

$txtNovaInstancia = New-Object System.Windows.Forms.TextBox
$txtNovaInstancia.Location = New-Object System.Drawing.Point(450,45)
$txtNovaInstancia.Size = New-Object System.Drawing.Size(100,20)
$txtNovaInstancia.MaxLength = 30
$txtNovaInstancia.Enabled = $false
$txtNovaInstancia.Parent = $groupPastas

$chkCriarCopia.Add_CheckedChanged({
    $txtNovaInstancia.Enabled = $chkCriarCopia.Checked
    if (-not $chkCriarCopia.Checked) { $txtNovaInstancia.Clear() }
})

### GROUPBOX 2 – PARÂMETROS GERAIS
$groupParametros = New-Object System.Windows.Forms.GroupBox
$groupParametros.Text = "Parâmetros Gerais"
$groupParametros.Location = New-Object System.Drawing.Point(10, 100)
$groupParametros.Size = New-Object System.Drawing.Size(560, 120)
$tabPage1.Controls.Add($groupParametros)

$chkAtalhos = New-Object System.Windows.Forms.CheckBox
$chkAtalhos.Location = New-Object System.Drawing.Point(10,20)
$chkAtalhos.Size = New-Object System.Drawing.Size(20,20)
$chkAtalhos.Parent = $groupParametros

$lblAtalhos = New-Object System.Windows.Forms.Label
$lblAtalhos.Text = "Atalhos"
$lblAtalhos.AutoSize = $true
$lblAtalhos.Location = New-Object System.Drawing.Point(30,22)
$lblAtalhos.Parent = $groupParametros

$chkBootstrap = New-Object System.Windows.Forms.CheckBox
$chkBootstrap.Location = New-Object System.Drawing.Point(110,20)
$chkBootstrap.Size = New-Object System.Drawing.Size(20,20)
$chkBootstrap.Parent = $groupParametros

$lblBootstrap = New-Object System.Windows.Forms.Label
$lblBootstrap.Text = "BootStrap"
$lblBootstrap.AutoSize = $true
$lblBootstrap.Location = New-Object System.Drawing.Point(130,22)
$lblBootstrap.Parent = $groupParametros

$chkLocArm = New-Object System.Windows.Forms.CheckBox
$chkLocArm.Location = New-Object System.Drawing.Point(230,20)
$chkLocArm.Size = New-Object System.Drawing.Size(20,20)
$chkLocArm.Parent = $groupParametros

$lblLocArm = New-Object System.Windows.Forms.Label
$lblLocArm.Text = "LOC_ARM"
$lblLocArm.AutoSize = $true
$lblLocArm.Location = New-Object System.Drawing.Point(250,22)
$lblLocArm.Parent = $groupParametros

$txtLocArm = New-Object System.Windows.Forms.TextBox
$txtLocArm.Location = New-Object System.Drawing.Point(350,20)
$txtLocArm.Size = New-Object System.Drawing.Size(100,20)
$txtLocArm.MaxLength = 20
$txtLocArm.Enabled = $false
$txtLocArm.Parent = $groupParametros

$chkLocArm.Add_CheckedChanged({
    $txtLocArm.Enabled = $chkLocArm.Checked
    if (-not $chkLocArm.Checked) { $txtLocArm.Clear() }
})

$chkCodHemo = New-Object System.Windows.Forms.CheckBox
$chkCodHemo.Location = New-Object System.Drawing.Point(10,60)
$chkCodHemo.Size = New-Object System.Drawing.Size(20,20)
$chkCodHemo.Parent = $groupParametros
$chkCodHemo.Add_CheckedChanged({
    $txtCodHemo.Enabled = $chkCodHemo.Checked
    if (-not $chkCodHemo.Checked) { $txtCodHemo.Clear() }
})

$lblCodHemo = New-Object System.Windows.Forms.Label
$lblCodHemo.Text = "COD_HEMO"
$lblCodHemo.AutoSize = $true
$lblCodHemo.Location = New-Object System.Drawing.Point(30,62)
$lblCodHemo.Parent = $groupParametros

$txtCodHemo = New-Object System.Windows.Forms.TextBox
$txtCodHemo.Location = New-Object System.Drawing.Point(110,60)
$txtCodHemo.Size = New-Object System.Drawing.Size(100,20)
$txtCodHemo.MaxLength = 15
$txtCodHemo.Enabled = $false
$txtCodHemo.Parent = $groupParametros

$chkInstituicao = New-Object System.Windows.Forms.CheckBox
$chkInstituicao.Location = New-Object System.Drawing.Point(230,60)
$chkInstituicao.Size = New-Object System.Drawing.Size(20,20)
$chkInstituicao.Parent = $groupParametros
$chkInstituicao.Add_CheckedChanged({
    $txtInstituicao.Enabled = $chkInstituicao.Checked
    if (-not $chkInstituicao.Checked) { $txtInstituicao.Clear() }
})

$lblInstituicao = New-Object System.Windows.Forms.Label
$lblInstituicao.Text = "INSTITUIÇÃO"
$lblInstituicao.AutoSize = $true
$lblInstituicao.Location = New-Object System.Drawing.Point(250,62)
$lblInstituicao.Parent = $groupParametros

$txtInstituicao = New-Object System.Windows.Forms.TextBox
$txtInstituicao.Location = New-Object System.Drawing.Point(350,60)
$txtInstituicao.Size = New-Object System.Drawing.Size(100,20)
$txtInstituicao.MaxLength = 10
$txtInstituicao.Enabled = $false
$txtInstituicao.Parent = $groupParametros

### GROUPBOX 3 – IMPRESSORAS
$groupImpressoras = New-Object System.Windows.Forms.GroupBox
$groupImpressoras.Text = "Impressoras"
$groupImpressoras.Location = New-Object System.Drawing.Point(10, 230)
$groupImpressoras.Size = New-Object System.Drawing.Size(560, 100)
$tabPage1.Controls.Add($groupImpressoras)

$chkPrinterEtiqueta = New-Object System.Windows.Forms.CheckBox
$chkPrinterEtiqueta.Location = New-Object System.Drawing.Point(10,20)
$chkPrinterEtiqueta.Size = New-Object System.Drawing.Size(20,20)
$chkPrinterEtiqueta.Parent = $groupImpressoras
$chkPrinterEtiqueta.Add_CheckedChanged({
    $comboPrinterEtiqueta.Enabled = $chkPrinterEtiqueta.Checked
    if (-not $chkPrinterEtiqueta.Checked) { $comboPrinterEtiqueta.SelectedIndex = -1 }
})

$lblPrinterEtiqueta = New-Object System.Windows.Forms.Label
$lblPrinterEtiqueta.Text = "Impressora Etiqueta"
$lblPrinterEtiqueta.AutoSize = $true
$lblPrinterEtiqueta.Location = New-Object System.Drawing.Point(30,22)
$lblPrinterEtiqueta.Parent = $groupImpressoras

$comboPrinterEtiqueta = New-Object System.Windows.Forms.ComboBox
$comboPrinterEtiqueta.DropDownStyle = 'DropDownList'
$comboPrinterEtiqueta.Location = New-Object System.Drawing.Point(150,20)
$comboPrinterEtiqueta.Size = New-Object System.Drawing.Size(185,20)
$comboPrinterEtiqueta.Enabled = $false
$comboPrinterEtiqueta.Parent = $groupImpressoras

$chkTipoImp = New-Object System.Windows.Forms.CheckBox
$chkTipoImp.Location = New-Object System.Drawing.Point(345,20)
$chkTipoImp.Size = New-Object System.Drawing.Size(20,20)
$chkTipoImp.Parent = $groupImpressoras
$chkTipoImp.Add_CheckedChanged({
    $comboTipoImp.Enabled = $chkTipoImp.Checked
    if (-not $chkTipoImp.Checked) { $comboTipoImp.SelectedIndex = -1 }
})

$lblTipoImp = New-Object System.Windows.Forms.Label
$lblTipoImp.Text = "Tipo Impressora"
$lblTipoImp.AutoSize = $true
$lblTipoImp.Location = New-Object System.Drawing.Point(365,22)
$lblTipoImp.Parent = $groupImpressoras

$comboTipoImp = New-Object System.Windows.Forms.ComboBox
$comboTipoImp.DropDownStyle = 'DropDownList'
$comboTipoImp.Items.AddRange(@("1 - ALEGRO","2 - ZEBRA"))
$comboTipoImp.Location = New-Object System.Drawing.Point(460,20)
$comboTipoImp.Size = New-Object System.Drawing.Size(90,20)
$comboTipoImp.Enabled = $false
$comboTipoImp.Parent = $groupImpressoras

# Impressora Gráfica
$chkPrinterGraphics = New-Object System.Windows.Forms.CheckBox
$chkPrinterGraphics.Location = New-Object System.Drawing.Point(10,60)
$chkPrinterGraphics.Size = New-Object System.Drawing.Size(20,20)
$chkPrinterGraphics.Parent = $groupImpressoras
$chkPrinterGraphics.Add_CheckedChanged({
    $comboPrinterGraphics.Enabled = $chkPrinterGraphics.Checked
    if (-not $chkPrinterGraphics.Checked) { $comboPrinterGraphics.SelectedIndex = -1 }
})

$lblPrinterGraphics = New-Object System.Windows.Forms.Label
$lblPrinterGraphics.Text = "Impressora Gráfica"
$lblPrinterGraphics.AutoSize = $true
$lblPrinterGraphics.Location = New-Object System.Drawing.Point(30,62)
$lblPrinterGraphics.Parent = $groupImpressoras

$comboPrinterGraphics = New-Object System.Windows.Forms.ComboBox
$comboPrinterGraphics.DropDownStyle = 'DropDownList'
$comboPrinterGraphics.Location = New-Object System.Drawing.Point(150,60)
$comboPrinterGraphics.Size = New-Object System.Drawing.Size(185,20)
$comboPrinterGraphics.Enabled = $false
$comboPrinterGraphics.Parent = $groupImpressoras

# DPI
$chkDPI2 = New-Object System.Windows.Forms.CheckBox
$chkDPI2.Location = New-Object System.Drawing.Point(345,60)
$chkDPI2.Size = New-Object System.Drawing.Size(20,20)
$chkDPI2.Parent = $groupImpressoras
$chkDPI2.Add_CheckedChanged({
    $txtDPI2.Enabled = $chkDPI2.Checked
    if (-not $chkDPI2.Checked) { $txtDPI2.Clear() }
})

$lblDPI2 = New-Object System.Windows.Forms.Label
$lblDPI2.Text = "DPI"
$lblDPI2.AutoSize = $true
$lblDPI2.Location = New-Object System.Drawing.Point(365,62)
$lblDPI2.Parent = $groupImpressoras

$txtDPI2 = New-Object System.Windows.Forms.TextBox
$txtDPI2.Location = New-Object System.Drawing.Point(460,60)
$txtDPI2.Size = New-Object System.Drawing.Size(50,20)
$txtDPI2.MaxLength = 3
$txtDPI2.Enabled = $false
$txtDPI2.Parent = $groupImpressoras

# Botão Confirmar
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text     = "Confirmar"
$updateButton.Size     = New-Object System.Drawing.Size(100,30)
$updateButton.Location = New-Object System.Drawing.Point(464,340)
$tabPage1.Controls.Add($updateButton)

# Início do evento do botão Confirmar
$updateButton.Add_Click({
    $anyChanged    = $false
    $paramsChanged = $false

    if (-not $comboBox.SelectedItem -and -not $chkCriarCopia.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Selecione uma pasta ou crie uma nova instância.","Aviso",
          [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $usarCopia = $chkCriarCopia.Checked -and $txtNovaInstancia.Text.Trim()
    $nomeSubpasta = if ($usarCopia) {
        "sacs_" + $txtNovaInstancia.Text.Trim().ToLower()
    } elseif ($comboBox.SelectedItem) {
        $comboBox.SelectedItem.ToString().Trim().ToLower()
    } else { "" }

    if (-not $nomeSubpasta) {
        [System.Windows.Forms.MessageBox]::Show("Não foi possível determinar a pasta de destino.","Erro",
          [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $caminhoNovo = Join-Path 'C:\' $nomeSubpasta

    if ($usarCopia) {
        $origem = $comboBox.SelectedItem.ToString().Trim()
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
                $matchLocal = [regex]::Match($lines[$i], '^PathLocalRoot=C:\\[^\\]+\\update')
                if ($matchLocal.Success) {
                    $lines[$i] = "PathLocalRoot=C:\$nomeSubpasta\update"
                    $changed = $true
                }

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
	
# Evento ao abrir o formulário – carrega pastas e impressoras
$form.Add_Shown({
    $comboBox.Items.Clear()
    if (Test-Path "C:\SACS") {
        $comboBox.Items.Add("SACS")
    }
    Get-ChildItem -Path 'C:\' -Directory -Filter 'sacs_*' |
        Sort-Object Name |
        ForEach-Object { $comboBox.Items.Add($_.Name) }

    if ($comboBox.Items.Contains($defaultSacs)) {
        $comboBox.SelectedItem = $defaultSacs
    } elseif ($comboBox.Items.Count -gt 0) {
        $comboBox.SelectedIndex = 0
    }

    # Carrega impressoras instaladas
    $comboPrinterGraphics.Items.Clear()
    $comboPrinterEtiqueta.Items.Clear()
    foreach ($p in [System.Drawing.Printing.PrinterSettings]::InstalledPrinters) {
        $comboPrinterGraphics.Items.Add($p)
        $comboPrinterEtiqueta.Items.Add($p)
    }
})

# Aba Sobre – descrição da ferramenta
$txtSobre = New-Object System.Windows.Forms.TextBox
$txtSobre.Multiline   = $true
$txtSobre.ReadOnly    = $true
$txtSobre.ScrollBars  = 'Vertical'
$txtSobre.Dock        = [System.Windows.Forms.DockStyle]::Fill
$txtSobre.Text = @"

Hemote Tool — Eficiência e controle na configuração do SACS

O Hemote Tool foi desenvolvido para simplificar e agilizar o processo de configuração de instâncias do sistema SACS. Com uma interface intuitiva e recursos automatizados, a ferramenta ajuda equipes técnicas e operacionais a manterem consistência, reduzirem erros e ganharem tempo em tarefas recorrentes.

Principais recursos:

 • Criação de novas instâncias com estrutura replicada e nome personalizado
 • Atualização automática de atalhos (.lnk) com redirecionamento para a nova instância
 • Ajuste direto de parâmetros como COD_HEMO, INSTITUIÇÃO, LOC_ARM, DPI e tipo de impressora
 • Seleção de impressoras de forma prática, evitando erro de digitação.
 • Correção dos caminhos nos arquivos de configuração (_data_access.ini, WebUpdate.ini, configuracao.ini)
 • Detecção automática de pastas SACS presentes no C:\ e preenchimento dinâmico dos campos.
 • Interface organizada por função, com dicas explicativas em cada campo.



___________________________________________________________



Versão: 2.0  
Autor: Felipe Almeida  
Atualizado em: Setembro de 2025
"@
$tabPageSobre.Controls.Clear()
$tabPageSobre.Controls.Add($txtSobre)

# ToolTips
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