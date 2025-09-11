Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Cria o formulário
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Atualizador de Atalhos"
$form.ClientSize      = New-Object System.Drawing.Size(500,300)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $true
$form.Topmost         = $true

# Cria o TabControl e abas
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$tabPage1         = New-Object System.Windows.Forms.TabPage("Atualizador")
$tabPageBootstrap = New-Object System.Windows.Forms.TabPage("Bootstrap")
$tabPage2         = New-Object System.Windows.Forms.TabPage("Sobre")
$tabControl.TabPages.Add($tabPage1)
$tabControl.TabPages.Add($tabPageBootstrap)
$tabControl.TabPages.Add($tabPage2)
$form.Controls.Add($tabControl)

# ---------------------- ABA ATUALIZADOR ----------------------
$label = New-Object System.Windows.Forms.Label
$label.Text     = "Informe a pasta SACS (ex: sacs_teste):"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,20)
$tabPage1.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size     = New-Object System.Drawing.Size(350,20)
$textBox.Location = New-Object System.Drawing.Point(10,50)
$tabPage1.Controls.Add($textBox)

$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text     = "Atualizar"
$updateButton.Location = New-Object System.Drawing.Point(370,48)
$tabPage1.Controls.Add($updateButton)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline   = $true
$statusBox.ScrollBars  = 'Vertical'
$statusBox.ReadOnly    = $true
$statusBox.Size        = New-Object System.Drawing.Size(460,80)
$statusBox.Location    = New-Object System.Drawing.Point(10,90)
$tabPage1.Controls.Add($statusBox)

$updateButton.Add_Click({
    $nomeSubpasta = $textBox.Text.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($nomeSubpasta)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Digite o nome da subpasta.","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $caminhoNovo    = "C:\$nomeSubpasta"
    $caminhoAtalhos = Join-Path $caminhoNovo "atalhos"

    if (-not (Test-Path $caminhoAtalhos)) {
        [System.Windows.Forms.MessageBox]::Show(
            "A pasta 'atalhos' não foi encontrada em '$caminhoNovo'.","Erro",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Atualiza atalhos
    $atalhos   = Get-ChildItem -Path $caminhoAtalhos -Recurse -Filter *.lnk -ErrorAction SilentlyContinue |
                 Where-Object { -not $_.PSIsContainer }
    $alterados = 0
    foreach ($atalho in $atalhos) {
        try {
            $shell    = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($atalho.FullName)
            $foiAlterado = $false
            foreach ($campo in @("TargetPath","WorkingDirectory")) {
                $valorOriginal = $shortcut.$campo
                if ($valorOriginal -like "C:\*") {
                    $partes = $valorOriginal -split "\\",4
                    if ($partes.Count -ge 3) {
                        $partes[1] = $nomeSubpasta
                        $novoValor = $partes -join "\"
                        if ($shortcut.$campo -ne $novoValor) {
                            $shortcut.$campo = $novoValor
                            $foiAlterado    = $true
                        }
                    }
                }
            }
            if ($foiAlterado) {
                $shortcut.Save()
                $alterados++
            }
        } catch {}
    }

    # Exibe resultado atualização
    if ($alterados -gt 0) {
        $statusBox.Text = "✅ Total de atalhos atualizados: $alterados"
    } else {
        $statusBox.Text = "ℹ️ Nenhum atalho precisava ser atualizado."
    }

    # Se existir pasta bootstrap e houver alterações, pergunta para ajustar WebUpdate.ini
    $caminhoBootstrap = Join-Path $caminhoNovo "bootstrap"
    if ($alterados -gt 0 -and (Test-Path $caminhoBootstrap)) {
        $resp = [System.Windows.Forms.MessageBox]::Show(
            "A pasta 'bootstrap' foi encontrada em '$caminhoNovo'.`nDeseja ajustar o WebUpdate.ini agora?",
            "Ajuste Bootstrap",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($resp -eq [System.Windows.Forms.DialogResult]::Yes) {
            $arquivoIni = Join-Path $caminhoBootstrap "WebUpdate.ini"
            if (-not (Test-Path $arquivoIni)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Arquivo WebUpdate.ini não encontrado em '$caminhoBootstrap'.",
                    "Erro",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }

            $linhas   = Get-Content $arquivoIni
            $alterIni = 0
            for ($i = 0; $i -lt $linhas.Count; $i++) {
                if ($linhas[$i] -match '^PathLocalRoot=C:\\[^\\]+\\update\\?$') {
                    $linhas[$i] = "PathLocalRoot=C:\$nomeSubpasta\update"
                    $alterIni++
                }
                if ($linhas[$i] -match '^FinalApp="C:\\[^\\]+\\atalhos\\.*"$') {
                    $linhas[$i] = $linhas[$i] -replace '^FinalApp="C:\\[^\\]+\\atalhos\\', "FinalApp=`"C:\$nomeSubpasta\atalhos\"
                    $alterIni++
                }
            }
            if ($alterIni -gt 0) {
                Set-Content -Path $arquivoIni -Value $linhas
                [System.Windows.Forms.MessageBox]::Show(
                    "WebUpdate.ini ajustado: $alterIni alteração(ões).",
                    "Sucesso",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Nenhum parâmetro precisava ser ajustado no WebUpdate.ini.",
                    "Info",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    }
})

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

# ---------------------- ABA SOBRE ----------------------
$aboutLabel = New-Object System.Windows.Forms.Label
$aboutLabel.AutoSize = $false
$aboutLabel.Size     = New-Object System.Drawing.Size(460,180)
$aboutLabel.Location = New-Object System.Drawing.Point(10,10)
$aboutLabel.Text     = @"
Instruções de uso:

- Aba Atualizador: informe o nome da pasta que contém os atalhos do Hemote Plus. O script corrige o caminho de destino e o diretório de trabalho. Em seguida, pergunta se deseja ajustar o WebUpdate.ini na pasta Bootstrap; se confirmado, faz o ajuste automático.

- Aba Bootstrap: permite, de forma independente, ajustar os parâmetros PathLocalRoot e FinalApp no arquivo WebUpdate.ini presente na pasta Bootstrap.

Versão: 1.6
Autor: Felipe Almeida
"@
$tabPage2.Controls.Add($aboutLabel)

# Exibe o formulário
[void]$form.ShowDialog()