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

# Cria o TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'

# Cria as abas
$tabPage1         = New-Object System.Windows.Forms.TabPage("Atualizador")
$tabPageBootstrap = New-Object System.Windows.Forms.TabPage("Bootstrap")
$tabPageCopiaExe  = New-Object System.Windows.Forms.TabPage("Copiar exe")
$tabPage2         = New-Object System.Windows.Forms.TabPage("Sobre")

# Adiciona as abas na ordem desejada
$tabControl.TabPages.Add($tabPage1)         # Aba 0: Atualizador
$tabControl.TabPages.Add($tabPageBootstrap) # Aba 1: Bootstrap
$tabControl.TabPages.Add($tabPageCopiaExe)  # Aba 2: Copiar exe
$tabControl.TabPages.Add($tabPage2)         # Aba 3: Sobre

# Anexa o TabControl ao formulário
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

# Checkbox para ativar/desativar a caixa de texto (sem texto)
$chkCodHemo = New-Object System.Windows.Forms.CheckBox
$chkCodHemo.Location = New-Object System.Drawing.Point(10,80)
$chkCodHemo.Size     = New-Object System.Drawing.Size(20,20)
$chkCodHemo.Add_CheckedChanged({
    if ($chkCodHemo.Checked) {
        $txtCodHemo.Enabled = $true
    } else {
        $txtCodHemo.Enabled = $false
        $txtCodHemo.Text    = ""
    }
})
$tabPage1.Controls.Add($chkCodHemo)

# Label COD_HEMO
$lblCodHemo = New-Object System.Windows.Forms.Label
$lblCodHemo.Text     = "COD_HEMO"
$lblCodHemo.Location = New-Object System.Drawing.Point(30,82)
$lblCodHemo.Size     = New-Object System.Drawing.Size(80,20)
$tabPage1.Controls.Add($lblCodHemo)

# Caixa de texto COD_HEMO (desativada por padrão)
$txtCodHemo = New-Object System.Windows.Forms.TextBox
$txtCodHemo.Location = New-Object System.Drawing.Point(110,80)
$txtCodHemo.Size     = New-Object System.Drawing.Size(100,20)
$txtCodHemo.MaxLength = 15
$txtCodHemo.Enabled   = $false
$tabPage1.Controls.Add($txtCodHemo)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline   = $true
$statusBox.ScrollBars  = 'Vertical'
$statusBox.ReadOnly    = $true
$statusBox.Size        = New-Object System.Drawing.Size(460,100)
$statusBox.Location    = New-Object System.Drawing.Point(10,150)
$tabPage1.Controls.Add($statusBox)

# Evento do botão Atualizar
$updateButton.Add_Click({
    $nomeSubpasta    = $textBox.Text.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($nomeSubpasta)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Informe o caminho a ser atualizado.","Erro",
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

    # Exibe resultado atualização de atalhos
    if ($alterados -gt 0) {
        $statusBox.Text = "✅ Total de atalhos atualizados: $alterados"
    } else {
        $statusBox.Text = "ℹ️ Nenhum atalho precisava ser atualizado."
    }

    # Ajuste do WebUpdate.ini se houver bootstrap
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

    # Ajusta parâmetro COD_HEMO em _data_access.ini
    if ($chkCodHemo.Checked) {
        $arquivoDataAccess = Join-Path $caminhoNovo "_data_access.ini"
        if (Test-Path $arquivoDataAccess) {
            $linhasData = Get-Content $arquivoDataAccess
            for ($j = 0; $j -lt $linhasData.Count; $j++) {
                if ($linhasData[$j] -match '^N =\[_cod_hem\]=') {
                    $linhasData[$j] = "N =[_cod_hem]= $($txtCodHemo.Text)"
                    break
                }
            }
            Set-Content -Path $arquivoDataAccess -Value $linhasData
            [System.Windows.Forms.MessageBox]::Show(
                "_data_access alterado para $($txtCodHemo.Text)",
                "Sucesso",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Arquivo '_data_access.ini' não encontrado em '$caminhoNovo'.",
                "Erro",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
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