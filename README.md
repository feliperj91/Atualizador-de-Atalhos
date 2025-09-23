# Hemote Tool

**Hemote Tool** é uma aplicação gráfica desenvolvida para facilitar a configuração e manutenção de instâncias do sistema SACS. Com foco em agilidade, padronização e redução de erros operacionais, a ferramenta oferece uma interface intuitiva e recursos automatizados para equipes técnicas e operacionais.

## 🚀 Funcionalidades

- Criação de novas instâncias com estrutura replicada e nome personalizado
- Atualização automática de atalhos (.lnk) com redirecionamento para a nova instância
- Modificação segura dos arquivos:
  - `_data_access.ini`
  - `WebUpdate.ini`
  - `configuracao.ini`
- Ajuste direto de parâmetros como:
  - `COD_HEMO`
  - `INSTITUIÇÃO`
  - `LOC_ARM`
  - `DPI`
  - Tipo de Impressora
- Seleção de impressoras instaladas para etiquetas e fichas gráficas
- Detecção automática de pastas SACS e preenchimento dinâmico dos campos
- Interface organizada por função, com ToolTips explicativos em cada etapa
- Validação de campos e caminhos antes de qualquer operação crítica
- Registro de ações em log para rastreabilidade e suporte técnico

## 🖥️ Requisitos

- Windows 8.1 ou superior
- PowerShell 5.1+
- Permissões administrativas (recomendado para operações de cópia e edição de arquivos)

## 📦 Instalação

1. Clone ou baixe este repositório
2. Execute o script `HemoteTool.ps1` com PowerShell
3. (Opcional) Converta para `.exe` com [PS2EXE](https://github.com/MScholtes/PS2EXE)

## 📄 Licença

Este projeto é de uso interno e não possui licença pública definida. Para fins de distribuição ou adaptação, entre em contato com o autor.

## 👨‍💻 Autor

**Felipe Almeida**  
Atualizado em: Setembro de 2025  
Versão: 2.0
