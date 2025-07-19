#Requires -Version 5.1

<#
.SYNOPSIS
    Gera assinaturas de email personalizadas para o Outlook baseadas em informacoes do Active Directory.

.DESCRIPTION
    Este script cria assinaturas HTML e texto para o Outlook automaticamente, 
    baseando-se nas informacoes do usuario no Active Directory e em grupos especificos.

.PARAMETER CompanyName
    Nome da empresa (sobrescreve o valor do AD se fornecido)

.PARAMETER Website
    Website da empresa (sobrescreve o valor do AD se fornecido)

.PARAMETER Encoding
    Codificacao para o arquivo de texto da assinatura

.PARAMETER LogLevel
    Nivel de log (Error, Warning, Information, Verbose)

.EXAMPLE
    .\AssinaturaEmail_Final.ps1 -Verbose
    
.EXAMPLE
    .\AssinaturaEmail_Final.ps1 -CompanyName "Trust Group" -Website "https://trust.group"

.NOTES
    Versao: 1.0
    Autor: Sistema TI
    Data: 2025-01-15
#>

[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,
    
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Website,
    
    [Parameter()]
    [ValidateSet("ascii", "utf8", "unicode", "utf32")]
    [string]$Encoding = "unicode",
    
    [Parameter()]
    [ValidateSet("Error", "Warning", "Information", "Verbose")]
    [string]$LogLevel = "Information"
)

#region Configuration
$Script:Config = @{
    SignatureFolder = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'
    FileNameTemplate = "Trust Group ({0}@trust.group)"
    SupportedExtensions = @(".htm", ".html", ".txt", ".rtf")
    Groups = @{
        FotoComercial = "Assinatura Email - Foto Comercial"
        Comex = "Assinatura Email - Comex"
        Externo = "Assinatura Email - Externo"
    }
    ImageBaseUrl = "https://trust-signatures.s3.sa-east-1.amazonaws.com"
    DefaultOutlookVersions = @("16.0", "15.0", "14.0") # Office 365/2019/2016, 2013, 2010
}
#endregion

#region Classes
class UserAccount {
    [string]$Name
    [string]$DistinguishedName
    [string]$UserPrincipalName
    [string]$DisplayName
    [string]$GivenName
    [string]$Initials
    [string]$Surname
    [string]$PhysicalDeliveryOfficeName
    [string]$Description
    [string]$JobTitle
    [string]$Department
    [string]$Company
    [string]$EmailAddress
    [string]$StreetAddress
    [string]$City
    [string]$State
    [string]$PostalCode
    [string]$Country
    [string]$TelephoneNumber
    [string]$Mobile
    [string]$Pager
    [string]$Fax
    [string]$HomePhoneNumber
    [string]$OtherHomePhoneNumber
    [string]$HomeFax
    [string]$OtherFax
    [string]$IPPhone
    [string]$OtherIPPhone
    [string]$WebPage
    [string]$ExtensionAttribute1
    [string]$ExtensionAttribute2
    [string]$ExtensionAttribute3
    [bool]$IsFotoComercial
    [bool]$IsComex
    [bool]$IsExterno
    
    UserAccount() {
        # Constructor vazio para permitir inicializacao
    }
}

class SignatureGenerator {
    [UserAccount]$User
    [hashtable]$Config
    
    SignatureGenerator([UserAccount]$user, [hashtable]$config) {
        $this.User = $user
        $this.Config = $config
    }
    
    [string] GenerateHTML() {
        if ($this.User.IsFotoComercial) {
            return $this.GetFotoComercialSignature()
        }
        elseif ($this.User.IsComex) {
            return $this.GetComexSignature()
        }
        elseif ($this.User.IsExterno) {
            return $this.GetExternoSignature()
        }
        else {
            return $this.GetDefaultSignature()
        }
    }
    
    [string] GeneratePlainText() {
        return @"
$($this.User.DisplayName)
$($this.User.JobTitle)
___________________________________________________________

$($this.User.Company)
$($this.User.PhysicalDeliveryOfficeName)
$(if($this.User.StreetAddress){ "$($this.User.StreetAddress), " })$(if($this.User.City){ "$($this.User.City), " })$(if($this.User.State){ "$($this.User.State), " })$($this.User.PostalCode)

$(if($this.User.TelephoneNumber){"T: $($this.User.TelephoneNumber)"})
$(if($this.User.Mobile){"M: $($this.User.Mobile)"})
$(if($this.User.EmailAddress){"E: $($this.User.EmailAddress)"})
$(if($this.User.WebPage){"W: $($this.User.WebPage)"})
"@
    }
    
    hidden [string] GetFotoComercialSignature() {
        $baseUrl = $this.Config.ImageBaseUrl
        return @"
<table><tr style="font-size: 16px"><td><img src="$($this.Config.ImageBaseUrl)/usuarios/$($this.User.GivenName)$($this.User.Surname).png" width="130"></td><td><div style="margin: 0px 20px 0px 30px"><span>Empresa Exemplo</span><br><span style="font-weight: 800">$($this.User.GivenName) $($this.User.Surname)</span><br><span>$($this.User.JobTitle)</span><br><br><span>$($this.User.PhysicalDeliveryOfficeName)</span><br><span>$($this.User.TelephoneNumber)</span><br></div></td><td><div style="margin: 0px 30px 0px 0px"><img src="$($this.Config.ImageBaseUrl)/logomarcas/logo.png" width="130"></div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/barra.png" width="3" height="150"></td><td><div style="margin: 0px 30px 0px 30px; font-size: 18px">Acreditamos em <br>relações virtuosas <br>e confiáveis.</div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/gptw.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ods.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ecovadis.png" width="50"></td></tr></table><br>$($this.GetSocialMediaLinks())<br>$($this.GetLegalDisclaimer())
"@
    }
    
    hidden [string] GetComexSignature() {
        $baseUrl = $this.Config.ImageBaseUrl
        return @"
<table><tr style="font-size: 16px"><td><img src="$($this.Config.ImageBaseUrl)/usuarios/$($this.User.GivenName)$($this.User.Surname).png" width="130"></td><td><div style="margin: 0px 20px 0px 30px"><span>Empresa Exemplo</span><br><span style="font-weight: 800">$($this.User.GivenName) $($this.User.Surname)</span><br><span>$($this.User.JobTitle)</span><br><br><span>$($this.User.PhysicalDeliveryOfficeName)</span><br><span>$($this.User.TelephoneNumber)</span><br></div></td><td><div style="margin: 0px 30px 0px 0px"><img src="$($this.Config.ImageBaseUrl)/logomarcas/logo.png" width="130"></div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/barra.png" width="3" height="150"></td><td><div style="margin: 0px 30px 0px 30px; font-size: 18px">Acreditamos em <br>relações virtuosas <br>e confiáveis.</div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/gptw.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ods.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ecovadis.png" width="50"></td></tr></table><br>$($this.GetSocialMediaLinks())<br>$($this.GetLegalDisclaimer())
"@
    }
    
    hidden [string] GetExternoSignature() {
        $baseUrl = $this.Config.ImageBaseUrl
        return @"
<table><tr style="font-size: 16px"><td><img src="$($this.Config.ImageBaseUrl)/usuarios/$($this.User.GivenName)$($this.User.Surname).png" width="130"></td><td><div style="margin: 0px 20px 0px 30px"><span>Empresa Exemplo</span><br><span style="font-weight: 800">$($this.User.GivenName) $($this.User.Surname)</span><br><span>$($this.User.JobTitle)</span><br><br><span>$($this.User.PhysicalDeliveryOfficeName)</span><br><span>$($this.User.TelephoneNumber)</span><br></div></td><td><div style="margin: 0px 30px 0px 0px"><img src="$($this.Config.ImageBaseUrl)/logomarcas/logo.png" width="130"></div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/barra.png" width="3" height="150"></td><td><div style="margin: 0px 30px 0px 30px; font-size: 18px">Acreditamos em <br>relações virtuosas <br>e confiáveis.</div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/gptw.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ods.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ecovadis.png" width="50"></td></tr></table><br>$($this.GetSocialMediaLinks())<br>$($this.GetLegalDisclaimer())
"@
    }
    
    hidden [string] GetDefaultSignature() {
        $baseUrl = $this.Config.ImageBaseUrl
        return @"
<table><tr style="font-size: 16px"><td><img src="$($this.Config.ImageBaseUrl)/usuarios/$($this.User.GivenName)$($this.User.Surname).png" width="130"></td><td><div style="margin: 0px 20px 0px 30px"><span>Empresa Exemplo</span><br><span style="font-weight: 800">$($this.User.GivenName) $($this.User.Surname)</span><br><span>$($this.User.JobTitle)</span><br><br><span>$($this.User.PhysicalDeliveryOfficeName)</span><br><span>$($this.User.TelephoneNumber)</span><br></div></td><td><div style="margin: 0px 30px 0px 0px"><img src="$($this.Config.ImageBaseUrl)/logomarcas/logo.png" width="130"></div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/barra.png" width="3" height="150"></td><td><div style="margin: 0px 30px 0px 30px; font-size: 18px">Acreditamos em <br>relações virtuosas <br>e confiáveis.</div></td><td><img src="$($this.Config.ImageBaseUrl)/logomarcas/gptw.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ods.png" width="50"><img src="$($this.Config.ImageBaseUrl)/logomarcas/ecovadis.png" width="50"></td></tr></table><br>$($this.GetSocialMediaLinks())<br>$($this.GetLegalDisclaimer())
"@
    }
    
    hidden [string] GetSocialMediaLinks() {
        $baseUrl = $this.Config.ImageBaseUrl
        return @"
<table width="850"><tr><td><div style="margin: 0px 0px 0px 30px"><a href="https://wa.me/$($this.User.Mobile)"><img src="$($this.Config.ImageBaseUrl)/logomarcas/whatsapp.png" width="40"></a></div></td><td><div style="margin: 0px 0px 0px 30px"><a href="https://www.google.com/search?q=empresa+exemplo"><img src="$($this.Config.ImageBaseUrl)/logomarcas/localizacao.png" width="40"></a></div></td><td><div style="margin: 0px 0px 0px 30px"><a href="https://empresaexemplo.com.br/"><img src="$($this.Config.ImageBaseUrl)/logomarcas/site.png" width="40"></a></div></td><td><div style="margin: 0px 0px 0px 30px"><a href="https://linkedin.com/company/empresaexemplo"><img src="$($this.Config.ImageBaseUrl)/logomarcas/linkedin.png" width="40"></a></div></td><td><div style="margin: 0px 0px 0px 30px"><a href="https://instagram.com/empresa_exemplo"><img src="$($this.Config.ImageBaseUrl)/logomarcas/instagram.png" width="40"></a></div></td><td><div style="margin: 0px 0px 0px 30px"><a href="tel:$($this.User.Fax)"><img src="$($this.Config.ImageBaseUrl)/logomarcas/fone.png" width="40"></a></div></td></tr></table>
"@
    }
    
    hidden [string] GetLegalDisclaimer() {
        return @"
<table><div style="font-size: 11px">O conteúdo desta mensagem é de propriedade da Empresa Exemplo S.A. e é destinado exclusivamente para seu uso, sendo vedado o encaminhamento para fins particulares.<br> A divulgação indevida do teor desta mensagem é crime nos termos da legislação vigente, pois resultará em violação de segredo profissional, empresarial ou de proteção<br> à marca da proprietária nos termos do inc. XXIX da CF/88 e Lei 9.279/96. Se você não for o destinatário da mesma, favor inutilizá-la. <br><br> The content of this message is owned by Empresa Exemplo S.A. and is intended solely for your own use, being forbidden its use for private purposes. The improper disclosure of <br> this message is a crime under current Brazilian law, as it will represent a violation of professional, business secrets and/or brand protection under the terms of inc. XXIX of <br> CF/88 and Law 9.279/96. As it is so, if you are not the addressee, please make it unusable.</div></table>
"@
    }
}
#endregion

#region Logging Functions
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet("Error", "Warning", "Information", "Verbose")]
        [string]$Level = "Information",
        
        [Parameter()]
        [System.Exception]$Exception
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    if ($Exception) {
        $logMessage += " - Exception: $($Exception.Message)"
    }
    
    switch ($Level) {
        "Error" { 
            Write-Error $logMessage
            if ($Exception) { 
                Write-Error $Exception.StackTrace 
            }
        }
        "Warning" { Write-Warning $logMessage }
        "Information" { Write-Information $logMessage -InformationAction Continue }
        "Verbose" { Write-Verbose $logMessage }
    }
}
#endregion

#region Core Functions
function Test-Prerequisites {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Verificando pre-requisitos..." -Level Verbose
        
        # Verifica se esta em dominio
        if (-not $env:USERDOMAIN -or $env:USERDOMAIN -eq $env:COMPUTERNAME) {
            throw "Este script deve ser executado em um ambiente de dominio."
        }
        
        # Verifica se pode acessar AD
        try {
            $null = ([ADSISearcher]"(&(objectCategory=person)(sAMAccountName=$env:USERNAME))").FindOne()
        }
        catch {
            throw "Nao foi possivel acessar o Active Directory. Verifique as permissoes."
        }
        
        Write-Log "Pre-requisitos verificados com sucesso." -Level Information
        return $true
    }
    catch {
        Write-Log "Falha na verificacao de pre-requisitos" -Level Error -Exception $_.Exception
        return $false
    }
}

function Test-UserGroupMembership {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$GroupName
    )
    
    try {
        Write-Log "Verificando membership do grupo: $GroupName" -Level Verbose
        
        $searcher = [ADSISearcher]"(&(objectCategory=person)(sAMAccountName=$env:USERNAME))"
        $user = $searcher.FindOne()
        
        if (-not $user) {
            Write-Log "Usuario nao encontrado no AD" -Level Warning
            return $false
        }
        
        $groups = $user.Properties.memberof
        $isMember = $groups | Where-Object { $_ -match [regex]::Escape($GroupName) }
        
        Write-Log "Resultado da verificacao do grupo '$GroupName': $([bool]$isMember)" -Level Verbose
        return [bool]$isMember
    }
    catch {
        Write-Log "Erro ao verificar membership do grupo '$GroupName'" -Level Warning -Exception $_.Exception
        return $false
    }
}

function Get-UserDetails {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Obtendo detalhes do usuario do Active Directory..." -Level Information
        
        $searcher = [ADSISearcher]"(&(objectCategory=User)(samaccountname=$env:username))"
        $result = $searcher.FindOne()
        
        if (-not $result) {
            throw "Usuario '$env:username' nao encontrado no Active Directory."
        }
        
        $user = $result.Properties
        Write-Log "Usuario encontrado: $($user['displayname'] -join '')" -Level Information
        
        $userAccount = [UserAccount]::new()
        $userAccount.Name = $user['name'] -join ''
        $userAccount.DistinguishedName = $user['distinguishedname'] -join ''
        $userAccount.UserPrincipalName = $user['userprincipalname'] -join ''
        $userAccount.DisplayName = Convert-ToHtmlEntities (Convert-SpecificLocations -Text ($user['displayname'] -join ''))
        $userAccount.GivenName = Convert-ToHtmlEntities (Convert-SpecificLocations -Text ($user['givenname'] -join ''))
        $userAccount.Initials = $user['initials'] -join ''
        $userAccount.Surname = Convert-ToHtmlEntities (Convert-SpecificLocations -Text ($user['sn'] -join ''))
        
        $rawOfficeName = if ($user['physicaldeliveryofficename']) { $user['physicaldeliveryofficename'] -join '' } else { "" }
        $userAccount.PhysicalDeliveryOfficeName = Convert-ToHtmlEntities (Convert-SpecificLocations -Text $rawOfficeName)
        
        $userAccount.Description = $user['description'] -join ''
        $userAccount.JobTitle = Convert-ToHtmlEntities ($user['title'] -join '')
        $userAccount.Department = Convert-ToHtmlEntities (Convert-SpecificLocations -Text ($user['department'] -join ''))
        $userAccount.EmailAddress = $user['mail'] -join ''
        $userAccount.StreetAddress = $user['streetaddress'] -join ''
        $userAccount.City = $user['l'] -join ''
        $userAccount.State = $user['st'] -join ''
        $userAccount.PostalCode = $user['postalcode'] -join ''
        $userAccount.Country = $user['c'] -join ''
        $userAccount.TelephoneNumber = $user['telephonenumber'] -join ''
        $userAccount.Mobile = $user['mobile'] -join ''
        $userAccount.Pager = $user['pager'] -join ''
        $userAccount.Fax = $user['facsimiletelephonenumber'] -join ''
        $userAccount.HomePhoneNumber = $user['homephone'] -join ''
        $userAccount.OtherHomePhoneNumber = $user['otherhomephone'] -join ''
        $userAccount.HomeFax = $user['homefax'] -join ''
        $userAccount.OtherFax = $user['otherfacsimiletelephonenumber'] -join ''
        $userAccount.IPPhone = $user['ipphone'] -join ''
        $userAccount.OtherIPPhone = $user['otheripphone'] -join ''
        $userAccount.Company = if ($CompanyName) { $CompanyName } else { $user['company'] -join '' }
        $userAccount.WebPage = if ($Website) { $Website } else { $user['wWWHomePage'] -join '' }
        $userAccount.ExtensionAttribute1 = $user['extensionattribute1'] -join ''
        $userAccount.ExtensionAttribute2 = $user['extensionattribute2'] -join ''
        $userAccount.ExtensionAttribute3 = $user['extensionattribute3'] -join ''
        
        # Verifica grupos
        $userAccount.IsFotoComercial = Test-UserGroupMembership -GroupName $Script:Config.Groups.FotoComercial
        $userAccount.IsComex = Test-UserGroupMembership -GroupName $Script:Config.Groups.Comex
        $userAccount.IsExterno = Test-UserGroupMembership -GroupName $Script:Config.Groups.Externo
        
        Write-Log "Detalhes do usuario obtidos com sucesso." -Level Information
        return $userAccount
    }
    catch {
        Write-Log "Erro ao obter detalhes do usuario" -Level Error -Exception $_.Exception
        throw
    }
}

function Convert-SpecificLocations {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text
    )
    
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }
    
    # Mapeamento sem acentos para evitar problemas de codificacao
    $result = $Text
    $result = $result -replace "Filial Itajai", "Filial Itaja&iacute;"
    $result = $result -replace "Matriz Itajai", "Matriz Itaja&iacute;"
    $result = $result -replace "Filial Sao Paulo", "Filial S&atilde;o Paulo"
    $result = $result -replace "Filial Sumare", "Filial Sumar&eacute;"
    $result = $result -replace "Comercio Exterior", "Com&eacute;rcio Exterior"
    $result = $result -replace "Distribuicao", "Distribui&ccedil;&atilde;o"
    $result = $result -replace "Juridico", "Jur&iacute;dico"
    $result = $result -replace "Regulatorio", "regulat&oacute;rio"
    $result = $result -replace "Joao", "Jo&atilde;o"
    $result = $result -replace "Cezar", "C&eacute;zar"
    $result = $result -replace "Goncalves", "Gon&ccedil;alves"
    $result = $result -replace "Marcal", "Mar&ccedil;al"
    $result = $result -replace "Carvalho", "&Ccedil;arvalho"
    $result = $result -replace "Sao", "S&atilde;o"
    
    return $result
}

function Convert-ToHtmlEntities {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text
    )
    
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }
    
    # Funcao simplificada usando apenas codigos ASCII
    $result = $Text
    
    # Conversao usando char codes para evitar problemas de codificacao
    $result = $result -replace [char]193, "&Aacute;"  # A com acento agudo
    $result = $result -replace [char]225, "&aacute;"  # a com acento agudo
    $result = $result -replace [char]192, "&Agrave;"  # A com acento grave
    $result = $result -replace [char]224, "&agrave;"  # a com acento grave
    $result = $result -replace [char]194, "&Acirc;"   # A com circunflexo
    $result = $result -replace [char]226, "&acirc;"   # a com circunflexo
    $result = $result -replace [char]195, "&Atilde;"  # A com til
    $result = $result -replace [char]227, "&atilde;"  # a com til
    $result = $result -replace [char]201, "&Eacute;"  # E com acento agudo
    $result = $result -replace [char]233, "&eacute;"  # e com acento agudo
    $result = $result -replace [char]200, "&Egrave;"  # E com acento grave
    $result = $result -replace [char]232, "&egrave;"  # e com acento grave
    $result = $result -replace [char]202, "&Ecirc;"   # E com circunflexo
    $result = $result -replace [char]234, "&ecirc;"   # e com circunflexo
    $result = $result -replace [char]205, "&Iacute;"  # I com acento agudo
    $result = $result -replace [char]237, "&iacute;"  # i com acento agudo
    $result = $result -replace [char]204, "&Igrave;"  # I com acento grave
    $result = $result -replace [char]236, "&igrave;"  # i com acento grave
    $result = $result -replace [char]206, "&Icirc;"   # I com circunflexo
    $result = $result -replace [char]238, "&icirc;"   # i com circunflexo
    $result = $result -replace [char]211, "&Oacute;"  # O com acento agudo
    $result = $result -replace [char]243, "&oacute;"  # o com acento agudo
    $result = $result -replace [char]210, "&Ograve;"  # O com acento grave
    $result = $result -replace [char]242, "&ograve;"  # o com acento grave
    $result = $result -replace [char]212, "&Ocirc;"   # O com circunflexo
    $result = $result -replace [char]244, "&ocirc;"   # o com circunflexo
    $result = $result -replace [char]213, "&Otilde;"  # O com til
    $result = $result -replace [char]245, "&otilde;"  # o com til
    $result = $result -replace [char]218, "&Uacute;"  # U com acento agudo
    $result = $result -replace [char]250, "&uacute;"  # u com acento agudo
    $result = $result -replace [char]217, "&Ugrave;"  # U com acento grave
    $result = $result -replace [char]249, "&ugrave;"  # u com acento grave
    $result = $result -replace [char]219, "&Ucirc;"   # U com circunflexo
    $result = $result -replace [char]251, "&ucirc;"   # u com circunflexo
    $result = $result -replace [char]199, "&Ccedil;"  # C com cedilha
    $result = $result -replace [char]231, "&ccedil;"  # c com cedilha
    $result = $result -replace [char]209, "&Ntilde;"  # N com til
    $result = $result -replace [char]241, "&ntilde;"  # n com til
    
    return $result
}

function Remove-OldSignatureFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$BasePath
    )
    
    try {
        Write-Log "Removendo arquivos de assinatura antigos..." -Level Information
        
        $removedCount = 0
        foreach ($extension in $Script:Config.SupportedExtensions) {
            $targetFile = "$BasePath$extension"
            if (Test-Path $targetFile) {
                Remove-Item $targetFile -Force
                Write-Log "Arquivo removido: $targetFile" -Level Verbose
                $removedCount++
            }
        }
        
        if ($removedCount -gt 0) {
            Write-Log "Removidos $removedCount arquivo(s) antigo(s)." -Level Information
        } else {
            Write-Log "Nenhum arquivo antigo encontrado para remocao." -Level Verbose
        }
    }
    catch {
        Write-Log "Erro ao remover arquivos antigos" -Level Warning -Exception $_.Exception
    }
}

function New-SignatureDirectory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    try {
        if (-not (Test-Path -Path $Path)) {
            Write-Log "Criando diretorio de assinaturas: $Path" -Level Information
            $null = New-Item -ItemType Directory -Path $Path -Force
            Write-Log "Diretorio criado com sucesso." -Level Information
        } else {
            Write-Log "Diretorio de assinaturas ja existe." -Level Verbose
        }
        return $true
    }
    catch {
        Write-Log "Erro ao criar diretorio de assinaturas" -Level Error -Exception $_.Exception
        return $false
    }
}

function Save-SignatureFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$BasePath,
        
        [Parameter(Mandatory)]
        [string]$HtmlContent,
        
        [Parameter(Mandatory)]
        [string]$PlainTextContent,
        
        [Parameter(Mandatory)]
        [string]$TextEncoding
    )
    
    try {
        # Salva arquivo HTML
        Write-Log "Salvando assinatura HTML..." -Level Information
        $utf8Encoding = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText("$BasePath.htm", $HtmlContent, $utf8Encoding)
        Write-Log "Assinatura HTML salva: $BasePath.htm" -Level Information
        
        # Salva arquivo de texto
        Write-Log "Salvando assinatura de texto..." -Level Information
        $PlainTextContent | Out-File "$BasePath.txt" -Encoding $TextEncoding
        Write-Log "Assinatura de texto salva: $BasePath.txt" -Level Information
        
        return $true
    }
    catch {
        Write-Log "Erro ao salvar arquivos de assinatura" -Level Error -Exception $_.Exception
        return $false
    }
}

function Get-OutlookVersion {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Detectando versao do Outlook..." -Level Verbose
        
        foreach ($version in $Script:Config.DefaultOutlookVersions) {
            $regPath = "HKCU:\Software\Microsoft\Office\$version\Outlook"
            if (Test-Path $regPath) {
                Write-Log "Outlook versao $version detectado." -Level Information
                return $version
            }
        }
        
        # Fallback para versao mais recente
        Write-Log "Versao especifica nao detectada, usando fallback: $($Script:Config.DefaultOutlookVersions[0])" -Level Warning
        return $Script:Config.DefaultOutlookVersions[0]
    }
    catch {
        Write-Log "Erro ao detectar versao do Outlook" -Level Warning -Exception $_.Exception
        return $Script:Config.DefaultOutlookVersions[0]
    }
}

function Set-RegistryEntries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Version,
        
        [Parameter(Mandatory)]
        [string]$SignatureName
    )
    
    try {
        Write-Log "Configurando entradas de registro para Outlook versao $Version..." -Level Information
        
        $regPaths = @{
            General = "HKCU:\Software\Microsoft\Office\$Version\Common\General"
            MailSettings = "HKCU:\Software\Microsoft\Office\$Version\Common\MailSettings"
            OutlookSetup = "HKCU:\Software\Microsoft\Office\$Version\Outlook\Setup"
        }
        
        # Garante que as chaves existam
        foreach ($path in $regPaths.Values) {
            if (-not (Test-Path $path)) {
                $null = New-Item -Path $path -Force
                Write-Log "Chave de registro criada: $path" -Level Verbose
            }
        }
        
        # Define entradas de registro
        $null = New-ItemProperty -Path $regPaths.General -Name "Signatures" -Value "signatures" -PropertyType String -Force
        $null = New-ItemProperty -Path $regPaths.MailSettings -Name "NewSignature" -Value $SignatureName -PropertyType String -Force
        $null = New-ItemProperty -Path $regPaths.MailSettings -Name "ReplySignature" -Value $SignatureName -PropertyType String -Force
        
        Write-Log "Entradas de registro configuradas com sucesso." -Level Information
        
        # Remove First-Run se existir
        try {
            if (Test-Path $regPaths.OutlookSetup) {
                Remove-ItemProperty -Path $regPaths.OutlookSetup -Name "First-Run" -ErrorAction SilentlyContinue
                Write-Log "Chave 'First-Run' removida." -Level Verbose
            }
        }
        catch {
            Write-Log "Aviso: Nao foi possivel remover a chave 'First-Run'" -Level Verbose
        }
        
        return $true
    }
    catch {
        Write-Log "Erro ao configurar entradas de registro" -Level Error -Exception $_.Exception
        return $false
    }
}
#endregion

#region Main Script
function Invoke-SignatureGeneration {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "=== Iniciando geracao de assinatura de email ===" -Level Information
        
        # Verificacoes iniciais
        if (-not (Test-Prerequisites)) {
            throw "Falha na verificacao de pre-requisitos."
        }
        
        # Obter detalhes do usuario
        $user = Get-UserDetails
        if (-not $user) {
            throw "Nao foi possivel obter detalhes do usuario."
        }
        
        # Preparar diretorio
        if (-not (New-SignatureDirectory -Path $Script:Config.SignatureFolder)) {
            throw "Nao foi possivel criar o diretorio de assinaturas."
        }
        
        # Preparar arquivos
        $fileName = $Script:Config.FileNameTemplate -f $env:USERNAME
        $filePath = Join-Path -Path $Script:Config.SignatureFolder -ChildPath $fileName
        
        Remove-OldSignatureFiles -BasePath $filePath
        
        # Gerar assinaturas
        Write-Log "Gerando assinaturas..." -Level Information
        $signatureGenerator = [SignatureGenerator]::new($user, $Script:Config)
        $htmlSignature = $signatureGenerator.GenerateHTML()
        $plainTextSignature = $signatureGenerator.GeneratePlainText()
        
        # Salvar arquivos
        if (-not (Save-SignatureFiles -BasePath $filePath -HtmlContent $htmlSignature -PlainTextContent $plainTextSignature -TextEncoding $Encoding)) {
            throw "Erro ao salvar arquivos de assinatura."
        }
        
        # Configurar Outlook
        $outlookVersion = Get-OutlookVersion
        if (-not (Set-RegistryEntries -Version $outlookVersion -SignatureName $fileName)) {
            Write-Log "Aviso: Nao foi possivel configurar o Outlook automaticamente." -Level Warning
        }
        
        Write-Log "=== Assinatura gerada com sucesso ===" -Level Information
        Write-Log "Arquivos salvos em: $filePath" -Level Information
        
        return $true
    }
    catch {
        Write-Log "Erro critico durante a geracao da assinatura" -Level Error -Exception $_.Exception
        return $false
    }
}

# Execucao principal
try {
    $result = Invoke-SignatureGeneration
    if ($result) {
        exit 0
    } else {
        exit 1
    }
}
catch {
    Write-Log "Erro nao tratado" -Level Error -Exception $_.Exception
    exit 1
}
#endregion