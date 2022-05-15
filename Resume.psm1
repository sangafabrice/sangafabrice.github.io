$ModuleDir = (Get-Item $MyInvocation.MyCommand.Path).DirectoryName
$MJML = $(where.exe mjml 2> $Null) ? 'mjml':$(
    "$ModuleDir\node_modules\.bin\mjml.ps1" |
    ForEach-Object {
        If (!(Test-Path $_)) {
            Set-Location $ModuleDir
            [void] (npm install mjml@latest --yes)
            Set-Location -
        }
        $_
    }
)
$Email = git worktree list |
ForEach-Object {
    If ($_.EndsWith('[html-email]', 1)) {
        ($_ -replace '\s+[a-z0-9]{7,} \[html-email\]$').Trim() + '\index.html'
    }
}

Function Set-NewsLetter {
    Set-Location $Script:ModuleDir
    '.\index-no-comment.mjml' |
    ForEach-Object {
        Get-Content .\index.mjml |
        Where-Object { $_ -notlike '*<!--*-->*' } |
        Out-File $_
        Invoke-Expression "$MJML $_ --config.minify true --output ${Script:Email}"
        Remove-Item $_ -Force
    }
    Set-Location -
}

Function Send-NewsLetter {
    Param (
        [Parameter(Mandatory)]
        [Alias('To')]
        [string] $Receiver,
        [Parameter(Mandatory, ParameterSetName='UseCmdLine')]
        [Alias('Host')]
        [string] $SmtpServer,
        [Parameter(Mandatory, ParameterSetName='UseDefault')]
        [switch] $UseDefault,
        [Alias('Port')]
        [int] $SmtpPort = 587
    )
    
    $UseDefault ? $(
        Remove-Variable SmtpServer
        Get-Content "${Script:ModuleDir}\secret.toml" |
        ForEach-Object {
            ,($_ -split '=') |
            ForEach-Object {
                @{
                    Option = 'ReadOnly';
                    Name = $_[0].Trim();
                    Value = ($_[1] -replace '"').Trim();
                    Force = $true;
                }
            } |
            ForEach-Object { Set-Variable @_ }
        }
        [pscredential]::new(
            $UserName,
            (
                @{
                    String = $PassWord;
                    AsPlainText = $true;
                    Force = $true;
                } | 
                ForEach-Object { ConvertTo-SecureString @_ }
            )
        )
    ):(Get-Credential) |
    ForEach-Object {
        @{
            To = $Receiver;
            Subject = 'Fabrice Sanga Summary Statement';
            Body = (Get-Content $Script:Email -Raw -ErrorAction Stop);
            SmtpServer = $SmtpServer;
            From = 'Fabrice Sanga <' + $_.UserName + '>';
            BodyAsHtml = $true;
            DeliveryNotificationOption = 'OnSuccess';
            Credential = $_;
            UseSsl = $true;
            Port = $SmtpPort;
            WarningAction = 'SilentlyContinue';
        }
    } | 
    ForEach-Object { Send-MailMessage @_ }
}

Export-ModuleMember -Function '*-NewsLetter'