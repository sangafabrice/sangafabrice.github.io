$ModuleDir = $MyInvocation.MyCommand.Path -replace '\\[^\\]+$'
$Email = (Get-Item $ModuleDir).Parent.FullName + '\index.html'

Function Set-NewsLetter {
    Set-Location -Path $Script:ModuleDir
    $MjmlNoComment = '.\index-no-comment.mjml'
    Get-Content .\index.mjml |
    Where-Object { $_ -notlike '*<!--*-->*' } |
    Out-File $MjmlNoComment
    mjml $MjmlNoComment --config.minify true --output $Script:Email
    Remove-Item $MjmlNoComment -Force
    Set-Location -
}

Function Send-NewsLetter {
    Param (
        [Parameter(Mandatory=$true)]
        [Alias('To')]
        [string] $Receiver,
        [Parameter(Mandatory=$true)]
        [Alias('Host')]
        [string] $SmtpServer,
        [Alias('Port')]
        [int] $SmtpPort = 587
    )

    (Get-Credential) |
    ForEach-Object {
        @{
            To = $Receiver;
            Subject = 'Fabrice Sanga Summary Statement';
            Body = (Get-Content $Script:Email -Raw);
            SmtpServer = $SmtpServer;
            From = 'Fabrice Sanga <' + $_.UserName + '>';
            BodyAsHtml = $true;
            DeliveryNotificationOption = 'OnSuccess';
            Credential = $_;
            UseSsl = $true;
            Port = $SmtpPort;
        }
    } | 
    ForEach-Object { Send-MailMessage @_ }
}

Export-ModuleMember -Function '*-NewsLetter'