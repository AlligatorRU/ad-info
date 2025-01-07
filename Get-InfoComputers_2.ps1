#импорт модуля activedirectory, если модуль не доступен, вывод сообщения об ошибке
Import-Module activedirectory -ErrorAction SilentlyContinue
if (Get-Module -name ActiveDirectory -ErrorAction SilentlyContinue)
    {
#вывод комментраия
    $comment = @"
       #############################################
       #В поиск можно включить ФИО, имя компьютера,#
       #IP адрес, Версию ОС                        #
       #############################################
"@
    Write-Host -ForegroundColor DarkCyan $comment
    if ($args.Count -eq 0) 
        {[string]$search=Read-Host "Поиск"}
        else {$search=$args[0]}
#вывод сообщения о начале работы сценария
    Write-Host -ForegroundColor Yellow "выполняется поиск..."
#-notmatch 'OU=test,OU=computers,DC=domain,DC=com'

            write-host -ForegroundColor DarkGreen Получение сведений...
            $date_with_offset=(Get-Date).AddDays(-90)
            $PC=Get-ADComputer -Properties * -SearchBase 'OU=computers,DC=domain,DC=com' -Filter {LastLogonDate -gt $date_with_offset}  |
            where {$_.name, $_.IPv4Address, $_.operatingSystem, $_.description, $_.distinguishedName -match $search } |
            Sort-Object operatingSystem | ft -Property name, IPv4Address, operatingSystem, description, distinguishedName  -Autosize
            $PC_last_logon=Get-ADComputer -SearchBase 'OU=computers,DC=domain,DC=com'`
            -Properties * -Filter {LastLogonDate -lt $date_with_offset} |
            where {$_.name, $_.IPv4Address, $_.operatingSystem, $_.description, $_.distinguishedName -match $search } | Sort-Object operatingSystem |
             ft -Property name, IPv4Address, operatingSystem, description, distinguishedName -Autosize
            $PC_no_logon=Get-ADComputer -Server FUTURAMA -SearchBase 'OU=computers,DC=domain,DC=com' -Filter {LogonCount -eq 0 } -Properties * |
            where {$_.name, $_.IPv4Address, $_.operatingSystem, $_.description, $_.distinguishedName -match $search } | ft -Property operatingSystem,  description, distinguishedName
            ################################################################################################################
            #вывод информации на экран
            write-host -ForegroundColor DarkGray "=========================================================="
            write-host -ForegroundColor Green "`t Активные компьютеры" ($PC.count - '4')
            write-host -ForegroundColor DarkGray "==========================================================" 
            $PC
            write-host -ForegroundColor DarkGray "=========================================================="
            write-host -ForegroundColor Red "`t Давно нет в сети" ($PC_last_logon.count - '4')
            write-host -ForegroundColor DarkGray "=========================================================="
            $PC_last_logon
            write-host -ForegroundColor DarkGray "=========================================================="
            write-host -ForegroundColor Yellow "`t Никогда не подключались" ($PC_no_logon.count - '4')
            write-host -ForegroundColor DarkGray "=========================================================="
            $PC_no_logon
            write-host -ForegroundColor DarkGray "=========================================================="
            ################################################################################################################
    } else {
write-host -ForegroundColor Red "Модуль Active Directory для Windows PowerShell не установлен.`n
Для установки модуля требуется набор утилит Microsoft Remote Server Administration Tools (RSAT).`n
RSAT можно найти в папке \\SR05\soft\OS\RSAT."
}
