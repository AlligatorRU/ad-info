#импорт модуля activedirectory, если модуль не доступен, вывод сообщения об ошибке
Import-Module activedirectory -ErrorAction SilentlyContinue
if (Get-Module -name ActiveDirectory -ErrorAction SilentlyContinue)
{
################################################################################################################

#комментраий
$comment = @"
##################################################
#В поиск можно включить логин пользователя       #
#имя компьютера, IP адрес, подразделение.        #   
##################################################
"@
Write-Host -ForegroundColor DarkCyan $comment
#задать значение переменной $search
    if ($args.Count -eq 0) 
        {[string]$search=Read-Host "Поиск"}
        else {$search=$args[0]}
    Write-Host -ForegroundColor Yellow "выполняется поиск..."
#поиск объекта в AD (-notmatch 'OU=test,OU=computers,DC=domain,DC=com') можно исключить подразделение

write-host -ForegroundColor DarkGreen Получение сведений...
#дата минус 30 дней
$date_with_offset=(Get-Date).AddDays(-30)
#Подключение к контейнеру с компьютерами, выполнявшими вход не менее 30 дней назад
$PC=Get-ADComputer -Properties * -SearchBase 'OU=computers,DC=domain,DC=com' -Filter {LastLogonDate -gt $date_with_offset}  |
where {$_.name, $_.IPv4Address, $_.operatingSystem, $_.description, $_.CanonicalName -match $search } 
#Подключение к контейнеру с компьютерами, которых давно небыло видно
$PC_last_logon=Get-ADComputer -SearchBase 'OU=computers,DC=domain,DC=com' -Properties * -Filter {LastLogonDate -lt $date_with_offset} |
where {$_.name, $_.IPv4Address, $_.operatingSystem, $_.description, $_.CanonicalName -match $search } 
#Подключение к контейнеру с компьютерами, которые не были в сети
$PC_no_logon=Get-ADComputer -Server FUTURAMA -SearchBase 'OU=computers,DC=domain,DC=com' -Filter {LogonCount -eq 0 } -Properties * |
where {$_.name, $_.IPv4Address, $_.operatingSystem, $_.description, $_.CanonicalName -match $search }

#Запись переменных при условии что есть переменная description
$PC | ForEach-Object {
    if ($_.description -ne $null)
        {
            $name=$_.name
            $IPv4Address=$_.IPv4Address
            $description=$_.description
            $CanonicalName=$_.CanonicalName
            $operatingSystem=$_.operatingSystem
#ICMP запрос
            if (Test-Connection -Count 1 -ComputerName $IPv4Address -Quiet)
                {
                $status=write-host -ForegroundColor Green "ONLINE"
                }
                    else {
                    $status=write-host -ForegroundColor Red "OFFLINE"
                }

               
            write-host -ForegroundColor White "Имя компьютера" $name
            write-host -ForegroundColor White "IP-адрес" $IPv4Address 
            write-host -ForegroundColor White "Пользователь"$description
            write-host -ForegroundColor White "Подразделение " $CanonicalName
            $status

            $DisplayName
            $Department
            $Title
            $telephoneNumber
            Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
         }
}

$PC_last_logon | ForEach-Object {
    if ($_.description -ne $null)
        {
            $logon=write-host -ForegroundColor Red "Давно не подключался"
            $name=$_.name
            $IPv4Address=$_.IPv4Address
            $description=$_.description
            $CanonicalName=$_.CanonicalName
            $operatingSystem=$_.operatingSystem
            $logon
            write-host -ForegroundColor White "Имя компьютера" $name
            write-host -ForegroundColor White "IP-адрес" $IPv4Address 
            write-host -ForegroundColor White "Пользователь"$description
            write-host -ForegroundColor White "Подразделение " $CanonicalName
            Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
         }
}
$PC_no_logon | ForEach-Object {
    if ($_.description -ne $null)
        {
            $logon="ERROR"
            $name=$_.name
            $IPv4Address=$_.IPv4Address
            $description=$_.description
            $CanonicalName=$_.CanonicalName
            $operatingSystem=$_.operatingSystem
            write-host -ForegroundColor White $logon $name $IPv4Address $description $CanonicalName $operatingSystem
            Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
         }
}

    if ($PC.cn.count -eq 0) {
        Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "ни чего не найдено`n попробуте изменить запрос." `a
        } else {
            if ($PC.cn.count -eq 1)
            {
            Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "найденa 1 запись" `a
            } else {
                if ($PC.cn.count -match "[2-4]")
                {
                Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "найдено" $PC.count "записи" `a
                } else {
                Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "найдено" $PC.count "записей" `a
                }
            }
        } 
    Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
    write-host -ForegroundColor Gray -BackgroundColor DarkGray "© Кинешемская ЦРБ | Олег Груздев | 2017 - 2020 г."

################################################################################################################
    } else {
write-host -ForegroundColor Red "Модуль Active Directory для Windows PowerShell не установлен.`n
Для установки модуля требуется набор утилит Microsoft Remote Server Administration Tools (RSAT).`n
RSAT можно найти в папке \\FS\soft\OS\RSAT."
}
  
