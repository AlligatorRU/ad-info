#импорт модуля ActiveDirectory, если модуль не доступен, вывод сообщения об ошибке
Import-Module activedirectory -ErrorAction SilentlyContinue
if (Get-Module -name ActiveDirectory -ErrorAction SilentlyContinue)
    { 
#вывод комментраия
    $comment = @"
       #############################################
       #В поиск можно включить ФИО, имя компьютера,#
       #отдел, адрес, должность или номер телефона.#
       #############################################
"@
    Write-Host -ForegroundColor DarkCyan $comment
#задать значение переменной $search
    if ($args.Count -eq 0) 
        {[string]$search=Read-Host "Поиск"}
        else {$search=$args[0]}
#вывод сообщения о начале работы сценария
    Write-Host -ForegroundColor Yellow "выполняется поиск..."
    $SObject=Get-ADUser -Filter {Enabled -eq $True}  -Properties * -SearchBase 'OU=users,DC=domain,DC=com' |
    where {$_.description,$_.DisplayName,$_.telephoneNumber,$_.Department,$_.Title,$_.l -match $search} | Sort-Object description
    $date_with_offset=(Get-Date).AddDays(-20)
    Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
       $SObject | ForEach-Object {
    if ($_.description -ne $null)
        {
#проверка доступен ли компьютер
        if (Test-Connection -Count 1 -ComputerName $_.description -Quiet)
            {
            $PC=Get-ADComputer -Identity $_.description -Properties ipv4Address 
            Write-Host -ForegroundColor white $_.description `t -NoNewline
            Write-Host -ForegroundColor green "OK" `t -NoNewline
            Write-Host -ForegroundColor white "IP" $PC.IPv4Address
            Write-Host -ForegroundColor Gray "ФИО:" $_.DisplayName
            Write-Host -ForegroundColor Gray "отдел:" $_.Department
            Write-Host -ForegroundColor Gray "телефон:" $_.telephoneNumber
            Write-Host -ForegroundColor Gray "должность:" $_.Title
           #Write-Host -ForegroundColor Gray "Почта:" $_.EmailAddress
            Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
            } else {
            $PC=Get-ADComputer -Identity $_.description -Properties ipv4Address
            Write-Host -ForegroundColor white $_.description `t -NoNewline
            if ($_.LastLogonDate -lt $date_with_offset){
                   Write-Host -ForegroundColor red  "давно не подключался" `t -NoNewline
                } 
            Write-Host -ForegroundColor white "IP" $PC.IPv4Address
                
            Write-Host -ForegroundColor Gray "ФИО:" $_.DisplayName
            Write-Host -ForegroundColor Gray "отдел:" $_.Department
            Write-Host -ForegroundColor Gray "телефон:" $_.telephoneNumber
            Write-Host -ForegroundColor Gray "должность:" $_.Title
           #Write-Host -ForegroundColor Gray "Почта:" $_.EmailAddress
            Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
            }
        } else {
        Write-Host -ForegroundColor red "имя компьютера неизвестно"
        Write-Host -ForegroundColor Gray "ФИО:" $_.DisplayName
        Write-Host -ForegroundColor Gray "отдел:" $_.Department
        Write-Host -ForegroundColor Gray "телефон:" $_.telephoneNumber
        Write-Host -ForegroundColor Gray "должность:" $_.Title
       #Write-Host -ForegroundColor Gray "Почта:" $_.EmailAddress
        Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
        }
    }
    if ($SObject.cn.count -eq 0) {
        Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "ни чего не найдено`n попробуте изменить запрос." `a
        } else {
            if ($SObject.cn.count -eq 1)
            {
            Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "найденa 1 запись" `a
            } else {
                if ($SObject.cn.count -match "[2-4]")
                {
                Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "найдено" $SObject.count "записи" `a
                } else {
                Write-Host -ForegroundColor Yellow "Поиск завершен, по запросу" $search "найдено" $SObject.count "записей" `a
                }

            }
        } 
    Write-Host -ForegroundColor DarkGray "____________________________________________________________________________"
    write-host -ForegroundColor Gray -BackgroundColor DarkGray "© Кинешемская ЦРБ | Олег Груздев | 2017 - 2020 г."
    } else {
write-host -ForegroundColor Red "Модуль Active Directory для Windows PowerShell не установлен.`n
Для установки модуля требуется набор утилит Microsoft Remote Server Administration Tools (RSAT)."
}
