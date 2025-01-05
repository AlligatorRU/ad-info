#powershell
#Вывод сообщения о начале работы сценария
#импорт модуля activedirectory, если модуль не доступен, вывод сообщения об ошибке
Import-Module activedirectory -ErrorAction SilentlyContinue
if (Get-Module -name ActiveDirectory -ErrorAction SilentlyContinue)
    {

    write-host -ForegroundColor DarkGreen Получение сведений из базы AD...
    $date_with_offset=(Get-Date).AddDays(-30)
    #Находим объекты компьютера, которые выполняли вход не поздней чем 30 дней назад.
    $PC=Get-ADComputer -SearchBase 'OU=computers,DC=domain,DC=com' -Filter {LastLogonDate -gt $date_with_offset} -Properties * |
    where {$_.Name  -notmatch 'computerName'} #notmatch исключает из поиска компьютер с указанным именем
        $PC | ForEach-Object {
            if (Test-Connection -Count 1 -ComputerName $_.Name -Quiet) #проверка доступен ли компьютер
                { $username=(Get-WmiObject -Class Win32_ComputerSystem -ComputerName $_.Name).username #получим имя пользователя AD выполнившего вход
            if ($username -ne $null)
                    {
                    $usr=$username.Remove(0,4)
                    $cmp=$_.Name
                    #устанавливаем атрибуты
                    Set-ADUser -Identity $usr -Description $cmp
                    Set-ADComputer -Identity $cmp -Description $usr
                    #сообщение о выполнении сценария
                    Write-Host -ForegroundColor green "атрибут успешно записан | пользователь" $usr "выполнил вход на компьютер" $_.Name
                    } else { Write-Host -ForegroundColor Gray $_.Name вход не выполнен }
                } else { Write-Host -ForegroundColor DarkGray $_.Name не в сети }
        }

    }
else {
write-host -ForegroundColor Red "Модуль Active Directory для Windows PowerShell не установлен.`n
Для установки модуля требуется набор утилит Microsoft Remote Server Administration Tools (RSAT)."
}
