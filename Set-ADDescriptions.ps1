#powershell
<#Сценарий находит пользователей выолнивших вход на компьютер и записывает атрибут "Description" в объект AD. В объект пользователя Active Directory, в атрибут "Description", будет записано имя компьютера, на который пользователь выполнил вход. А в объект компьютера, в атрибут "Description", будет записано имя этого пользователя. Таким образом, у нас будет информация об имени компьютера в описании пользователя и информация об имени пользователя в описании компьютера. Сценарий можно добавитьв планировщик заданий, чтобы иметь актуальную информацию о выполнивших вход пользователях. Далее мы можем использовать атрибут "Description" в других сценариях, для извлечения информации.#>

#Вывод сообщения о начале работы сценария
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
