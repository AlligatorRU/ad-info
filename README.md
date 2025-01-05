# ad-info
<h3>1. Active Directory, добавление имени пользователя и компьютера в атрибут "Description". </h3>
<h3>2. Поиск пользователей и компьютеров. </h3>


<p>1. В Active Directory нет специального атрибута, где бы хранилась информация о том, кто выполнил вход на компьютер. Но имеется атрибут есть атрибут "Description", который мы можем использовать для записи информации о пользователях и компьютерах. Сценарий Set-ADDescriptions.ps1 добавляет в объекты компьютера информацию о пользователе, который выполнил вход на данный компьютер, а в объект пользователя информацию о компьютере на который он вошёл. Таким образом, у нас будет информация об имени компьютера в описании пользователя и информация об имени пользователя в описании компьютера. Сценарий можно добавить в планировщик заданий, чтобы иметь актуальную информацию в описании. Далее атрибут "Description" будет удобно использовать для поиска пользователей и компьютеров в других сценариях.</p>

<p>2. Сценарий  Get-InfoUsers.ps1 выполняет поиск по разным атрибутам в объектах пользователей, например: description, DisplayName, telephoneNumber, Department, Title. Get-InfoComputers.ps1 выполняет поиск в по разным атрибутам в объектах компьютеров, например: name, IPv4Address, operatingSystem, description, CanonicalName. Get-InfoComputers_2.ps1 выводит информацию о компьютерах, которые не подключались к сети или подключались давно. Сценарий Get-InfoUsersTable.ps1 работает аналогично Get-InfoUsers.ps1, но выводит информацию в таблицу excel.</p>
