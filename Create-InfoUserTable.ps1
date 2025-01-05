#импорт модуля activedirectory, если модуль не доступен, вывод сообщения об ошибке
Import-Module activedirectory -ErrorAction SilentlyContinue
if (Get-Module -name ActiveDirectory -ErrorAction SilentlyContinue)
    { 

#комментраий
    $comment = @"
       #############################################
       #В поиск можно включить ФИО, имя компьютера,#
       #отдел, должность или номер телефона.       #
       #############################################
"@
    Write-Host -ForegroundColor DarkCyan $comment
#задать значение переменной $search
    if ($args.Count -eq 0) 
        {[string]$search=Read-Host "Создать"}
        else {$search=$args[0]}

#поиск объекта в AD
    $SObject=Get-ADUser -Filter {Enabled -eq $True}  -Properties * -SearchBase 'OU=People,DC=CRB,DC=KIN' |
    where {$_.description,$_.DisplayName,$_.telephoneNumber,$_.Department,$_.Title,$_.l -match $search} | Sort-Object Department
    $date_with_offset=(Get-Date).AddDays(-30)
#поиск по объекту
#проверка существования атрибута
   # Созадём объект Excel
$Excel = New-Object -ComObject Excel.Application

# Делаем его видимым
$Excel.Visible = $true
# Добавляем рабочую книгу
$WorkBook = $Excel.Workbooks.Add()
$people = $WorkBook.Worksheets.Item(1)
$people.Name = 'Пользователи домена'

# Заголовок таблицы (самая первая ячейка)
$Row = 1
$Column = 1
$people.Cells.Item($Row, $Column) = 'Сведения о пользователях '+ $search

# Форматируем текст, чтобы он был похож на заголовок
$people.Cells.Item($Row, $Column).Font.Size = 16
$people.Cells.Item($Row, $Column).Font.Bold = $true
$people.Cells.Item($Row, $Column).Font.ThemeFont = 1
$people.Cells.Item($Row, $Column).Font.ThemeColor = 4
$people.Cells.Item($Row, $Column).Font.ColorIndex = 55
$people.Cells.Item($Row, $Column).Font.Color = 8210719

# Объединяем диапазон ячеек
$Range = $people.Range('A1','J2')
$Range.Merge()
$Range.VerticalAlignment = -4108

# Переходим на следующую строку
$Row++; $Row++

# Номер начальной строки
$InitialRow = $Row

# Заполняем ячейки - шапку таблицы
$people.Cells.Item($Row, $Column) = 'Отдел'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true		
$Column++
$people.Cells.Item($Row, $Column) = 'Имя'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true		
$Column++
$people.Cells.Item($Row, $Column) = 'Должность'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
$Column++
$people.Cells.Item($Row, $Column) = 'Телефон'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
$Column++
$people.Cells.Item($Row, $Column) = 'Мобильный'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
$Column++
$people.Cells.Item($Row, $Column) = 'Город'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
$Column++
$people.Cells.Item($Row, $Column) = 'Улица'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
$Column++
$people.Cells.Item($Row, $Column)= 'Почта'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
$Column++
$people.Cells.Item($Row, $Column) = 'Имя компьютера'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	

$Column++
$people.Cells.Item($Row, $Column) = 'Подключение'
$people.Cells.Item($Row, $Column).Interior.ColorIndex = 15
$people.Cells.Item($Row, $Column).Font.Bold = $true	
		

# Переходим на следующую строку, возвращаемся в первый столбец
$Row++
$Column = 1

       $SObject | ForEach-Object {
    if ($_.description -ne $null)
        {
        #ICMP запрос
        if (Test-Connection -Count 1 -ComputerName $_.description -Quiet)
            {
            $people.Cells.Item($Row, $Column) = $_.Department 
            $Column++
            $people.Cells.Item($Row, $Column) =  $_.DisplayName
            $Column++
            $people.Cells.Item($Row, $Column) = $_.Title
            $Column++
			$people.Cells.Item($Row, $Column).NumberFormat = "@" 
            $people.Cells.Item($Row, $Column) = $_.telephoneNumber
            $Column++
			$people.Cells.Item($Row, $Column).NumberFormat = "@" 
            $people.Cells.Item($Row, $Column) = $_.mobile
            $Column++
            $people.Cells.Item($Row, $Column) = $_.l
            $Column++
			$people.Cells.Item($Row, $Column) = $_.streetAddress
            $Column++
			$people.Cells.Item($Row, $Column) = $_.EmailAddress
            $Column++
            $people.Cells.Item($Row, $Column) = $_.description
            $Column++ 
            $people.Cells.Item($Row, $Column) = "в сети"
			$people.Cells.Item($Row, $Column).Interior.ColorIndex = 10
            $Row++
            $Column = 1 			
            } else {
            if ($_.LastLogonDate -lt $date_with_offset){
                   
                } 
            $people.Cells.Item($Row, $Column) = $_.Department 
            $Column++
            $people.Cells.Item($Row, $Column) =  $_.DisplayName
            $Column++
            $people.Cells.Item($Row, $Column) = $_.Title
            $Column++
			$people.Cells.Item($Row, $Column).NumberFormat = "@" 
            $people.Cells.Item($Row, $Column) = $_.telephoneNumber
            $Column++
			$people.Cells.Item($Row, $Column).NumberFormat = "@" 
            $people.Cells.Item($Row, $Column) = $_.mobile
            $Column++
            $people.Cells.Item($Row, $Column) = $_.l
            $Column++
			$people.Cells.Item($Row, $Column) = $_.streetAddress
            $Column++
			$people.Cells.Item($Row, $Column) = $_.EmailAddress
            $Column++
            $people.Cells.Item($Row, $Column) = $_.description
            $Column++ 
			$people.Cells.Item($Row, $Column) = "не в сети"
			$people.Cells.Item($Row, $Column).Interior.ColorIndex = 6
            $Row++
            $Column = 1          
            }
        } else {
            $people.Cells.Item($Row, $Column) = $_.Department 
            $Column++
            $people.Cells.Item($Row, $Column) =  $_.DisplayName
            $Column++
            $people.Cells.Item($Row, $Column) = $_.Title
            $Column++
			$people.Cells.Item($Row, $Column).NumberFormat = "@" 
            $people.Cells.Item($Row, $Column) = $_.telephoneNumber
            $Column++
			$people.Cells.Item($Row, $Column).NumberFormat = "@" 
            $people.Cells.Item($Row, $Column) = $_.mobile
            $Column++
            $people.Cells.Item($Row, $Column) = $_.l
            $Column++
			$people.Cells.Item($Row, $Column) = $_.streetAddress
            $Column++
			$people.Cells.Item($Row, $Column) = $_.EmailAddress
            $Column++
            $people.Cells.Item($Row, $Column) = $_.description
            $Column++ 
			$people.Cells.Item($Row, $Column) = "давно не подключался"
			$people.Cells.Item($Row, $Column).Interior.ColorIndex = 3
            $Row++
            $Column = 1          
        }
    }
	# Возвращаемся на одну строку назад
$Row--

# Выделяем нашу таблицу
$DataRange = $people.Range(("A{0}" -f $InitialRow), ("J{0}" -f $Row))
7..12 | ForEach-Object `
{
    $DataRange.Borders.Item($_).LineStyle = 1
    $DataRange.Borders.Item($_).Weight = 2
}
$UsedRange = $people.UsedRange
$UsedRange.EntireColumn.AutoFit() | Out-Null

  } else {
write-host -ForegroundColor Red "Модуль Active Directory для Windows PowerShell не установлен.`n
Для установки модуля требуется набор утилит Microsoft Remote Server Administration Tools (RSAT).`n
RSAT можно найти в папке \\FS\soft\OS\RSAT."
}
