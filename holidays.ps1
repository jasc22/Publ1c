$Year = 2026

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9) # Calendar
$items = $calendar.Items
$items.IncludeRecurrences = $true

function Add-Holiday {
    param (
        [string]$Name,
        [datetime]$Date
    )

    # Adjust for observed date
    if ($Date.DayOfWeek -eq 'Saturday') {
        $Observed = $Date.AddDays(-1)
    }
    elseif ($Date.DayOfWeek -eq 'Sunday') {
        $Observed = $Date.AddDays(1)
    }
    else {
        $Observed = $Date
    }

    $exists = $items | Where-Object {
        $_.Subject -eq $Name -and $_.Start.Date -eq $Observed.Date
    }

    if (-not $exists) {
        $appt = $calendar.Items.Add()
        $appt.Subject = $Name
        $appt.Start = $Observed
        $appt.AllDayEvent = $true
        $appt.BusyStatus = 0
        $appt.ReminderSet = $false
        $appt.Categories = "Holiday"
        $appt.Save()

        Write-Host "Added: $Name ($($Observed.ToShortDateString()))"
    }
    else {
        Write-Host "Skipped: $Name already exists"
    }
}

function Get-NthWeekday {
    param ($Year, $Month, $DayOfWeek, $Nth)
    $date = Get-Date "$Year-$Month-01"
    while ($date.DayOfWeek -ne $DayOfWeek) { $date = $date.AddDays(1) }
    return $date.AddDays(7 * ($Nth - 1))
}

function Get-LastWeekday {
    param ($Year, $Month, $DayOfWeek)
    $date = (Get-Date "$Year-$Month-01").AddMonths(1).AddDays(-1)
    while ($date.DayOfWeek -ne $DayOfWeek) { $date = $date.AddDays(-1) }
    return $date
}

# Fixed-date holidays
Add-Holiday "New Year's Day"        (Get-Date "$Year-01-01")
Add-Holiday "Juneteenth"             (Get-Date "$Year-06-19")
Add-Holiday "Independence Day"       (Get-Date "$Year-07-04")
Add-Holiday "Veterans Day"           (Get-Date "$Year-11-11")
Add-Holiday "Christmas Day"          (Get-Date "$Year-12-25")

# Floating holidays
Add-Holiday "Martin Luther King Jr. Day" (Get-NthWeekday $Year 1 Monday 3)
Add-Holiday "Presidents Day"              (Get-NthWeekday $Year 2 Monday 3)
Add-Holiday "Memorial Day"                (Get-LastWeekday $Year 5 Monday)
Add-Holiday "Labor Day"                   (Get-NthWeekday $Year 9 Monday 1)
Add-Holiday "Columbus Day"                (Get-NthWeekday $Year 10 Monday 2)
Add-Holiday "Thanksgiving Day"            (Get-NthWeekday $Year 11 Thursday 4)

Write-Host "US Bank Holidays added for $Year."
