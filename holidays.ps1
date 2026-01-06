$Year = 2026

# --- Outlook setup ---
$outlook   = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar  = $namespace.GetDefaultFolder(9) # olFolderCalendar
$items     = $calendar.Items
$items.IncludeRecurrences = $true

# --- Helper functions ---

function Get-NthWeekday {
    param ($Year, $Month, $DayOfWeek, $Nth)

    $date = Get-Date "$Year-$Month-01"
    while ($date.DayOfWeek -ne $DayOfWeek) {
        $date = $date.AddDays(1)
    }
    return $date.AddDays(7 * ($Nth - 1))
}

function Get-LastWeekday {
    param ($Year, $Month, $DayOfWeek)

    $date = (Get-Date "$Year-$Month-01").AddMonths(1).AddDays(-1)
    while ($date.DayOfWeek -ne $DayOfWeek) {
        $date = $date.AddDays(-1)
    }
    return $date
}

function Get-ObservedDate {
    param ([datetime]$Date)

    switch ($Date.DayOfWeek) {
        'Saturday' { return $Date.AddDays(-1) }
        'Sunday'   { return $Date.AddDays(1) }
        default    { return $Date }
    }
}

function Add-OrUpdateHoliday {
    param (
        [string]$Name,
        [datetime]$Date
    )

    $Observed = Get-ObservedDate $Date

    $existing = $items | Where-Object {
        $_.Subject -eq $Name -and $_.Start.Date -eq $Observed.Date
    }

    if ($existing) {
        foreach ($appt in $existing) {
            $appt.AllDayEvent = $true
            $appt.BusyStatus  = 3      # Out of Office
            $appt.ReminderSet = $false
            $appt.Categories  = "Holiday"
            $appt.Save()
        }
        Write-Host "Updated: $Name ($($Observed.ToShortDateString()))"
    }
    else {
        $appt = $calendar.Items.Add()
        $appt.Subject       = $Name
        $appt.Start         = $Observed
        $appt.AllDayEvent   = $true
        $appt.BusyStatus    = 3      # Out of Office
        $appt.ReminderSet   = $false
        $appt.Categories    = "Holiday"
        $appt.Save()

        Write-Host "Added: $Name ($($Observed.ToShortDateString()))"
    }
}

# --- Fixed-date holidays ---
Add-OrUpdateHoliday "New Year's Day"        (Get-Date "$Year-01-01")
Add-OrUpdateHoliday "Juneteenth"             (Get-Date "$Year-06-19")
Add-OrUpdateHoliday "Independence Day"       (Get-Date "$Year-07-04")
Add-OrUpdateHoliday "Veterans Day"           (Get-Date "$Year-11-11")
Add-OrUpdateHoliday "Christmas Day"          (Get-Date "$Year-12-25")

# --- Floating holidays ---
Add-OrUpdateHoliday "Martin Luther King Jr. Day" (Get-NthWeekday $Year 1  Monday   3)
Add-OrUpdateHoliday "Presidents Day"              (Get-NthWeekday $Year 2  Monday   3)
Add-OrUpdateHoliday "Memorial Day"                (Get-LastWeekday $Year 5  Monday)
Add-OrUpdateHoliday "Labor Day"                   (Get-NthWeekday $Year 9  Monday   1)
Add-OrUpdateHoliday "Columbus Day"                (Get-NthWeekday $Year 10 Monday   2)
Add-OrUpdateHoliday "Thanksgiving Day"            (Get-NthWeekday $Year 11 Thursday 4)

Write-Host "US Bank Holidays processed for $Year."
