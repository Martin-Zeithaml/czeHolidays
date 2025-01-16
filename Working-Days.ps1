# ============ Modify section start ============
# Modify year and Easter holidays for that year
$year=2025
$czeHolidays=@(
    1,                          # Jan
    0,                          # Feb
    0,                          # Mar
    (18,21),                    # Apr
    (1,8),                      # May
    0,                          # Jun
    (5,6),                      # Jul
    0,                          # Aug
    28,                         # Sep
    28,                         # Oct
    17,                         # Nov
    (24,25,26)                  # Dec
)

# ============ Modify section end ==============

# Print out NETWORKDAYS with holidays (if any)
1..12 | ForEach-Object {
    # Get the last day of month by adding months to the last day of previous year
    $lastDay=(New-Object DateTime($($year-1),12,31)).AddMonths($_).ToString("yyyy,MM,dd")    
    # Zero is no holidays
    if ($czeHolidays[$_-1] -eq 0) {
        $holidays = ""
    }
    # Otherwise create holidays list
    else {
        $holidays = ", "
        # Opening '{' only if 2 or more days
        $holidays += "{" * [int][Math]::Sign($czeHolidays[$_-1].count-1)
        for ($i = 0; $i -lt $czeHolidays[$_-1].count; $i++) {
            $holidays += "date($($year), $($_), $($czeHolidays[$_-1][$i]))"
            if ($i+1 -eq $czeHolidays[$_-1].count) {
                # Closing '}' only if 2 or more days
                $holidays += "}" * [int][Math]::Sign($czeHolidays[$_-1].count-1)
            } else {
                $holidays += ", "
            }
        }
    }
    # Hours
    Write-Output "=NETWORKDAYS(date($($year),$_,1),date($($lastDay))$($holidays))*8"
}
