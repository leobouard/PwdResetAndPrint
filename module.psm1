function New-Password {
    param([int]$Length = 24)

    [string]$password = ''
    $low = 'abcdefghijkmnopqrstuvwxyz'.ToCharArray() 
    $upp = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $spe = '!#$%&*+-./=?@_'.ToCharArray()
    $num = 1..9
    $all = $low + $upp + $spe + $num

    # Get one character of every type
    'low', 'upp', 'spe', 'num' | ForEach-Object {
        $password += Get-Variable $_ -ValueOnly | Get-Random
    }
    while ($password.Length -lt $Length) { $password += $all | Get-Random }
    ($password.ToCharArray() | Get-Random -Count $Length) -join ''
}

function ConvertTo-HTMLPasswordTable {
    param(
        [string]$Password,
        [int]$Index = 1    
    )

    $i = $Index
    $passwordTable = [PSCustomObject]@{}
    $Password -split '' | Where-Object { $_ } | ForEach-Object {
        $passwordTable | Add-Member -MemberType NoteProperty -Name $i -Value $_
        $i++
    }

    [string]($passwordTable | ConvertTo-Html -Fragment)
}

function Add-Emphasis {
    param([string]$HTMLTable)

    'low', 'upp', 'num', 'spe' | ForEach-Object {
        $class = $_
        $regex = switch ($class) {
            'low' { '<td>[a-z]</td>' }
            'upp' { '<td>[A-Z]</td>' }
            'num' { '<td>[0-9]</td>' }
            'spe' { '<td>' }
        }

        ($HTMLTable | Select-String $regex -CaseSensitive -AllMatches).Matches | Sort-Object -Unique -Property Value | ForEach-Object {
            $td = $_.Value
            $newTd = $td -replace '<td>', "<td class=`"$class`">"
            $HTMLTable = $HTMLTable -creplace $td, $newTd
        }
    }

    $HTMLTable
}

function Split-Password {
    param([string]$Password)

    $length = $Password.Length
    $half = [int]($length / 2)
    $end = $length - $half

    @($Password.Substring(0, $half); $Password.Substring($half, $end))
}