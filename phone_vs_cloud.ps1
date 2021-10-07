param(
    [string]$phoneName,
    [string]$phoneFolderPath,
    [string]$cloudFolderPath,
    [string]$destinationFolderPath,
    [string]$filter='.(jpg|mp4)$',
    [switch]$confirmCopy=$false
)

function Get-ShellProxy
{
    if( -not $global:ShellProxy)
    {
        $global:ShellProxy = new-object -com Shell.Application
    }
    $global:ShellProxy
}

function Get-Phone
{
    param($phoneName)
    $shell = Get-ShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" â€” the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    $shellItem = $shell.NameSpace(17).self
    $phone = $shellItem.GetFolder.items() | where { $_.name -eq $phoneName }
    return $phone
}

function Get-SubFolder
{
    param($parent,[string]$path)
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    $current = $parent
    foreach ($pathPart in $pathParts)
    {
        if ($pathPart)
        {
            $current = $current.GetFolder.items() | where { $_.Name -eq $pathPart }
        }
    }
    return $current
}


$phone = Get-Phone -phoneName $phoneName
$phoneFolder = Get-SubFolder -parent $phone -path $phoneFolderPath

$items = @( $phoneFolder.GetFolder.items() | where { $_.Name -match $filter } )

$phonePath = "$phoneName\$phoneFolderPath"
if ($items) {
    $totalItems = $items.count
    if ($totalItems -gt 0) {
        Write-Output "Processing path: $phonePath"
        Write-Output "Looking for files in: $cloudFolderPath"
        Write-Output "Will copy any missing item to: $destinationFolderPath"

        $shell = Get-ShellProxy
        $destinationFolder = $shell.Namespace($destinationFolderPath).self

        $count = 0;
        $num_missing = 0;
        foreach ($item in $items) {
            $fileName = $item.Name

            ++$count
            $percent = [int](($count * 100) / $totalItems)
            Write-Progress -Activity "Processing Files in $phonePath" `
                -status "Processing File ${count} / ${totalItems} (${percent}%)" `
                -CurrentOperation $fileName `
                -PercentComplete $percent

            $filter = "{0}*" -f $fileName
            $found_items = @( Get-ChildItem -Path $cloudFolderPath -Filter $filter -Recurse )
            if ($found_items.Length -eq 0) {
                $confirmed = $true
                if ($confirmCopy) {
                    $confirmation = Read-Host "$fileName seems to be missing from $cloudFolderPath "`
                                              "or any of its sub-folders. "`
                                              "Shall we copy it to $destinationFolderPath? (y/n)"
                    if ($confirmation -ne 'y') {
                        $confirmed = $false
                    }
                }
                if ($confirmed) {
                    Write-Output "Copying $fileName to $destinationFolderPath..."
                    $destinationFolder.GetFolder.CopyHere($item)
                    Write-Output "Copied"
                    ++$copied
                }
            }
        }
        Write-Output "$copied/$totalItems item(s) were copied to $destinationFolderPath"
    }
} else {
    Write-Output "No files for $phonePath"
}
