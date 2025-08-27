# Load the WSUS Administration module
Import-Module UpdateServices


# Datos de conexión WSUS (genéricos)
$serverName = "wsus.midominio.local"
$portNumber = 8080

# Conectarse al servidor WSUS
$wsus = Get-WsusServer -Name $serverName -PortNumber $portNumber

# Obtener todos los equipos del servidor WSUS
$computers = $wsus.GetComputerTargets()

# Crear una lista para almacenar los datos
$WSUSData = @()

# Iterar a través de cada equipo y obtener la fecha del último informe de estado
foreach ($computer in $computers) {
    $lastReport = $computer.LastReportedStatusTime.ToString("yyyy-MM-dd HH:mm:ss")
    $WSUSData += [PSCustomObject]@{
    Equipo         = $computer.Name
        LastReportDate = $lastReport
    }
}

# Exportar los datos a un archivo CSV en la misma ubicación que el archivo PS1
$csvPath = Join-Path -Path $PSScriptRoot -ChildPath "WSUSData.csv"
$WSUSData | Export-Csv -Path $csvPath -NoTypeInformation

Write-Output "WSUS data extraction completed. Data saved to $csvPath."