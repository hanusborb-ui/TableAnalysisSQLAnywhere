<#
.SYNOPSIS
Read-only SQL Anywhere 17 table inventory, eligibility checks, and CSV batching.
.DESCRIPTION
Requires Windows PowerShell 5.1 and the SQL Anywhere 17 ODBC driver matching
PowerShell bitness. Defaults to Windows integrated login (INT=YES), using the
Windows identity running PowerShell. -IntegratedSecurity explicitly selects
this mode. Supply -Credential instead for database authentication (INT=NO).
These authentication options are mutually exclusive. Integrated authentication
requires the database login_mode and Windows login mapping already configured.
Uses SQL Anywhere ODBC, not SQL Server SqlClient.
ServerName accepts host names or host:port endpoints; DatabaseName is required
and applies to each endpoint. Each database is connected independently; there
is no SQL Server master database discovery. EngineName optionally identifies
the engine (same value for every endpoint in one invocation).
Encryption accepts SQL Anywhere ENC syntax, e.g.
TLS(trusted_certificates=C:\Certs\root.crt). If omitted the driver/server defaults
apply. SQL Server's TrustServerCertificate switch is not supported.
Only local persistent base tables are included. System-owner tables, views,
materialized views, temporary tables and remote proxy tables are excluded.
RowCount is the catalog count from SYSTAB, updated at successful checkpoints;
it can lag live data. No COUNT(*) scans or checkpoints are executed.
DateTimeColumn1 and DateTimeColumn2 contain the first two datetime/timestamp
column names in table column order (including aliases and time-zone variants).
Missing columns are blank. Only column names are returned, not data values.
PK contains the primary key column names, comma-separated in key order,
from SYSIDX/SYSIDXCOL/SYSTABCOL, without square brackets;
tables without a primary key show No. Eligibility uses a separate existence
flag, so changing the displayed PK name does not change Is_Mirror. XML checks both built-in and named XML types. Arbitrary
text columns containing XML cannot be identified by this metadata-only check.
HasVarbinaryColumn conservatively flags the binary family (BINARY/VARBINARY,
LONG BINARY/IMAGE), including aliases through their base domain.
SQL Server memory-optimized and partition-scheme checks have no direct mapping
here: their CSV fields are N/A, not tested No values. Is_Mirror means only that
a primary key exists and no XML/binary columns were detected. It does not certify
compatibility with a target mirroring/replication product, or assess SQL Anywhere
database-level in-memory mode. Confirm product support separately.
Batches restart at 1 per database/server and keep Yes/No tables separate.
Tables >= BatchRowLimit are isolated; other tables use first-fit decreasing.
No table is split. One CSV per server/database, with both names in the filename.
CsvPath supplies the base filename; server and database names are appended.
CsvPath and OutputDirectory are exclusive.
Credentials must have complete catalog visibility. Failed databases are warned
and omitted; a failed run must not be treated as a complete inventory.
.EXAMPLE
.\Get-TableMirrorCheck-SQLAnywhere17.ps1 -ServerName 'prov-s-7118:40901' -DatabaseName 'DB_CHR' -IntegratedSecurity -OutputDirectory 'C:\Temp\TableChecks'
.EXAMPLE
$cred = Get-Credential -UserName 'DBA' -Message 'SQL Anywhere credentials'
.\Get-TableMirrorCheck-SQLAnywhere17.ps1 -ServerName 'sa-host:2638' -DatabaseName 'AppDB' -Credential $cred -OutputDirectory 'C:\Temp\TableChecks'
.EXAMPLE
.\Get-TableMirrorCheck-SQLAnywhere17.ps1 -ServerName 'host1:2638','host2:2638' -DatabaseName 'DB1','DB2' -Credential (Get-Credential) -OutputDirectory 'C:\Temp\TableChecks'
#>
[CmdletBinding(PositionalBinding = $false, DefaultParameterSetName = 'Integrated')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ServerName,
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $true)]
    [string[]]$DatabaseName,
    [Parameter(Mandatory = $true, ParameterSetName = 'Database')]
    [System.Management.Automation.PSCredential]$Credential,
    [Parameter(ParameterSetName = 'Integrated')]
    [switch]$IntegratedSecurity,
    [string]$DriverName = 'SQL Anywhere 17',
    [string]$EngineName,
    [string]$Encryption,
    [ValidateRange(1, 2147483647)]
    [int]$ConnectionTimeoutSeconds = 15,
    [string]$CsvPath,
    [string]$OutputDirectory,
    [ValidateRange(1, 2147483647)]
    [int]$QueryTimeoutSeconds = 120,
    [ValidateRange(1, 9223372036854775807)]
    [long]$BatchRowLimit = 100000000
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

function Invoke-DatabaseQuery {
    param([string]$TargetServerName, [string]$DatabaseName, [string]$Query)

    $builder = New-Object System.Data.Odbc.OdbcConnectionStringBuilder
    $builder.Driver = $DriverName
    $builder['Host'] = $TargetServerName
    $builder['DBN'] = $DatabaseName
    if ($EngineName) { $builder['ServerName'] = $EngineName }
    if ($Encryption) { $builder['ENC'] = $Encryption }
    if ($null -ne $Credential) {
        $builder['INT'] = 'NO'
        $builder['UID'] = $Credential.UserName
        $builder['PWD'] = $Credential.GetNetworkCredential().Password
    }
    else {
        # Use the Windows identity running PowerShell; do not send UID/PWD.
        $builder['INT'] = 'YES'
    }
    $builder['CON'] = 'TableMirrorCheck'
    # Do not auto-start a database or server during an inventory.
    $builder['ASTART'] = 'NO'
    $connection = New-Object System.Data.Odbc.OdbcConnection($builder.ConnectionString)
    $connection.ConnectionTimeout = $ConnectionTimeoutSeconds
    $builder.Clear()
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $command.CommandTimeout = $QueryTimeoutSeconds
    $adapter = New-Object System.Data.Odbc.OdbcDataAdapter($command)
    $table = New-Object System.Data.DataTable
    try {
        $connection.Open()
        [void]$adapter.Fill($table)
        # Preserve the DataTable even when the query returns zero or one row.
        return ,$table
    }
    finally {
        $adapter.Dispose()
        $command.Dispose()
        $connection.Dispose()
    }
}

$query = @'
SELECT USER_NAME(t.creator) AS SchemaName,
       t.table_name AS TableName,
       dt.DateTimeColumn1,
       dt.DateTimeColumn2,
       CAST(t."count" AS BIGINT) AS "RowCount",
       CASE WHEN EXISTS (SELECT 1 FROM SYS.SYSIDX i
                         WHERE i.table_id = t.table_id AND i.index_category = 1)
            THEN 1 ELSE 0 END AS HasPrimaryKey,
       (SELECT LIST(TRIM(pc.column_name), ', ' ORDER BY ic.sequence)
        FROM SYS.SYSIDX i
        JOIN SYS.SYSIDXCOL ic
          ON ic.table_id = i.table_id AND ic.index_id = i.index_id
        JOIN SYS.SYSTABCOL pc
          ON pc.table_id = ic.table_id AND pc.column_id = ic.column_id
        WHERE i.table_id = t.table_id AND i.index_category = 1) AS PK,
       CASE WHEN EXISTS (SELECT 1 FROM SYS.SYSTABCOL c
                         JOIN SYS.SYSDOMAIN d ON d.domain_id = c.domain_id
                         LEFT JOIN SYS.SYSUSERTYPE u ON u.type_id = c.user_type
                         WHERE c.table_id = t.table_id
                           AND (LOWER(TRIM(d.domain_name)) = 'xml'
                             OR LOWER(TRIM(u.type_name)) = 'xml'))
            THEN 'Yes' ELSE 'No' END AS HasXmlColumn,
       CASE WHEN EXISTS (SELECT 1 FROM SYS.SYSTABCOL c
                         JOIN SYS.SYSDOMAIN d ON d.domain_id = c.domain_id
                         WHERE c.table_id = t.table_id
                           AND LOWER(TRIM(d.domain_name)) IN
                               ('binary', 'varbinary', 'long binary', 'image'))
            THEN 'Yes' ELSE 'No' END AS HasVarbinaryColumn
FROM SYS.SYSTAB t
LEFT JOIN (
    SELECT ranked.table_id,
           MAX(CASE WHEN ranked.ColumnNumber = 1 THEN ranked.column_name END) AS DateTimeColumn1,
           MAX(CASE WHEN ranked.ColumnNumber = 2 THEN ranked.column_name END) AS DateTimeColumn2
    FROM (
        SELECT c.table_id, TRIM(c.column_name) AS column_name,
               ROW_NUMBER() OVER (PARTITION BY c.table_id ORDER BY c.column_id) AS ColumnNumber
        FROM SYS.SYSTABCOL c
        JOIN SYS.SYSDOMAIN d ON d.domain_id = c.domain_id
        WHERE LOWER(TRIM(d.domain_name)) IN
              ('timestamp', 'datetime', 'smalldatetime', 'timestamp with time zone', 'datetimeoffset')
    ) ranked
    WHERE ranked.ColumnNumber <= 2
    GROUP BY ranked.table_id
) dt ON dt.table_id = t.table_id
WHERE t.table_type = 1 AND t.server_type = 1
  AND LOWER(USER_NAME(t.creator)) NOT IN ('sys', 'dbo', 'rs_systabgroup', 'sa_diagnostic')
ORDER BY SchemaName, TableName;
'@

if ($CsvPath -and $OutputDirectory) {
    throw 'Use either -CsvPath or -OutputDirectory, not both.'
}
$servers = @($ServerName | ForEach-Object {
    if ([string]::IsNullOrWhiteSpace($_)) { throw 'Server names cannot be blank.' }
    $_.Trim()
} | Select-Object -Unique)
if (-not $CsvPath -and -not $OutputDirectory) {
    $OutputDirectory = (Get-Location).Path
}
$targetDatabases = @($DatabaseName | ForEach-Object {
    if ([string]::IsNullOrWhiteSpace($_)) { throw 'Database names cannot be blank.' }
    $_.Trim()
} | Select-Object -Unique)
$usedCsvPaths = @{}
foreach ($currentServerName in $servers) {
    foreach ($targetDatabaseName in $targetDatabases) {
        # Replace Windows-invalid filename characters in both names.
        $safeServerName = $currentServerName -replace '[<>:"/\\|?*\x00-\x1F]', '_'
        $safeDatabaseName = $targetDatabaseName -replace '[<>:"/\\|?*\x00-\x1F]', '_'
        if ($CsvPath) {
            $parent = Split-Path -Path $CsvPath -Parent
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
            $leaf = "$stem - $safeServerName - $safeDatabaseName.csv"
            if ($parent) { $serverCsvPath = Join-Path $parent $leaf }
            else { $serverCsvPath = $leaf }
        }
        else {
            $serverCsvPath = Join-Path $OutputDirectory "TableMirrorCheck - $safeServerName - $safeDatabaseName.csv"
        }
        $serverCsvPath = [System.IO.Path]::GetFullPath($serverCsvPath)
        # Disambiguate server/database names which sanitize to the same filename.
        $candidatePath = $serverCsvPath
        $suffix = 1
        while ($usedCsvPaths.ContainsKey($serverCsvPath)) {
            $suffix++
            $parent = [System.IO.Path]::GetDirectoryName($candidatePath)
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($candidatePath)
            $serverCsvPath = Join-Path $parent "$stem ($suffix).csv"
        }
        $usedCsvPaths[$serverCsvPath] = $true
        try {
            $parent = Split-Path -Path $serverCsvPath -Parent
            if (-not (Test-Path -LiteralPath $parent)) {
                [void](New-Item -ItemType Directory -Path $parent -Force)
            }
            $results = New-Object 'System.Collections.Generic.List[object]'
            Write-Verbose "Checking [$currentServerName].[$targetDatabaseName]"
            $data = Invoke-DatabaseQuery -TargetServerName $currentServerName -DatabaseName $targetDatabaseName -Query $query
            foreach ($row in $data.Rows) {
                $results.Add([pscustomobject][ordered]@{
                    ServerName         = $currentServerName
                    DatabaseName       = $targetDatabaseName
                    SchemaName         = [string]$row.SchemaName
                    TableName          = [string]$row.TableName
                    DateTimeColumn1    = [string]$row.DateTimeColumn1
                    DateTimeColumn2    = [string]$row.DateTimeColumn2
                    RowCount           = [long]$row.RowCount
                    BatchNumber        = 0
                    BatchTotalRows     = [long]0
                    BatchTableCount    = 0
                    PK                 = $(if ([int]$row.HasPrimaryKey -eq 1) { ([string]$row.PK).Trim().Replace('[', '').Replace(']', '') } else { 'No' })
                    HasXmlColumn       = [string]$row.HasXmlColumn
                    HasVarbinaryColumn = [string]$row.HasVarbinaryColumn
                    IsMemoryOptimized  = 'N/A'
                    IsPartitioned      = 'N/A'
                    Is_Mirror          = $(if ([int]$row.HasPrimaryKey -eq 1 -and $row.HasXmlColumn -eq 'No' -and $row.HasVarbinaryColumn -eq 'No') { 'Yes' } else { 'No' })
                })
            }

            # First-fit decreasing: fill earlier batches where a whole table fits.
            # Isolated batches are closed even if the table has exactly the row limit.
            foreach ($databaseGroup in ($results | Group-Object DatabaseName)) {
                $nextBatchNumber = 0
                $batches = New-Object 'System.Collections.Generic.List[object]'
                $sortedTables = @($databaseGroup.Group | Sort-Object -Property @{Expression = 'Is_Mirror'; Descending = $true}, @{Expression = 'RowCount'; Descending = $true}, SchemaName, TableName)
                foreach ($tableResult in $sortedTables) {
                    $selectedBatch = $null
                    if ($tableResult.RowCount -lt $BatchRowLimit) {
                        foreach ($batch in $batches) {
                            if ($batch.IsMirror -eq $tableResult.Is_Mirror -and -not $batch.Closed -and $tableResult.RowCount -le ($BatchRowLimit - $batch.TotalRows)) {
                                $selectedBatch = $batch
                                break
                            }
                        }
                    }
                    if ($null -eq $selectedBatch) {
                        $nextBatchNumber++
                        $selectedBatch = [pscustomobject]@{
                            Number = $nextBatchNumber
                            IsMirror = $tableResult.Is_Mirror
                            TotalRows = [long]0
                            Closed = ($tableResult.RowCount -ge $BatchRowLimit)
                            Tables = (New-Object 'System.Collections.Generic.List[object]')
                        }
                        $batches.Add($selectedBatch)
                    }
                    $selectedBatch.Tables.Add($tableResult)
                    $selectedBatch.TotalRows += $tableResult.RowCount
                }
                foreach ($batch in $batches) {
                    foreach ($tableResult in $batch.Tables) {
                        $tableResult.BatchNumber = $batch.Number
                        $tableResult.BatchTotalRows = $batch.TotalRows
                        $tableResult.BatchTableCount = $batch.Tables.Count
                    }
                }
            }
            $results = @($results | Sort-Object DatabaseName, BatchNumber, SchemaName, TableName)

            if ($serverCsvPath) {
                if ($results.Count -gt 0) {
                    $results | Export-Csv -LiteralPath $serverCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                }
                else {
                    # Include column headers even when no tables were returned.
                    '"ServerName","DatabaseName","SchemaName","TableName","DateTimeColumn1","DateTimeColumn2","RowCount","BatchNumber","BatchTotalRows","BatchTableCount","PK","HasXmlColumn","HasVarbinaryColumn","IsMemoryOptimized","IsPartitioned","Is_Mirror"' |
                        Set-Content -LiteralPath $serverCsvPath -Encoding UTF8 -ErrorAction Stop
                }
            }
            Write-Host "[$currentServerName].[$targetDatabaseName] CSV: $serverCsvPath"
            Write-Verbose "Tables returned: $($results.Count)"
            $results
        }
        catch {
            Write-Warning "Database [$currentServerName].[$targetDatabaseName] failed: $($_.Exception.Message). No complete CSV was produced for this database; any previous CSV at that path may be stale. Continuing."
        }
    }
}
