<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER Files
.PARAMETER PathToScriptDomLibrary
.PARAMETER UseQuotedIdentifier
.NOTES
.LINK
.EXAMPLE
.EXAMPLE
#>

[cmdletbinding()]
param(
    [Parameter(Mandatory=$true)] [string] $ServerInstance,
    [Parameter(Mandatory=$true)] [string] $DatabaseName,
    [Parameter(Mandatory=$false)] [string] $PathToScriptDomLibrary = $null,
    [Parameter(Mandatory=$false)] [string] $UseQuotedIdentifier = $true
)

begin {
    Write-Verbose "Testing for required modules being loaded..."
    if ((Get-Module sqlps) -eq $null -and (Get-Module sqlserver) -eq $null)
    {
        Throw "SQLPS/SQLSERVER module not loaded! Unable to continue"
    }

    $LibraryLoaded = $false
    $ObjectCreated = $false
    $LibraryVersions = @(13,12,11)

    if ($PathToScriptDomLibrary -ne "") {
        try {
            Add-Type -Path $PathToScriptDomLibrary -ErrorAction SilentlyContinue
            Write-Verbose "Loaded library from path $PathToScriptDomLibrary"
        } catch {
            throw "Couldn't load the required ScriptDom library from the path specified!"
        }
    } else {
        ForEach ($v in $LibraryVersions)
        {
            if (!$LibraryLoaded) {
                try {
                    Add-Type -AssemblyName "Microsoft.SqlServer.TransactSql.ScriptDom,Version=$v.0.0.0,Culture=neutral,PublicKeyToken=89845dcd8080cc91"  -ErrorAction SilentlyContinue
                    Write-Verbose "Loaded version $v.0.0.0 of the ScriptDom library."
                    $LibraryLoaded = $true                
                } catch {
                    Write-Verbose "Couldn't load version $v.0.0.0 of the ScriptDom library."
                }
            }
        }
    }

    ForEach ($v in $LibraryVersions)
    {
        if (!$ObjectCreated) {
            try {
                $ParserNameSpace = "Microsoft.SqlServer.TransactSql.ScriptDom.TSql" + $v + "0Parser"
                $Parser = New-Object $ParserNameSpace($UseQuotedIdentifier)
                $ObjectCreated = $true
                Write-Verbose "Created parser object for version $v..."
            } catch {
                Write-Verbose "Couldn't load version $v.0.0.0 of the ScriptDom library."
            }
        }
    }
}

process {
    $ObjectScripts = @()

    if ($ServerInstance -like "*\*") {
        $SplitInstance = $ServerInstance.split("\")
        $ServerName = $SplitInstance[0]
        $InstanceName = $SplitInstance[1]
    } else {
        $ServerName = $ServerInstance
        $InstanceName = "DEFAULT"
    }

    Write-Verbose "Getting table triggers..."
    $Tables = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Tables"
    $TableTriggers = $Tables.Triggers
    Write-Verbose "Getting database triggers..."
    $DBTriggers = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Triggers"
    Write-Verbose "Getting synonyms..."
    $Synonyms = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Synonyms"
    Write-Verbose "Getting views..."
    $Views = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Views"
    Write-Verbose "Getting table functions..."
    $Functions = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\UserDefinedFunctions"
    Write-Verbose "Getting stored procedures..."
    $Procs = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\StoredProcedures"

    $ObjectScripts += $Views
    $ObjectScripts += $TableTriggers
    $ObjectScripts += $DBTriggers
    $ObjectScripts += $Synonyms
    $ObjectScripts += $Functions
    $ObjectScripts += $Procs
    
    ForEach ($o in $ObjectScripts) {
        $CurrentObject = $o.Schema + "." + $o.Name
        $VerboseMessage = "Parsing $currentObject"
        Write-Verbose $VerboseMessage

        $ObjectDDL = $o.Script() -join "`r`nGO`r`n"

        $memoryStream = New-Object System.IO.MemoryStream
        $streamWriter = New-Object System.IO.StreamWriter($memoryStream)
        $streamWriter.Write($ObjectDDL)
        $streamWriter.Flush()
        $memoryStream.Position = 0
        
        $streamReader = New-Object System.IO.StreamReader($memoryStream)

        $Errors = $null
        $Fragments = $Parser.Parse($streamReader, [ref] $Errors)

        $Tokens = $Fragments.ScriptTokenStream
    
        $IdentifierCount = 0
        $Iteration = 0
        
        ForEach ($t in $Tokens) {

            if ($t.TokenType -ne "dot") {
                if ($t.TokenType -eq "QuotedIdentifier" -or $t.TokenType -eq "Identifier") {
                    $IdentifierCount++
                } else {
                    $IdentifierCount = 0
                }
                if ($IdentifierCount -eq 4) {
                    $IdentifierCount = 0
                    $LinkedServerReference = [PSCustomObject] @{
                        ReferencingObjectSchema = $o.Schema
                        ReferencingObjectName = $o.Name
                        ReferencingObjectType = $o.GetType().Name
                        LinkedServerName = $Tokens[$Iteration - 6].Text.Replace("[","").Replace("]","")
                        Database = $Tokens[$Iteration - 4].Text.Replace("[","").Replace("]","")
                        Schema = $Tokens[$Iteration - 2].Text.Replace("[","").Replace("]","")
                        Object = $Tokens[$Iteration].Text.Replace("[","").Replace("]","")
                    }
                    $LinkedServerReference
                }
            }
            $Iteration++
        }
        $streamReader.Close()
    }
}