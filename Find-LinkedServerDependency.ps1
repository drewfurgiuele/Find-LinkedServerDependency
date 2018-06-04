<#
.SYNOPSIS
This function will scan various objects for any use of four part identifiers that indicate they use a linked server reference

.DESCRIPTION
Did you ever want to know how many objects in your database reference a linked server? It can be hard! You could do a find and replace in SSMS with tools like (the excellent) SQL Search from RedGate, or maybe
using pattern matching with sp_helptext or other metadata views. But what your linked server shares, say, a schema name? You'd get lots of false positives. Instead, why not be thorough? This function will parse each individual
object's DDL code and look for a consecutive quoted or non-quoted identifier string and flag the object as having at least one reference. Note that, currently, if the same linked server is referenced twice in the same object, this function will
return two results for the same object.

Returns an object that contains the object containing the linked server reference, the linked server name, database, schema, and object being referenced, and the definition of the referencing object (as a hidden property)

.PARAMETER ServerInstance
The SQL Server instance's host name to run this function against. Can be either just a host name for a default instance, or a named instance (sqlserver/instancename). This is a required parameter.

.PARAMETER DatabaseName
The database that contains the objects you want to parse out. This is a required parameter.

.PARAMETER PathToScriptDomLibrary
The explicit path to the Microsoft.SqlServer.TransactSql.ScriptDom.dll. If you run this function on a machine that has SQL Server installed, this parameter should not be needed. However, if this library cannot be found in any default location
(or you're running on a machine/VM that doesn't have SQL Server installed) you can manually supply a path to this library (see NOTES below)./

.NOTES
This function requires:
    1. The SQLSERVER PowerShell Module ()
    2. An accessable Microsoft.SqlServer.TransactSql.ScriptDom.dll file (Most commonly available via SQL Server Install, but can be copied to another location and referenced via the -PathToScriptDomLibrary parameter)

.LINK
http://port1433.com/2018/04/27/finding-linked-server-references-using-powershell/

.EXAMPLE
$references = ./Find-LinkedServerDependency.ps1 -ServerName sqlserver -DatabaseName databaseName

#>

[cmdletbinding()]
param(
    [Parameter(Position=0,Mandatory=$true)] [string] $ServerInstance,
    [Parameter(Position=1,Mandatory=$true)] [string] $DatabaseName,
    [Parameter(Position=2,Mandatory=$false)] [string] $PathToScriptDomLibrary = $null,
    [Parameter(Position=3,Mandatory=$false)] [string] $UseQuotedIdentifier = $true
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

    $DictionaryObject = New-Object 'System.Collections.Generic.Dictionary[string, string]'

    Class ReferenceExpression
    {
        [string] $TextToReplace
        [bool] $IsQuotedIdentifier = $false
    }

    Class LinkedServerReference
    {
        [string] $ReferencingObjectSchema
        [string] $ReferencingObjectName
        [string] $ReferencingObjectType
        [string] $LinkedServerName
        [string] $Database
        [string] $Schema
        [string] $Object
        [int] $TotalReferences = 0
        hidden [string] $Definition
        hidden [ReferenceExpression[]] $RefExpression
        hidden [System.Collections.Generic.Dictionary[string,string]] $Dictionary
        
        [String] ReplaceLinkedServer([String] $FromServer, [String] $ToServer)
        {
            $ReplacedCode = $this.Definition
            $Pattern = New-Object System.Text.RegularExpressions.Regex("CREATE", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $ReplacedCode = $pattern.replace($ReplacedCode, "ALTER", 1)                  

            if ($this.LinkedServerName -eq $FromServer) {
                ForEach ($e in $this.RefExpression) {
                    if ($e.IsQuotedIdentifier) {
                        $PatternToLookFor = New-Object System.Text.RegularExpressions.Regex(("\[" + $this.LinkedServerName + "\]."), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        $ReplacedServerName = "[" + $ToServer + "]."
                    } else {
                        $PatternToLookFor = New-Object System.Text.RegularExpressions.Regex((" " + $this.LinkedServerName + "."), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        $ReplacedServerName =  $ToServer + "."
                    }
                    $RefExpressionFixed = $PatternToLookFor.Replace($e.TextToReplace,$ReplacedServerName)
                    $ReplacedCode = $ReplacedCode.Replace($e.TextToReplace,$RefExpressionFixed)
               }
            } 
            $ReplacedCode = $ReplacedCode.Replace("SET ANSI_NULLS ON","")
            $ReplacedCode = $ReplacedCode.Replace("SET QUOTED_IDENTIFIER ON","")
            $ReplacedCode = $ReplacedCode + "`r`nGO`r`n"
            $ReplacedCode = $ReplacedCode.TrimStart()
            
            return $ReplacedCode
        }

        [String] RemoveLinkedServer([String] $FromServer)
        {
            $ReplacedCode = $this.Definition
            $Pattern = New-Object System.Text.RegularExpressions.Regex("CREATE", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $ReplacedCode = $pattern.replace($ReplacedCode, "ALTER", 1)        
            if ($this.LinkedServerName -eq $FromServer) {
                ForEach ($e in $this.RefExpression) {
                    if ($e.IsQuotedIdentifier) {
                        $PatternToLookFor = New-Object System.Text.RegularExpressions.Regex(("\[" + $this.LinkedServerName + "\]."), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    } else {
                        $PatternToLookFor = New-Object System.Text.RegularExpressions.Regex((" " + $this.LinkedServerName + "."), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    }
                    $RefExpressionFixed = $PatternToLookFor.Replace($e.TextToReplace,"")
                    $ReplacedCode = $ReplacedCode.Replace($e.TextToReplace,$RefExpressionFixed)
               }
            } 
            $ReplacedCode = $ReplacedCode.Replace("SET ANSI_NULLS ON","")
            $ReplacedCode = $ReplacedCode.Replace("SET QUOTED_IDENTIFIER ON","")
            $ReplacedCode = $ReplacedCode + "`r`nGO`r`n"
            $ReplacedCode = $ReplacedCode.TrimStart()

            return $ReplacedCode
        }


        [int] GetTotalReferences()
        {
            return $this.RefExpression.Count
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

    if ($GenerateChangeScript) {
        $InstanceObject = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName" | Where-Object {$_.DisplayName -eq "$Instancename"}
        $DropScripterObject = New-Object Microsoft.SqlServer.Management.Smo.Scripter($InstanceObject)
        $DropScripterObject.Options.ScriptDrops = $True
    }

    if (Test-Path -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Tables") {
        Write-Verbose "Getting table triggers..."
        $Tables = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Tables"
        $TableTriggers = $Tables.Triggers
        $DictionaryObject.Add('Trigger','TRIGGER')
    }
    if (Test-Path -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Triggers") {
        Write-Verbose "Getting database triggers..."
        $DBTriggers = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Triggers"
    }
    if (Test-Path -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Synonyms") {
        Write-Verbose "Getting synonyms..."
        $Synonyms = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Synonyms"
        $DictionaryObject.Add('Synonym','SYNONYM')
    }
    if (Test-Path -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Views") {
        Write-Verbose "Getting views..."
        $Views = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\Views"
        $DictionaryObject.Add('View','VIEW')
    }
    if (Test-Path -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\UserDefinedFunctions") {
        Write-Verbose "Getting table functions..."
        $Functions = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\UserDefinedFunctions"
        $DictionaryObject.Add('UserDefinedFunction','FUNCTION')
    }
    if (Test-Path -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\StoredProcedures") {
        Write-Verbose "Getting stored procedures..."
        $Procs = Get-ChildItem -Path "SQLSERVER:\SQL\$ServerName\$InstanceName\Databases\$DatabaseName\StoredProcedures"
        $DictionaryObject.Add('StoredProcedure','PROCEDURE')
    }

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

        if ($o.GetType().Name -eq "Synonym") {
            $ObjectDDL = ($o.Script() -join "`r`n")
        } else {
            $ObjectDDL = $o.TextBody
        }

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
        
        $ReferencesForThisObject = @()

        ForEach ($t in $Tokens) {

            if ($t.TokenType -ne "dot") {
                if ($t.TokenType -eq "QuotedIdentifier" -or $t.TokenType -eq "Identifier") {
                    $IdentifierCount++
                } else {
                    $IdentifierCount = 0
                }
                if ($IdentifierCount -eq 4) {
                    $IdentifierCount = 0                    
                    if ($ReferencesForThisObject.LinkedServerName -contains $Tokens[$Iteration - 6].Text.Replace("[","").Replace("]","")) {
                        Write-Verbose "Already a linked server name for this reference, skipping..."
                        $ExistingReference = $ReferencesForThisObject | Where-Object {$_.LinkedServerName -eq $Tokens[$Iteration - 6].Text.Replace("[","").Replace("]","")}

                        $Expression = New-Object ReferenceExpression 
                        $Expression.TextToReplace = $Tokens[$Iteration - 6].Text + "." + $Tokens[$Iteration - 4].Text + "." + $Tokens[$Iteration - 2].Text + "." + $Tokens[$Iteration].Text
                        if ($Tokens[$Iteration - 6].TokenType -eq "QuotedIdentifier") {
                            $Expression.IsQuotedIdentifier = $true
                        }
                        $ExistingReference.RefExpression += $Expression
                        $ExistingReference.TotalReferences = $ExistingReference.GetTotalReferences()
                    } else {
                        $Reference = New-Object LinkedServerReference

                        $Reference.ReferencingObjectSchema = $o.Schema
                        $Reference.ReferencingObjectName = $o.Name
                        $Reference.ReferencingObjectType = $o.GetType().Name
                        $Reference.LinkedServerName = $Tokens[$Iteration - 6].Text.Replace("[","").Replace("]","")
                        $Reference.Database = $Tokens[$Iteration - 4].Text.Replace("[","").Replace("]","")
                        $Reference.Schema = $Tokens[$Iteration - 2].Text.Replace("[","").Replace("]","")
                        $Reference.Object = $Tokens[$Iteration].Text.Replace("[","").Replace("]","")
                        $Reference.Definition = ($o.Script() -join "`r`n")
                        $Reference.Dictionary = $DictionaryObject

                        $Expression = New-Object ReferenceExpression 
                        $Expression.TextToReplace = $Tokens[$Iteration - 6].Text + "." + $Tokens[$Iteration - 4].Text + "." + $Tokens[$Iteration - 2].Text + "." + $Tokens[$Iteration].Text
                        if ($Tokens[$Iteration - 6].TokenType -eq "QuotedIdentifier") {
                            $Expression.IsQuotedIdentifier = $true
                        }
                        $Reference.RefExpression = $Expression
                        $Reference.TotalReferences = $Reference.GetTotalReferences()

                        $ReferencesForThisObject += $Reference
                    }
                }

            }
            $Iteration++
        }
        $ReferencesForThisObject

        $streamReader.Close()
    }
}