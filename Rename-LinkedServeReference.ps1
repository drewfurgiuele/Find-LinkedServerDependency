
        if ($GenerateChangeScript) {
            $FileName = ($ScriptSavePath.Trim("\") + "\" + $CurrentObject + (Get-Date).ToFileTimeUtc() + ".sql")
            $DropScripterObject.Script($o) | Out-File $FileName
            $CreateCode = $o.Script() -join "`r`nGO`r`n"
            if ($NewLinkedServerName) {
                $CreateCode = $CreateCode.Replace(("[" + $LinkedServerReference.LinkedServerName + "]."),("[" + $NewLinkedServerName + "]."))
                $CreateCode = $CreateCode.Replace(($LinkedServerReference.LinkedServerName + "."),($NewLinkedServerName + "."))
            } else {
                $CreateCode = $CreateCode.Replace(("[" + $LinkedServerReference.LinkedServerName + "]."),"")
                $CreateCode = $CreateCode.Replace(($LinkedServerReference.LinkedServerName + "."),"")
            }
            Write-Verbose "Writing change file to $ScriptSavePath..."
            $CreateCode | Out-File $FileName -Append
        }