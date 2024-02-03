##-------------------------------------------------------------------------------
## Desc: Connect to IBM i and Query Data or run commands using 
##       IBM i Access/400 ODBC Driver.
## Note when developing and testing, you can remove an already loaded module 
## you just made changes to using the following PowerShell commandlet call
## from the Windows PowerShell ISE Editor console:
## Remove-Module  -Name "IbmiOdbcFunctions"
##-------------------------------------------------------------------------------

##-------------------------------------------------------------------------------
## Open ODBC, OLEDB or SQL Server connection - IBMi specific host and user info
##
## Returns a connection object or Null if no connection opened.
##-------------------------------------------------------------------------------
function dbOpenConnIbmi {
 Param([string] $ibmihost,
       [string] $ibmiuser, 
	   [string] $ibmipass, 
	   [string] $conntype="ODBC",[string]  
	   $connstring = "Driver={IBM i Access ODBC Driver};System=@@SYSTEM;Uid=@@USER;Pwd=@@PASS;"
	   ) 
 
    $global:lasterror=""
    $global:lastcpferror=""
    $global:lastcpfmsgid=""

    try {


        ## Set the selected connection type (ODBC, OLEDB or SQL Server Client-SQLCLIENT)
        if ($conntype="ODBC") {
          $dbconnection = New-Object System.Data.ODBC.ODBCConnection
        } elseif ($conntype="OLEDB") {
          $dbconnection = New-Object System.Data.OleDb.OleDbConnection
        } elseif ($conntype="SQLCLIENT") {
          $dbconnection = New-Object System.Data.SQLClient.SQLCOnnection
        } 
  
        ## Set IBM i ODBC connection string
        $connstring=$connstring.replace("@@SYSTEM",$ibmihost)
        $connstring=$connstring.replace("@@USER",$ibmiuser)
        $connstring=$connstring.replace("@@PASS",$ibmipass)

        ## Set the connection string
        $dbconnection.ConnectionString = $connstring
        ## Open the connection 
        $dbconnection.Open()

        $global:lasterror="Connection opened successfully to $ibmihost."
        return $dbconnection

    } catch [System.Exception] { 
	   $global:lasterror = "dbOpenConn error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $Null
    }

}

##-------------------------------------------------------------------------------
## Open ODBC, OLEDB or SQL Server connection
##
## Returns a connection object or Null if no connection opened.
##-------------------------------------------------------------------------------
function dbOpenConn {
 Param([string] $connstring, 
       [string] $conntype="ODBC"
	  ) 

    $global:lasterror=""
    $global:lastcpferror=""
    $global:lastcpfmsgid=""

    try {


        ## Set the selected connection type (ODBC, OLEDB or SQL Server Client-SQLCLIENT)
        if ($conntype="ODBC") {
          $dbconnection = New-Object System.Data.ODBC.ODBCConnection
        } elseif ($conntype="OLEDB") {
          $dbconnection = New-Object System.Data.OleDb.OleDbConnection
        } elseif ($conntype="SQLCLIENT") {
          $dbconnection = New-Object System.Data.SQLClient.SQLCOnnection
        } 
  
        ## Set the connection string
        $dbconnection.ConnectionString = $connstring
        ## Open the connection 
        $dbconnection.Open()

        $global:lasterror="Connection opened successfully."
        return $dbconnection

    } catch [System.Exception] { 
	   $global:lasterror = "dbOpenConn error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $Null
    }

}

##-------------------------------------------------------------------------------
## Close ODBC, OLEDB or SQL Server connection if open
##
## Returns $True=Successful close. $False=Unsuccessful close or connection not open
##-------------------------------------------------------------------------------
function dbCloseConn {
 Param($dbconnection) 

    try {
       
       $global:lasterror=""
       $global:lastcpferror=""
       $global:lastcpfmsgid=""

       ## Close the connection 
       if ($dbconnection -ne $Null) {
          $dbconnection.Close()
          $global:lasterror="Connection closed successfully."
          return $True
       } else {
          $global:lasterror="Connection was not open."
          return $False
       }

    } catch [System.Exception] { 
	   $global:lasterror = "dbCloseConn error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $False
    }

}

##-------------------------------------------------------------------------------
## Execute SQL query to DataTable using already open connection
## Note: Must return DataTable with "return,tblName", otherwise data rows get 
## returned instead of the actual DataTable. See link listed below for explanation
## https://stackoverflow.com/questions/1918190/strange-behavior-in-powershell-function-returning-dataset-datatable
##
## Returns DataTable object or $Null if query failed.
##-------------------------------------------------------------------------------
function dbExecQueryToDataTable {
 Param($dbconnection, 
       [string] $sqlquery
	  ) 
 
    try {
       
        $global:lasterror=""
        $global:lastcpferror=""
        $global:lastcpfmsgid=""

        ## Create SQL query comment
        $cmd = $dbconnection.CreateCommand()
        ## Set the query SQL
        $cmd.CommandText = $Sqlquery
        $result = $cmd.ExecuteReader()
        ## Create DataTable and load resulting data
        $tabletemp = new-object "System.Data.DataTable"
        $tabletemp.Load($result)
        ## Get rows andcolumn count
        $dbcolcount=$tabletemp.columns.count
        $dbrowcount=$tabletemp.rows.count

        return ,$tabletemp

    } catch [System.Exception] { 
	   $global:lasterror = "dbExecQueryToDataTable error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $Null
    }

}

##-------------------------------------------------------------------------------
## Execute SQL query to Output File using already open database connection
## Note: Must return DataTable with return,tblName, otherwise data rows get 
## returned instead of the actual DataTable. See link listed below for explanation
## https://stackoverflow.com/questions/1918190/strange-behavior-in-powershell-function-returning-dataset-datatable
##
## Returns $True=Success, $False=Error
##-------------------------------------------------------------------------------
function dbExecQueryToOutputFile {
 Param($dbconnection,[string] $sqlquery,[string] $outputfilename,[string] $delim="|",$outputtofile,$outputheadings) 

    try {

        $exportedrecords=0

        $global:lasterror=""
        ## Create SQL query comment
        $cmd = $dbconnection.CreateCommand()
        ## Set the query SQL
        $cmd.CommandText = $Sqlquery
        $result = $cmd.ExecuteReader()
        ## Create DataTable and load resulting data
        $tabletemp = new-object "System.Data.DataTable"
        $tabletemp.Load($result)
        ## Get rows andcolumn count
        $dbcolcount=$tabletemp.columns.count
        $dbrowcount=$tabletemp.rows.count
        $global:dbcolcount=$tabletemp.columns.count
        $global:dbrowcount=$tabletemp.rows.count

        $record=""
        $curfield=1
        $colcount=$tabletemp.columns.count
        $recorddelim=$delim
        
        ##-------------------------------------------------------------------------------
        ## Iterate all column name fields for current row and output each field name for heading row
        ##-------------------------------------------------------------------------------
        foreach ($col1 in $tabletemp.columns) {

           ## Clear delimiter before last column         
           if ($curfield -EQ $colcount) { 
              $recorddelim="" } 
           else { $curfield++ }
     
           ## Write field data to current record buffer
           $record += Write-Output "$($col1.ColumnName)$recorddelim"
        }
        ## Output the final heading record buffer for current row if enabled
        if ($outputheadings -EQ $True) {
           ## Output data to console
           Write-Output $record
           ## Output to file if selected
           if ($outputtofile -EQ $True) {
             Out-File -FilePath $outputfilename -InputObject $record -Encoding ASCII
           }
        }

        ## reset current field to 1 
        $record=""
        $curfield=1
        $colcount=$tabletemp.columns.count
        $recorddelim=$delim
        ## Iterate all table rows and output
        foreach ($row1 in $tabletemp.Rows)
        {

          $exportedrecords++
 
          ##-------------------------------------------------------------------------------
          ## Iterate all fields for current row and output each field value with a delimiter
          ##-------------------------------------------------------------------------------
          foreach ($col1 in $tabletemp.columns) {
    
              ## Clear delimiter before last column         
              if ($curfield -EQ $colcount) { 
                 $recorddelim="" } 
              else { $curfield++ }
        
              ## Write field data to current record buffer
              $record += Write-Output "$($row1[$($col1.ColumnName)].ToString().TrimEnd())$recorddelim"
          }
  
          ## Output the final record buffer for current row to console
          Write-Output $record
  
          ## Output to file if selected
          if ($outputtofile -EQ $True) {
            Out-File -Append -FilePath $outputfilename -InputObject $record -Encoding ASCII
          }
  
          ## Reset field buffers and oclumn counters
          $record=""
          $curfield=1
          $recorddelim=$delim
        }

 	    $global:lasterror = "execQueryToOutputFile completed. $exportedrecords records output to $outputfilename"
        return $True
}
##-------------------------------------------------------------------------------
## Catch and handle any errors and return useful info via console
##-------------------------------------------------------------------------------
catch [System.Exception] {
	   $global:lasterror = "dbExecQueryToOutputFile error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $False
}

}

##-------------------------------------------------------------------------------
## Execute IBM i CL command using already open database connection
##
## Returns $True=Success, $False=Error
##-------------------------------------------------------------------------------
function dbExecIbmiClCommand {
 Param($dbconnection, 
      [string] $clcommand
	  ) 

    try {

        $global:lasterror=""
        $global:lastcpferror=""
        $global:lastcpfmsgid=""

        ## Create command object
        $cmd = $dbconnection.CreateCommand()
        ## Set SQL to call commmand
        $cmd.CommandText = "call qsys2.qcmdexc('$clcommand')"
        $result = $cmd.ExecuteNonQuery()
        $cmd.Dispose()

 	    $global:lasterror = "dbExecIbmiClCommand completed."
        return $True
}
##-------------------------------------------------------------------------------
## Catch and handle any CL command errors and return useful info via last error
##-------------------------------------------------------------------------------
catch [System.Exception] {
	   $global:lasterror = "dbIbmiClCommand error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
	   $global:lastcpferror = $_.Exception.Message 
       $locationcpf = $global:lastcpferror.IndexOf("CPF")
       if ($locationcpf -gt 0) { 
  	      $global:lastcpferror = $global:lastcpferror.Substring($locationcpf)
  	      $global:lastcpfmsgid = $global:lastcpferror.Substring(0,7)
       }
       return $False
}

}


##-------------------------------------------------------------------------------
## Execute SQL query to DataTable over new ODBC connection. Closes connection after running.
## Note: Must return DataTable with "return,tblName", otherwise data rows get 
## returned instead of the actual DataTable. See link listed below for explanation
## https://stackoverflow.com/questions/1918190/strange-behavior-in-powershell-function-returning-dataset-datatable
##
## Returns DataTable object or $Null if query failed.
##-------------------------------------------------------------------------------
function execQueryToDataTable {
 Param([string] $connstring, 
       [string] $sqlquery,
	   [string] $conntype="ODBC"
	  ) 

    try {
       
        $global:lasterror=""
        $global:lastcpferror=""
        $global:lastcpfmsgid=""


        ## Set the selected connection type (ODBC, OLEDB or SQL Server Client-SQLCLIENT)
        if ($conntype="ODBC") {
          $dbconnection = New-Object System.Data.ODBC.ODBCConnection
        } elseif ($conntype="OLEDB") {
          $dbconnection = New-Object System.Data.OleDb.OleDbConnection
        } elseif ($conntype="SQLCLIENT") {
          $dbconnection = New-Object System.Data.SQLClient.SQLCOnnection
        } 
  
        ## Set the connection string
        $dbconnection.ConnectionString = $connstring
        ## Open the connection 
        $dbconnection.Open()
        ## Create SQL query comment
        $cmd = $dbconnection.CreateCommand()
        ## Set the query SQL
        $cmd.CommandText = $Sqlquery
        $result = $cmd.ExecuteReader()
        ## Create DataTable and load resulting data
        $tabletemp = new-object "System.Data.DataTable"
        $tabletemp.Load($result)
        ## Close the database connections. We're done.
        $dbconnection.Close()
        ## Get rows andcolumn count
        $dbcolcount=$tabletemp.columns.count
        $dbrowcount=$tabletemp.rows.count

        return ,$tabletemp

    } catch [System.Exception] { 
	   $global:lasterror = "execQueryToDataTable error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $Null
    }

}

##-------------------------------------------------------------------------------
## Execute SQL query to Output File over new ODBC connection. Closes connection after running.
##
## Returns $True=Success, $False=Error
##-------------------------------------------------------------------------------
function execQueryToOutputFile {
 Param([string] $connstring,
       [string] $sqlquery,
	   [string] $conntype="ODBC",
	   [string] $outputfilename,
	   [string] $delim="|",
	   $outputtofile,
	   $outputheadings
	  ) 

    try {

        $exportedrecords=0

        $global:lasterror=""

        ## Set the selected connection type (ODBC, OLEDB or SQL Server Client-SQLCLIENT)
        if ($conntype="ODBC") {
          $dbconnection = New-Object System.Data.ODBC.ODBCConnection
        } elseif ($conntype="OLEDB") {
          $dbconnection = New-Object System.Data.OleDb.OleDbConnection
        } elseif ($conntype="SQLCLIENT") {
          $dbconnection = New-Object System.Data.SQLClient.SQLCOnnection
        } 
  
        ## Set the connection string
        $dbconnection.ConnectionString = $connstring
        ## Open the connection 
        $dbconnection.Open()
        ## Create SQL query comment
        $cmd = $dbconnection.CreateCommand()
        ## Set the query SQL
        $cmd.CommandText = $Sqlquery
        $result = $cmd.ExecuteReader()
        ## Create DataTable and load resulting data
        $tabletemp = new-object "System.Data.DataTable"
        $tabletemp.Load($result)
        ## Close the database connections. We're done.
        $dbconnection.Close()
        ## Get rows andcolumn count
        $dbcolcount=$tabletemp.columns.count
        $dbrowcount=$tabletemp.rows.count
        $global:dbcolcount=$tabletemp.columns.count
        $global:dbrowcount=$tabletemp.rows.count

        $record=""
        $curfield=1
        $colcount=$tabletemp.columns.count
        $recorddelim=$delim
        
        ##-------------------------------------------------------------------------------
        ## Iterate all column name fields for current row and output each field name for heading row
        ##-------------------------------------------------------------------------------
        foreach ($col1 in $tabletemp.columns) {

           ## Clear delimiter before last column         
           if ($curfield -EQ $colcount) { 
              $recorddelim="" } 
           else { $curfield++ }
     
           ## Write field data to current record buffer
           $record += Write-Output "$($col1.ColumnName)$recorddelim"
        }
        ## Output the final heading record buffer for current row if enabled
        if ($outputheadings -EQ $True) {
           ## Output data to console
           Write-Output $record
           ## Output to file if selected
           if ($outputtofile -EQ $True) {
             Out-File -FilePath $outputfilename -InputObject $record -Encoding ASCII
           }
        }

        ## reset current field to 1 
        $record=""
        $curfield=1
        $colcount=$tabletemp.columns.count
        $recorddelim=$delim
        ## Iterate all table rows and output
        foreach ($row1 in $tabletemp.Rows)
        {

          $exportedrecords++
 
          ##-------------------------------------------------------------------------------
          ## Iterate all fields for current row and output each field value with a delimiter
          ##-------------------------------------------------------------------------------
          foreach ($col1 in $tabletemp.columns) {
    
              ## Clear delimiter before last column         
              if ($curfield -EQ $colcount) { 
                 $recorddelim="" } 
              else { $curfield++ }
        
              ## Write field data to current record buffer
              $record += Write-Output "$($row1[$($col1.ColumnName)].ToString().TrimEnd())$recorddelim"
          }
  
          ## Output the final record buffer for current row to console
          Write-Output $record
  
          ## Output to file if selected
          if ($outputtofile -EQ $True) {
            Out-File -Append -FilePath $outputfilename -InputObject $record -Encoding ASCII
          }
  
          ## Reset field buffers and oclumn counters
          $record=""
          $curfield=1
          $recorddelim=$delim
        }

 	    $global:lasterror = "execQueryToOutputFile completed. $exportedrecords records output to $outputfilename"
        return $True
}
##-------------------------------------------------------------------------------
## Catch and handle any errors and return useful info via console
##-------------------------------------------------------------------------------
catch [System.Exception] {
	   $global:lasterror = "execQueryToOutputFile error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $False
}

}

##-------------------------------------------------------------------------------
## Execute IBM i CL command over new ODBC connection. Closes connection after running.
##
## Returns $True=Success, $False=Error
##-------------------------------------------------------------------------------
function execIbmiClCommand {
 Param([string] $connstring,
       [string] $clcommand,
	   [string] $conntype="ODBC"
	  ) 

    try {

        $exportedrecords=0

        $global:lasterror=""

        ## Set the selected connection type (ODBC, OLEDB or SQL Server Client-SQLCLIENT)
        if ($conntype="ODBC") {
          $dbconnection = New-Object System.Data.ODBC.ODBCConnection
        } elseif ($conntype="OLEDB") {
          $dbconnection = New-Object System.Data.OleDb.OleDbConnection
        } elseif ($conntype="SQLCLIENT") {
          $dbconnection = New-Object System.Data.SQLClient.SQLCOnnection
        } 
  
        ## Set the connection string
        $dbconnection.ConnectionString = $connstring
        ## Open the connection 
        $dbconnection.Open()
        ## Create SQL query comment
        $cmd = $dbconnection.CreateCommand()
        ## Set the query SQL
        $cmd.CommandText = "call qsys2.qcmdexc('$clcommand')"
        $result = $cmd.ExecuteNonQuery()
        ## Close the database connections. We're done.
        $dbconnection.Close()

 	    $global:lasterror = "execIbmiClCommand completed. $exportedrecords records output to $outputfilename"
        return $True
}
##-------------------------------------------------------------------------------
## Catch and handle any errors and return useful info via console
##-------------------------------------------------------------------------------
catch [System.Exception] {
	   $global:lasterror = "execIbmiClCommand error: " + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString()
       return $False
}

}

##-------------------------------------------------------------------------------
## Get date time helper function
## Returns formatted date and time value
##-------------------------------------------------------------------------------
function GetDateTime() {
   Param([string] $datfmt="MM/dd/yyyy",  
         [string] $timfmt="HH:mm:ss"
		)
   $date = Get-Date -Format $datfmt
   $time = get-date -Format $timfmt
   $datetime = $date + " " + $time
   return $datetime
}

##-------------------------------------------------------------------------------
## Run SQL action query and return result
## Returns Action result or -2 on errors.
##-------------------------------------------------------------------------------
function execNonQuery 
{
Param([string] $strsql,
      [object] $cmd, 
	  [object] $conn
	 )
  try {
     ## Set command text 
     $cmd.CommandText = $strsql
     ## Run non query SQL
     $rtn=$command1.ExecuteNonQuery()
     return $rtn  
  } catch [System.Exception] {
	 return -2
  }
}


