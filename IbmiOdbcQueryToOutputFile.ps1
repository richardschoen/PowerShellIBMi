##-------------------------------------------------------------------------------
## Desc: Connect to IBM i and Query Data using ODBC Driver
##
## $ExitType - ENVIRONMENT=Use Environment.Exit. Only use when calling from DOS 
## or ShellExec. Otherwise it will kill the calling application. 
## EXIT=Use standard PowerShell exit with appropriate exit code. 
## RETURN=Do a standard PowerShell return. This might return a 0 for success and 1 for failure rather than actual return code.
## 
## Parms:
## -ibmihost=IBM i host name/ip address
## -ibmiuser=IBM i user id
## -ibmipass=IBM i password
## -sql=SQL select query. Ex:select * from qiws.qcustcdt
## -delim=Output record delimiter.  Ex: "," or "|"
## -outputheadings=Output column headings. Default=$True
## -outputtofile=Output to a file. Default=$True
## -outputfilename=Output file name. Ex:"c:\temp\qcustcdt.txt",
## -exittype=Type of exit from the script. Default="EXIT"
##  "ENVIRONMENT"=Use Environment.Exit. Only use when calling from DOS 
##                or ShellExec. Otherwise it will kill the calling application. 
##  "EXIT"=Use standard PowerShell exit with appropriate exit code. 
##  "RETURN"=Do a standard PowerShell return. This might return a 0 for success and 1 for failure rather than actual return code.
##
## Returns:
## ExitCode or 0=success, 99=errors
##-------------------------------------------------------------------------------

## Defined parameters
param(
[string]$ibmihost="",
[string]$ibmiuser="",
[string]$ibmipass="",
[string]$sql="",
[string]$delim=",",
[string]$outputheadings=$True,
[string]$outputtofile=$True,
[string]$outputfilename="c:\temp\qcustcdt.txt",
[string]$exittype="EXIT"
)

## Load shared IBM i PowerShell functions.
## Should be in same directory as this script
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path 
Import-Module (Join-Path $ScriptDir "IbmiOdbcFunctions.psm1")

## Init initial work variables
$exitcode=0
$columncount=0
$lasterror=""
$lastcpferror=""
$lastcpfmsgid=""
$lastdbcolcount=0
$lastdbrowcount=0
$conn=$Null

##-------------------------------------------------------------------------------
## Sample custom PowerShell function to write SQL output to command line.
## Illustrates creating a custotm PowerShell function.
##-------------------------------------------------------------------------------
function outputquerysql 
{
Param ([string] $strSql)
Write-Output "SQL Query: $strSql"
}

##-------------------------------------------------------------------------------
## Let's try to do our work now and nicely handle errors
## This script should always end normally with an appropriate exit code
##-------------------------------------------------------------------------------
try {

	## Output SQL to command line
    outputquerysql $sql

    ## Display parameters passsed
    Write-Output "Querying Database Table to Output File $outputfilename"
    Write-Output "SQL: $sql"
    Write-Output "ExitType: $exittype"

	## Close connection if already open beore running
    $rtnclose = dbCloseConn -dbconnection $conn

	## Connect to IBMi over CA/400 ODBC driver
    $conn = dbOpenConnIbmi -ibmihost $ibmihost `
			-ibmiuser $ibmiuser `
	        -ibmipass $ibmipass

	## Run query to output file
    $rtnquery=dbExecQueryToOutputFile -dbconnection $conn -sqlquery $sql -outputfilename $outputfilename -delim $delim -outputtofile $outputtofile -outputheadings $outputheadings
    if ($rtnquery -eq $False) {
       throw [System.Exception]::new($lasterror) 
    }
   
    ## Close IBM i database connection
    $rtnclose = dbCloseConn -dbconnection $conn

    ##-------------------------------------------------------------------------------
    ## Handle final output and returns
    ##-------------------------------------------------------------------------------
    Write-Output "ExitCode: $exitcode"
    Write-Output ("Message: $lasterror")
    Write-Output "StackTrace:"
    exit $exitcode

    ## Causes caller to exit. Only when you need to send a DOS return code to caller
    if ($exittype.ToUpper() -eq "ENVIRONMENT") { 
       [Environment]::Exit($exitcode)
    } ## Causes standard powershell exit 
    elseif ($exittype.ToUpper() -eq "EXIT") { 
       exit $exitcode
    }   
    elseif ($exittype.ToUpper() -eq "RETURN") { 
       return
    }   
    else { 
	   ## default if nothing else
	   return
    }

}
##-------------------------------------------------------------------------------
## Catch and handle any errors and return useful info via console
##-------------------------------------------------------------------------------
catch [System.Exception] {
	$exitcode=99
    ## Attempt to close open connection
	Write-Output "ExitCode: $exitcode"
	Write-Output ("Message:" + $_.Exception.Message + " Line:" + $_.InvocationInfo.ScriptLineNumber.ToString() + " Char:" + $_.InvocationInfo.OffsetInLine.ToString())
    Write-Output "StackTrace: $_.Exception.StackTrace" 

	## Causes caller to exit. Only when you need to send a DOS return code to caller
	if ($exittype.ToUpper() -eq "ENVIRONMENT") { 
	   [Environment]::Exit($exitcode)
    } ## Causes standard powershell exit 
	elseif ($exittype.ToUpper() -eq "EXIT") { 
	   exit $exitcode
    }
	elseif ($exittype.ToUpper() -eq "RETURN") { 
	   return
    }
    else { 
	   ## default if nothing else
	   return
    }

} finally {
    ## Always insure database closure at the end of our script even on exceptions
    $rtnclose = dbCloseConn -dbconnection $conn
}
