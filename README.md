# IBM i PowerShell Samples
PowerShell for IBM i provides sample IBM i PowerShell scripts for interacting with IBM i.

## Files
```IbmiOdbcFunctions.psm1```   
Shared functions for connecting to IBM i via ODBC. Ideally this file should be located in the same directory as any PowerShell scripts.   

```IbmiOdbcQueryToOutputFile.ps1```    
Run SQL ODBC query to select and export records to delimited PC file.   

Example command line to run IbmiOdbcQueryToOutputFile.ps1:   
```
pwsh IbmiOdbcQueryToOutputFile.ps1 -ibmihost "1.1.1.1" -ibmiuser "USER1" -ibmipass "PASS1" -sql "select * from qiws.qcustcdt" -outputfile "c:\temp\qcustcdt.txt"
```
