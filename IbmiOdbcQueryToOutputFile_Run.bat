cd %~dp0
rem powershell IbmiOdbcQueryToOutputFile.ps1 -ibmihost SYSi1 -ibmiuser RICHARD -ibmipass IJS2032 -sql 'select * from qiws.qcustcdt' -outputfile 'c:\rjstemp\qcustcdttxt'
pwsh IbmiOdbcQueryToOutputFile.ps1 -ibmihost "sysi1" -ibmiuser "RICHARD" -ibmipass "IJS2032" -sql "select * ddfrom qiws.qcustcdt" -outputfile "c:\rjstemp\qcustcdt.txt"

echo "Error level:" %errorlevel%

pause
