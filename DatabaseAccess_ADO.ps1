$adOpenStatic = 3
$adLockOptimistic = 3
$objConnection = New-Object -comobject ADODB.Connection
$objRecordset = New-Object -comobject ADODB.Recordset
$objConnection.Open("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Users\jonat\OneDrive\Databases\Northwind.accdb")
$objRecordset.Open("Select * from companies", $objConnection,$adOpenStatic,$adLockOptimistic)

$objRecordset.MoveFirst()
do {$objRecordset.Fields.Item("CompanyName").Value; $objRecordset.MoveNext()} until 
    ($objRecordset.EOF -eq $True)
$objRecordset.Close()
$objConnection.Close()


################################################
# https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/schemaenum
$objRecordset.openschema('adSchemaTables')