$OleDbConn = New-Object "System.Data.OleDb.OleDbConnection"
$OleDbCmd = New-Object "System.Data.OleDb.OleDbCommand"
$OleDbAdapter = New-Object "System.Data.OleDb.OleDbDataAdapter"
$DataTable = New-Object "System.Data.DataTable"

$OleDbConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\jonat\OneDrive\Databases\Northwind.accdb;"
$OleDbCmd.Connection = $OleDbConn
$OleDbCmd.CommandText = "select * from companies"


$OleDbConn.Open()

$OleDbCmd.ExecuteScalar()

$schemaTables = $OleDbConn.GetSchema('tables')|? table_type -eq 'table'
$schemaTables |ogv

$schemaIndexes = $OleDbConn.GetSchema('indexes') |? table_name -in $schemaTables.Table_Name # -eq 'table'
$schemaIndexes |ogv
$Indexes = $schemaIndexes|group table_name -AsHashTable

$schemaColumns = $OleDbConn.GetSchema('columns') |? table_name -in $schemaTables.Table_Name # -eq 'table'
$schemaColumns |ogv
$Columns = $schemaColumns|group table_name -AsHashTable


$OleDbConn

$o = foreach($table in $schemaTables ){
    # get rowcount
    $OleDbCmd.CommandText = ("Select count (*) from [{0}]" -f $table.TABLE_NAME)
    
    $RowCount = $OleDbCmd.ExecuteScalar()
    [PSCustomObject]@{
        Table_Name = $table.TABLE_NAME
        ColumnCount = ($Columns["$($table.TABLE_NAME)"] ).count
        RowCount = $RowCount
        IndexCount = ($Indexes["$($table.TABLE_NAME)"] ).count
    }
}
$o |ft


$OleDbConn.Close()