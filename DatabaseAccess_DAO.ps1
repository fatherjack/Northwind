# Load the DAO.DBEngine COM object
$dbe = New-Object -ComObject DAO.DBEngine.120

# Specify the database path
$DatabasePath = "C:\Users\jonat\OneDrive\Databases\Northwind.accdb"

# Open the database
$db = $dbe.OpenDatabase($DatabasePath)

# Get the TableDefs collection
$tableDefs = $db.TableDefs

# Loop through each table
$relationships = foreach ($tableDef in $tableDefs ) {
    # Loop through each relation in the table
    foreach ($relation in $db.Relations) {
        [PSCustomObject]@{
            FKName = $relation.Name
            Table = $relation.Table
            ForeignTable = $relation.ForeignTable
            Fields = $relation.Fields -join '; '
        }
#         if ($relation.Table -eq $tableDef.Name) {
#             # Loop through each field in the relation
#             foreach ($field in $relation.Fields) {
#                 # Check if the field is a foreign key
#                 if ($field.Attributes -band 0x80000000) {
# #                    if ($field.Attributes -eq 4096) {
#                     # Print the foreign key details
#                     Write-Output "Table: $($tableDef.Name), Field: $($field.Name), Foreign Table: $($relation.ForeignTable), Foreign Field: $($field.ForeignName)"
#                 }
#             }
#         }
    }
}

$relationships|group table -NoElement
# Close the database
$db.Close()


####################### 
$tableDefs|ogv

$relation |ogv