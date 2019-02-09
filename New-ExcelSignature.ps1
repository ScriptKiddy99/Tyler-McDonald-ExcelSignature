# Creates a new Excel Com Object.
$Excel = new-object -comobject excel.application 
 
# Set window view size if Excel opens
$Excel.WindowState = "xlMaximized"
# Sets the window vsilbility property to true so we can see the excel instance.
$Excel.visible = $true

# Stops all those annoying alerts and pops ups from getting in our way.
$Excel.DisplayAlerts = $false

# Create New Workbook.
$workbook = $Excel.Workbooks.Add() 

# Gets the first Woorksheet in the currently open workbook.
$WorkSheet_Sheet1 = $workbook.Worksheets.item(1)

# stores a range of cells (In this case we select just G2@
$range = $WorkSheet_Sheet1.Range("G2")

# Selects our chosen range
$range.Select()

# Adds a new singature to the work book. with the top left of the singaure being at G2
$workbook.Signatures.AddSignatureLine("{00000000-0000-0000-0000-000000000000}")

