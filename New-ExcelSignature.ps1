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

# Adds a new singature to the work book.
$workbook.Signatures.AddSignatureLine("{00000000-0000-0000-0000-000000000000}")
