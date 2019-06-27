# puttysessmgr.ps1 - yet another program to manage putty sessions
#
# 27-Jun-2019 - Paul Mason
#               Created initial version

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object system.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(400,600)
$tree = New-Object System.Windows.Forms.TreeView

$tree.Size = '380,500'
$tree.Location ='5,5'
$tree.Name = "Putty Sessions"
$sessions =  Get-ChildItem -Path Registry::HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions  | % {split-path -leaf $_.Name}

write-host "List of sessions:`n"
foreach ($sess in $sessions) 
{
    $newnode=$tree.Nodes.Add("$sess","$sess") |Out-Null
}

function rightclick ($node) {
$nodetext=$node.Text
write-host "right-click on: $nodetext"
}

$tree.add_MouseDoubleClick({
$nodetext=$this.SelectedNode.Text
write-host "double-click on: $nodetext"
})

$tree.add_NodeMouseClick({
$whichnode=$this.SelectedNode
if ($_.Button -eq 'Right')
{
  rightclick($whichnode)
}

})
$Form.controls.add($tree)
$Form.ShowDialog()