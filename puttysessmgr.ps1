# puttysessmgr.ps1 - yet another program to manage putty sessions
#
# 27-Jun-2019 - Paul Mason
#               Created initial version

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()



function prompt{
    param([string] $title,
          [string] $message,
          [string] $default)
    
    $answer=$default
    $prompt_label.Text = $message
    $Form2.Text = $title
    $txtbox.Text = $default 
    $Form2.Location = $Form.Location
    $Form2.ShowDialog() |Out-Null
    return $txtbox.Text
}
function rightclick {
param([System.Windows.Forms.TreeNode]$node)
$nodetext=$node.Name
$cat=$node_cats[$nodetext]
if ($cat -eq "none")
{
   $cat = ""
}

write-host "right-click on: $nodetext (cat=$cat)"
$prompt_msg = "Enter category for "+$nodetext+":"
$text = prompt "Category" $prompt_msg $cat
Write-Host "entered text = [$text]"

}

function launch {
    param([System.Windows.Forms.TreeNode]$node)

$tag = $node.Tag

   if ($tag -eq "item")
   {
        $sess_name=$node.Text
        $cmd="& "+ $putty_exe + " -load " + $sess_name
        write-host "Launching: $cmd"
        & $putty_exe -load $sess_name
   }
}

function add_node {
param([System.Windows.Forms.TreeNodeCollection]$rootnode,
      [string] $category,
      [string] $name)

    $newnode = New-Object System.Windows.Forms.TreeNode
    $newnode.Name = $name
    $newnode.Text = $name
    $newnode.Tag = "item"

    if ($category -eq "none")
    {
        Write-Host "$name has no category, adding to root node"
        $rootnode.Add($newnode) |out-null        
    }
    else
    {
        write-host "Searching for node called $category"
        $rootnode | ForEach-Object {"...Name: {0} Text: {1}" -f $_.Name, $_.Text }
        $addto_node = $rootnode | Where-Object {$_.Name -eq $category}

        if ($addto_node)
        {
            write-host "found node with $category - $addto_node.Name"
            $addto_node.Nodes.Add($newnode) |Out-Null
        }
        else
        {
            Write-Host "adding new category node $category"
            $newcatnode = New-Object System.Windows.Forms.TreeNode
            $newcatnode.Text=$category
            $newcatnode.Name=$category
            $newcatnode.Tag = "category"
            $rootnode.Add($newcatnode)|Out-Null
            $addto_node = $newcatnode
            $addto_node.Nodes.Add($newnode)|Out-Null
       
        }
    }
}

$puttypath = "C:\Program Files\PuTTY"
$putty_exe = $puttypath + "\putty.exe"
$putty_exe 

$Form = New-Object system.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(400,600)
$Form.Text = "Putty Session Launcher"
$putty_icon = [System.Drawing.Icon]::ExtractAssociatedIcon($putty_exe)
if ($putty_icon)
{
$Form.Icon = $putty_icon
}
else
{
Write-Host "No icon"
}

$tree = New-Object System.Windows.Forms.TreeView

$tree.Size = '370,500'
$tree.Location ='5,5'
$tree.Text = "Putty Sessions"
$tree.Font = '"Consolas",10'

$categories=@()
$node_cats=@{}

$sessions =  Get-ChildItem -Path Registry::HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions  | ForEach-Object {split-path -leaf $_.Name}

$i=0 
foreach ($sess in $sessions) 
{
    if ($i -lt 2)
    {
        $cat = "Red"
    } 
    elseif ($i -lt 4)
    {
        $cat = "Blue"
    }
    elseif ($i -lt 7)
    {
        $cat = "Green"
    }
    else
    {
        $cat = "none"
    }

    $i +=1

    $new_cat = $categories | Where-Object {$_ -eq $cat }
    if (-not $new_cat)
    {
    $categories += $cat
    }
    $node_cats.add($sess,$cat) 
}

$i=0 
foreach ($sess in $sessions) 
{
    write-host "`nAdding node - Session=[$sess] cat=[$($node_cats[$sess])]"
    
    add_node $tree.Nodes  $($node_cats[$sess]) $sess
    $i += 1
}

$tree.add_MouseDoubleClick({
launch($this.SelectedNode)
})

$tree.add_NodeMouseClick({
$whichnode=$this.SelectedNode
if ($_.Button -eq 'Right')
{
  rightclick($whichnode)
}

})
$Form.controls.add($tree)

$close_btn = New-Object System.Windows.Forms.Button
$close_btn.location = '160,520'
$close_btn.size = '80,25'
$close_btn.Text = 'Close'
$close_btn.Font = '"Arial",10'
$Form.Controls.Add($close_btn)

$Form2 = New-Object system.Windows.Forms.Form
$Form2.Size = New-Object System.Drawing.Size(320,130)
$Form2.Text = "question"
$Form2.FormBorderStyle='FixedDialog'

$txtbox = New-Object System.Windows.Forms.TextBox
$txtbox.size = '270,100'
$txtbox.Location = '15,30'
$prompt_label = New-Object System.Windows.Forms.Label
$prompt_label.Text = "Enter"
$prompt_label.location = '5,5'
$prompt_label.Size = '280,20'
$Form2.Controls.Add($prompt_label) 
$Form2.Controls.Add($txtbox)

$close_btn2 = New-Object System.Windows.Forms.Button
$close_btn2.location = '140,55'
$close_btn2.size = '40,25'
$close_btn2.Text = 'OK'
$close_btn2.Font = '"Arial",8'
$Form2.Controls.Add($close_btn2)
$close_btn2.Add_Click({$Form2.Close()})

$close_btn.Add_Click({$Form.Close()
$form2.Close()})

$Form.ShowDialog()
$form2.close()
