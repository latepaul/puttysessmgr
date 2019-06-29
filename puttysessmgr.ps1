# puttysessmgr.ps1 - yet another program to manage putty sessions
#
# 27-Jun-2019 - Paul Mason
#               Created initial version

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function choose_cat {

    $cat_list_form.Topmost = $true

    $result = $cat_list_form.ShowDialog()

    $x="other"
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $x = $cat_listBox.SelectedItem
        if ($x -eq "new...")
        {
        $x = "other"
        }
    }

    return $x 
}
function prompt {
    param([string] $title,
        [string] $message,
        [string] $default)
    

    $prompt_label.Text = $message
    $Form2.Text = $title
    $txtbox.Text = $default 
    [void] $Form2.ShowDialog() 
    return $txtbox.Text
}
function rightclick {
    param([System.Windows.Forms.TreeNode]$node)

    if ($node.Tag -ne 'item') {
        return
    }
    $nodetext = $node.Name

    $cat = $node_cats[$nodetext]
    $orig_cat=$cat 
    if ($cat -eq 'none') {
        $cat = ''
    }

    write-host "right-click on: $nodetext (cat=$cat)"

    $chosen_cat = choose_cat
    if ($chosen_cat -eq "other") {
        $prompt_msg = 'Enter category for ' + $nodetext + ':'
        $text = prompt 'Category' $prompt_msg $cat
        Write-Debug "entered text = [$text]"
        $chosen_cat = $text 
    }


    if ($chosen_cat -eq '' -or $chosen_cat -eq $orig_cat )
    {
    write-host "Not changing cat"
    $has_changed=0
    }
    else
    {
    Write-host "Change cat for $nodetext to $chosen_cat"
    $has_changed=1
    $node_cats[$nodetext] = $chosen_cat
    }
    return $has_changed
}

function launch {
    param([System.Windows.Forms.TreeNode]$node)

    $tag = $node.Tag

    if ($tag -eq 'item') {
        $sess_name = $node.Text
        if ($sess_name -eq 'Open Putty') { 
            $cmd = '& ' + $putty_exe
        } 
        else {
            $cmd = '& ' + $putty_exe + ' -load ' + $sess_name
        }
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

    if ($name -eq 'Open Putty') {
        $newnode.Tag = 'special'
    }
    else {    
        $newnode.Tag = 'item'
    }
    
    if ($category -eq 'none') {
        Write-Host "$name has no category, adding to root node"
        [void] $rootnode.Add($newnode) 
    }
    else {
        write-host "Searching for node called $category"
        $rootnode | ForEach-Object { '...Name: {0} Text: {1}' -f $_.Name, $_.Text }
        $addto_node = $rootnode | Where-Object { $_.Name -eq $category }

        if ($addto_node) {
            write-host "found node with $category - $addto_node.Name"
            [void] $addto_node.Nodes.Add($newnode) 
        }
        else {
            Write-Host 'adding new category node $category'
            $newcatnode = New-Object System.Windows.Forms.TreeNode
            $newcatnode.Text = $category
            $newcatnode.Name = $category
            $newcatnode.Tag = 'category'
            [void] $rootnode.Add($newcatnode)
            $addto_node = $newcatnode
            [void] $addto_node.Nodes.Add($newnode) 
       
        }
    }
}

$puttypath = 'C:\Program Files\PuTTY'
$putty_exe = $puttypath + '\putty.exe'
$putty_exe 

$Form = New-Object system.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(400, 600)
$Form.Text = 'Putty Session Launcher'
$Form.Startposition = 'CenterScreen'

$putty_icon = [System.Drawing.Icon]::ExtractAssociatedIcon($putty_exe)
if ($putty_icon) {
    $Form.Icon = $putty_icon
}
else {
    Write-Host "No icon"
}

$tree = New-Object System.Windows.Forms.TreeView

$tree.Size = '370,500'
$tree.Location = '5,5'
$tree.Text = 'Putty Sessions'
$tree.Font = '"Consolas",10'

$categories = @()
$node_cats = @{ }

$sessions = Get-ChildItem -Path Registry::HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions | `
    ForEach-Object { split-path -leaf $_.Name } | `
    ForEach-Object { [uri]::UnescapeDataString($_) }

$i = 0 
foreach ($sess in $sessions) {
    if ($i -lt 5) {
        $cat = 'Favourites'
    } 
    else {
        $cat = 'none'
    }

    $i += 1

    $new_cat = $categories | Where-Object { $_ -eq $cat }
    if (-not $new_cat) {
        $categories += $cat
    }
    $node_cats.add($sess, $cat) 
}

add_node $tree.Nodes 'none' 'Open Putty'
$i = 0 
foreach ($sess in $sessions) {
    write-host "`nAdding node - Session=[$sess] cat=[$($node_cats[$sess])]"
    
    add_node $tree.Nodes  $($node_cats[$sess]) $sess
    $i += 1
}

$tree.add_MouseDoubleClick( {
        launch($this.SelectedNode)
    })

$tree.add_NodeMouseClick( {
        $whichnode = $this.SelectedNode
        if ($_.Button -eq 'Right') {
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
$Form2.Size = New-Object System.Drawing.Size(320, 130)
$Form2.Text = 'question'
$Form2.FormBorderStyle = 'FixedDialog'
$Form2.StartPosition = 'CenterScreen'

$txtbox = New-Object System.Windows.Forms.TextBox
$txtbox.size = '270,100'
$txtbox.Location = '15,30'
$prompt_label = New-Object System.Windows.Forms.Label
$prompt_label.Text = 'Enter'
$prompt_label.Location = '5,5'
$prompt_label.Size = '280,20'
$Form2.Controls.Add($prompt_label) 
$Form2.Controls.Add($txtbox)

$close_btn2 = New-Object System.Windows.Forms.Button
$close_btn2.location = '140,55'
$close_btn2.size = '40,25'
$close_btn2.Text = 'OK'
$close_btn2.Font = '"Arial",8'
$Form2.Controls.Add($close_btn2)
$close_btn2.Add_Click( { $Form2.Close() })

$cat_list_form = New-Object system.Windows.Forms.Form
$cat_list_form.Size = New-Object System.Drawing.Size(320, 200)
$cat_list_form.Text = 'Choose a category'
$cat_list_form.FormBorderStyle = 'FixedDialog'
$cat_list_form.StartPosition = 'CenterParent'

$cl_OKbtn = New-Object System.Windows.Forms.Button
$cl_OKbtn.Location = New-Object System.Drawing.Point(75,120)
$cl_OKbtn.Size = New-Object System.Drawing.Size(75,23)
$cl_OKbtn.Text = 'OK'
$cl_OKbtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
$cat_list_form.AcceptButton = $cl_OKbtn
$cat_list_form.Controls.Add($cl_OKbtn)

$cl_CnclBtn = New-Object System.Windows.Forms.Button
$cl_CnclBtn.Location = New-Object System.Drawing.Point(150,120)
$cl_CnclBtn.Size = New-Object System.Drawing.Size(75,23)
$cl_CnclBtn.Text = 'Cancel'
$cl_CnclBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cat_list_form.CancelButton = $cl_CnclBtn
$cat_list_form.Controls.Add($cl_CnclBtn)

$cl_label = New-Object System.Windows.Forms.Label
$cl_label.Location = New-Object System.Drawing.Point(10,20)
$cl_label.Size = New-Object System.Drawing.Size(280,20)
$cl_label.Text = 'Please select a category:'
$cat_list_form.Controls.Add($cl_label)

$cat_listBox = New-Object System.Windows.Forms.listBox
$cat_listBox.Location = New-Object System.Drawing.Point(10,40)
$cat_listBox.Size = New-Object System.Drawing.Size(260,20)
$cat_listBox.Height = 80

foreach ($cat in $categories) {
    [void] $cat_listBox.Items.Add($cat)
    write-host "Adding $cat to cat list dialog"
}
[void] $cat_listBox.Items.Add('new...')

$cat_list_form.Controls.Add($cat_listBox)

$close_btn.Add_Click( { $Form.Close()
        $form2.Close()
        $cat_list_form.close() })

$Form.ShowDialog()
$form2.close()
$cat_list_form.Close()
