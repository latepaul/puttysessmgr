# puttysessmgr.ps1 - yet another program to manage putty sessions
#
# 27-Jun-2019 - Paul Mason
#               Created initial version
# 01-Jul-2019 - Paul Mason #13 
#               Implement keyboard shortcut for launch

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# cleanup - stuff we need to do when exiting
# this is a function because we can be called from the close
# button or by quitting the form
function cleanup {
    $prompt_form.close()
    $cat_list_form.Close()
    $txtbox.remove_KeyDown($txtbox_KeyDown)
}

# save_config - save current config to file

function save_config {
    if ($global:config_filename -eq "") {
        $def_filename_path = $Env:USERPROFILE
        if ($def_filename_path -eq "") {
            $def_filename_path = $Env:TEMP
        }
        $def_filename = $def_filename_path + '\puttysessmgr.ini'
        $global:config_filename = prompt "Config file" "Enter the filename of the config" $def_filename        
    }

    write-debug "`n`nSAVE_CONFIG - writing new file"
    write-debug "writing [putty_exe]"
    "[putty_exe]" | Out-File $global:config_filename
    write-debug "saving p_e=$global:putty_exe"
    $global:putty_exe | Out-File $global:config_filename -Append
    write-debug "writing [sessions]"
    "[sessions]" | Out-File $global:config_filename -Append
    Write-Debug "`n`nPAUL - header done`n`n"
    $global:node_cats.GetEnumerator() | ForEach-Object {
        $cat = $_.Value 
        $sess = $_.Key
        Write-Debug "save_config: cat: $cat Sess: $sess"
        "session=$sess" | Out-File $global:config_filename -Append
        "category=$cat" | Out-File $global:config_filename -Append
    }
}

function load_config {
    if ($global:config_filename -eq "") {
        $def_filename_path = $Env:USERPROFILE
        if ($def_filename_path -eq "") {
            $def_filename_path = $Env:TEMP
        }
        $def_filename = $def_filename_path + '\puttysessmgr.ini'
        $global:config_filename = prompt "Config file" "Enter the filename of the config" $def_filename        
    }

    $global:node_cats.Clear()
    $global:categories.Clear()
    $script:section = ""
    $global:putty_exe ="unset"
    Get-Content $global:config_filename | ForEach-Object {
        if ($script:section -eq "putty_exe" -and $script:section[0] -ne '[' -and $global:putty_exe -eq "unset") {
            $global:putty_exe = $_ 
        }
        elseif ($script:section -eq "sessions") {
            $type = $_.split("=")
            if ($type[0] -eq "session") {
                $sess = $type[1]
            }
            elseif ($type[0] -eq "category") {
                $cat = $type[1]
                $global:node_cats[$sess] = $cat 
                if ($global:categories -notcontains $cat) {
                    $global:categories += $cat
                }
            }
            
        }

        if ($_ -eq "[putty_exe]") {
            $script:section = "putty_exe"
        }
        elseif ($_ -eq "[sessions]") {
            $script:section = "sessions"
        }
    }

    
}
# refresh_cat_list
function refresh_cat_list {

    [void] $cat_listBox.Items.Clear()

    # add 'special' entries first 
    [void] $cat_listBox.Items.Add('new...')
    [void] $cat_listBox.Items.Add('none')
    [void] $cat_listBox.Items.Add('Favourites')

    # now add actual categories
    foreach ($cat in $global:categories) {
        if ($cat -ne "Favourites" -and $cat -ne "none" `
                -and $global:node_cats.ContainsValue($cat)) {
            [void] $cat_listBox.Items.Add($cat)
        }
        
    }

    # remove unused categories
    $remove_cats = $()
    $global:categories | ForEach-Object {
        if (-not $global:node_cats.ContainsValue($_)) {
            $remove_cats += $_
        }
    }

    $remove_cats | ForEach-Object {
        $global:categories.Remove($_)
    }

}
# rebuild_tree
function rebuild_tree {

    # remove existing item nodes
    $tree.Nodes.Clear()

    # first add a special node for "open Putty" which just opens Putty without
    # a specific session

    add_node $tree.Nodes 'none' 'Open Putty'

    # add a node for each category
    if ($global:categories -contains "Favourites") {
        Write-Debug "  Category: Favourites"
        $newcatnode = New-Object System.Windows.Forms.TreeNode
        $newcatnode.Text = "Favourites"
        $newcatnode.Name = "Favourites"
        $newcatnode.Tag = 'category'
        [void] $tree.Nodes.Add($newcatnode)
    }      

    $global:categories | Sort-Object | get-unique | ForEach-Object {
        $cat = $_ 
        write-debug "  Category: $cat"
        if ($cat -ne "Favourites" -and $cat -ne "none") {

            Write-Debug "      $cat not none/Faves"

            if ($global:node_cats.ContainsValue($cat)) {
                Write-Debug "      $cat is in node-cats"
                $existing_cat_node = $tree.Nodes.find($cat, $true)
                if (-not $existing_cat_node) {  
                    Write-Debug "         not found will create"
                    $newcatnode = New-Object System.Windows.Forms.TreeNode
                    $newcatnode.Text = $cat
                    $newcatnode.Name = $cat
                    $newcatnode.Tag = 'category'
                    [void] $tree.Nodes.Add($newcatnode)
                }
                else {
                    Write-Debug "         found will not create"
                
                }
            }
            else {
                Write-Debug "      $cat is NOT in node-cats"
            }
            Write-Debug "`n`n`n"
        }
        else {
            Write-Debug "      $cat not eligible for node creation"
        }
    }

    foreach ($sess in $global:sessions) {
        write-debug "Adding node - Session=[$sess] cat=[$($node_cats[$sess])]"

        add_node $tree.Nodes  $($global:node_cats[$sess]) $sess
    }

}

# choose_cat - choose a category
function choose_cat {

    $Form.TopMost = $false
    $cat_list_form.Topmost = $true

    $result = $cat_list_form.ShowDialog()
    $x = @{ }
    $x[0] = "cancel"
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $x[0] = "category"
        $x[1] = $cat_listBox.SelectedItem
        if ($x[1] -eq "new...") {
            $x[0] = "new"
        }
    }

    return $x 
}

# prompt - ask a question
function prompt {
    param([string] $title,
        [string] $message,
        [string] $default)
    

    $prompt_label.Text = $message
    $prompt_form.Text = $title
    $txtbox.Text = $default 
    $result = $prompt_form.ShowDialog() 
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return ""
    }
    else {
        return $txtbox.Text
    }
}

# rightclick - when we right click on a node in the main treeview
function rightclick {
    param([System.Windows.Forms.TreeNode]$node)

    $has_changed = $false 
    if ($node.Tag -ne 'item') {
        return $has_changed
    }
    $nodetext = $node.Name

    $cat = $global:node_cats[$nodetext]
    $orig_cat = $cat 
    if ($cat -eq 'none') {
        $cat = ''
    }

    write-debug "right-click on: $nodetext (cat=$cat)"

    $cc_return = choose_cat
    $cc_status = $cc_return[0]
    $chosen_cat = $cc_return[1]

    write-debug "PAUL:rightclick:choose_cat returned: $cc_status $chosen_cat"
    if ($cc_status -eq "cancel") {
        write-debug "right-click cancel returned from choose_cat"
        return $false
    }
    
    if ($cc_status -eq "new") {
        $prompt_msg = 'Enter category for ' + $nodetext + ':'
        $text = prompt 'Category' $prompt_msg $cat
        Write-Debug "entered text = [$text]"
        $chosen_cat = $text 
    }

    if ($chosen_cat -eq '' -or $chosen_cat -eq $orig_cat ) {
        write-debug "Not changing cat"
        $has_changed = $false 
    }
    else {
        write-debug "Change cat for $nodetext to $chosen_cat"
        $has_changed = $true 
        $global:node_cats[$nodetext] = $chosen_cat
      
        if ($global:categories -notcontains $chosen_cat) {
            $global:categories += $chosen_cat
            refresh_cat_list   
        }                   
    }
    return $has_changed
}

# launch - launch putty for a session
function launch {
    param([System.Windows.Forms.TreeNode]$node,
        [int] $from)

    $tag = $node.Tag
    $name = $node.Name

    Write-Debug "launch: node=[$name] tag=[$tag] from=[$from]"

    # if this is a session launch it in putty
    if ($tag -eq 'item') {
        $sess_name = $node.Text
        write-debug "Launching: $global:putty_exe -load $sess_name"
        & $global:putty_exe -load $sess_name
       
    }

    # if this is a category then expand/collapse it
    # but only if it's a key press - treeview has expand on double-click
    # built in already
    if ($tag -eq 'category' -and $from -eq $global:from_key) {
        if ($tree.SelectedNode.IsExpanded -eq $true) {
            $tree.SelectedNode.Collapse()
        }
        else {
            $tree.SelectedNode.Expand()    
        }
        
    }

    # if this is "Open Putty" then just launch putty (with no session)
    if ($tag -eq 'openputty') {
        write-debug "Launching: $global:putty_exe"
        & $global:putty_exe
    }
}

# add a new node to the tree
function add_node {
    param([System.Windows.Forms.TreeNodeCollection]$rootnode,
        [string] $category,
        [string] $name)

    $newnode = New-Object System.Windows.Forms.TreeNode
    $newnode.Name = $name
    $newnode.Text = $name

    if ($name -eq 'Open Putty') {
        $newnode.Tag = 'openputty'
    }
    else {    
        $newnode.Tag = 'item'
    }
    
    if ($category -eq 'none') {
        write-debug "$name has no category, adding to root node"
        [void] $rootnode.Add($newnode) 
    }
    else {
        write-debug "Searching for node called $category"
        $addto_node = $rootnode | Where-Object { $_.Name -eq $category }

        if ($addto_node) {
            write-debug "found node with $category"
            [void] $addto_node.Nodes.Add($newnode) 
        }
        else {
            write-debug 'adding new category node $category'
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

# main script

# DebugPreference determines whether Write-Debug statements get 
# output or not
#$DebugPreference = "SilentlyContinue"
$DebugPreference = "Continue"

# path to putty
$puttypath = 'C:\Program Files\PuTTY'
$global:putty_exe = $puttypath + '\putty.exe'
$global:putty_exe 

# main form is $Form
$Form = New-Object system.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(400, 600)
$Form.Text = 'Putty Session Launcher'
$Form.Startposition = 'CenterScreen'

# use the icon from the putty exe
$putty_icon = [System.Drawing.Icon]::ExtractAssociatedIcon($global:putty_exe)
if ($putty_icon) {
    $Form.Icon = $putty_icon
}
else {
    write-debug "No icon"
}

# $tree is the treeview
$tree = New-Object System.Windows.Forms.TreeView

$tree.Size = '370,500'
$tree.Location = '5,5'
$tree.Text = 'Putty Sessions'
$tree.Font = '"Consolas",10'

# categories is the category list
[System.Collections.ArrayList]$global:categories = @()
# node_cats maps nodes to their categories
$global:node_cats = @{ }

$global:config_filename = ""

# a couple of globals for which action launched launch
$global:from_key = 1
$global:from_mouse = 2

# List of sessions from the Putty registry key
$global:sessions = Get-ChildItem -Path Registry::HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions | `
    ForEach-Object { split-path -leaf $_.Name } | `
    ForEach-Object { [uri]::UnescapeDataString($_) }


# build categories and node_cats from the sessions
$i = 0 
foreach ($sess in $global:sessions) {
    if ($i -lt 5) {
        $cat = 'Favourites'
    } 
    else {
        $cat = 'none'
    }

    $i += 1

    if ($global:categories -notcontains $cat ) {
        $global:categories += $cat
    }
    $global:node_cats.add($sess, $cat) 
}


# start adding nodes to the treeview

rebuild_tree 

# set doubleclick to run launch() function
$tree.add_MouseDoubleClick( {
        launch $this.SelectedNode $global:from_mouse
    })

# set right-click to run rightclick function
$tree.add_NodeMouseClick( {
        $whichnode = $this.SelectedNode
        if ($_.Button -eq 'Right') {
            $cat_changed = rightclick($whichnode)
            if ($cat_changed) {
                rebuild_tree
                refresh_cat_list
            }
        }

    })

$tree_KeyDown = [System.Windows.Forms.KeyEventHandler] {
    if ($_.KeyCode -eq 'Enter') {
        launch $tree.SelectedNode $global:from_key
    }
}
    
$tree.add_KeyDown($tree_KeyDown)

# add tree to form
$Form.controls.add($tree)

# close button for the form
$close_btn = New-Object System.Windows.Forms.Button
$close_btn.location = '280,520'
$close_btn.size = '80,25'
$close_btn.Text = 'Close'
$close_btn.Font = '"Arial",10'
$Form.Controls.Add($close_btn)

# save button for the form
$save_btn = New-Object System.Windows.Forms.Button
$save_btn.location = '40,520'
$save_btn.size = '80,25'
$save_btn.Text = 'Save'
$save_btn.Font = '"Arial",10'
$Form.Controls.Add($save_btn)
$save_btn.Add_Click( { 
        save_config
    })

# reload button for the form
$reload_btn = New-Object System.Windows.Forms.Button
$reload_btn.location = '140,520'
$reload_btn.size = '80,25'
$reload_btn.Text = 'Reload'
$reload_btn.Font = '"Arial",10'
$Form.Controls.Add($reload_btn)
$reload_btn.Add_Click( { 
        load_config
        rebuild_tree
    })

# build prompt_form - Popup to ask a question
$prompt_form = New-Object system.Windows.Forms.Form
$prompt_form.Size = New-Object System.Drawing.Size(320, 130)
$prompt_form.Text = 'question'
$prompt_form.FormBorderStyle = 'FixedDialog'
$prompt_form.StartPosition = 'CenterScreen'

# txtbox is the box you type the answer in
$txtbox = New-Object System.Windows.Forms.TextBox
$txtbox.size = '270,100'
$txtbox.Location = '15,30'
$txtbox.AcceptsReturn = $true


$prompt_label = New-Object System.Windows.Forms.Label
$prompt_label.Text = 'Enter'
$prompt_label.Location = '5,5'
$prompt_label.Size = '280,20'
$prompt_form.Controls.Add($prompt_label) 
$prompt_form.Controls.Add($txtbox)

# close button for prompt form
$prmpt_OKbtn = New-Object System.Windows.Forms.Button
$prmpt_OKbtn.location = '80,55'
$prmpt_OKbtn.size = '40,25'
$prmpt_OKbtn.Text = 'OK'
$prmpt_OKbtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
$prompt_form.Controls.Add($prmpt_OKbtn)

# cancel button for prompt form
$prmpt_Cnclbtn = New-Object System.Windows.Forms.Button
$prmpt_Cnclbtn.location = '180,55'
$prmpt_Cnclbtn.size = '50,25'
$prmpt_Cnclbtn.Text = 'Cancel'
$prmpt_Cnclbtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$prompt_form.CancelButton = $prmpt_Cnclbtn
$prompt_form.Controls.Add($prmpt_Cnclbtn)

# make hitting enter in the txtbox do the same as if we
# pressed the close button
$txtbox_KeyDown = [System.Windows.Forms.KeyEventHandler] {
    if ($_.KeyCode -eq 'Enter') {
        $prmpt_OKbtn.PerformClick()
    }
}

$txtbox.add_KeyDown($txtbox_KeyDown)

# cat_list_form is the category list form
$cat_list_form = New-Object system.Windows.Forms.Form
$cat_list_form.Size = New-Object System.Drawing.Size(320, 200)
$cat_list_form.Text = 'Choose a category'
$cat_list_form.FormBorderStyle = 'FixedDialog'
$cat_list_form.StartPosition = 'CenterParent'

# OK button for cat list form
$cl_OKbtn = New-Object System.Windows.Forms.Button
$cl_OKbtn.Location = New-Object System.Drawing.Point(75, 120)
$cl_OKbtn.Size = New-Object System.Drawing.Size(75, 23)
$cl_OKbtn.Text = 'OK'
$cl_OKbtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
$cat_list_form.AcceptButton = $cl_OKbtn
$cat_list_form.Controls.Add($cl_OKbtn)

# cancel button for cat list form
$cl_CnclBtn = New-Object System.Windows.Forms.Button
$cl_CnclBtn.Location = New-Object System.Drawing.Point(150, 120)
$cl_CnclBtn.Size = New-Object System.Drawing.Size(75, 23)
$cl_CnclBtn.Text = 'Cancel'
$cl_CnclBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cat_list_form.CancelButton = $cl_CnclBtn
$cat_list_form.Controls.Add($cl_CnclBtn)

# add a label 
$cl_label = New-Object System.Windows.Forms.Label
$cl_label.Location = New-Object System.Drawing.Point(10, 20)
$cl_label.Size = New-Object System.Drawing.Size(280, 20)
$cl_label.Text = 'Please select a category:'
$cat_list_form.Controls.Add($cl_label)

# list box to choose a category
$cat_listBox = New-Object System.Windows.Forms.listBox
$cat_listBox.Location = New-Object System.Drawing.Point(10, 40)
$cat_listBox.Size = New-Object System.Drawing.Size(260, 20)
$cat_listBox.Height = 80

# add categories to listbox
refresh_cat_list

$cat_list_form.Controls.Add($cat_listBox)

# make hitting enter in the list box do the same as if we
# pressed the OK button
$lstbox_KeyDown = [System.Windows.Forms.KeyEventHandler] {
    if ($_.KeyCode -eq 'Enter') {
        $cl_OKbtn.PerformClick()
    }
}

$cat_listBox.add_KeyDown($lstbox_KeyDown)

# set doubleclick to run launch() function
$cat_listBox.add_MouseDoubleClick( {
        $cl_OKbtn.PerformClick()
    })

# if close button is pressed run cleanup function
$close_btn.Add_Click( { $Form.Close()
        cleanup
    })

# kick off actual script by showing main form
$Form.ShowDialog()

# we get here if user closed the form but not by the 
# close button
cleanup