# Putty Session Launcher

## What it Does

Putty Session Launcher will read in putty sessions from the registry and 
create a diaglog box with a tree control. 

The First 6 of these are added to categories (Red, Green, Blue). Think of these
as folders. 

If you double-click on a session name it will launch putty with that session.

If you right-click it will open a dialog that will offer to change the 
category of that item. Currently that doesn't actually do anything.

## TODO

In no particular order, as they say,

* make the new categories active (i.e. it recreates the tree and re-displays
* Add a drop-down list of current categories + "new category"
* save/restore to a config file (registry?)
* prompt if putty is not in expected path
* make example category "Favourites" instead of Red, Green, Blue
* UI polish
