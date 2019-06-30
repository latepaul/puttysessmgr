# Putty Session Launcher

## What it Does

Putty Session Launcher will read in putty sessions from the registry and 
create a diaglog box with a tree control. 

The First 5 of these are added to a 'Favourites' category.

If you double-click on a session name it will launch putty with that session.

If you right-click it will open a dialog that will offer to change the 
category of that item. Currently that doesn't actually do anything.

## TODO

In no particular order, as they say,

* make the new categories active (i.e. it recreates the tree and re-displays)
* save/restore to a config file (registry?)
* build tree as you load from file - compare list from registry, with node-cat pairings from config file
* Add refresh button which reloads config and reloads registry and rebuilds tree and lists - so if we add a session we can see it straight away
* save button? 
* save on exit automatically - or prompt-for-save-on-exit automatically.
* prompt if putty is not in expected path
* make none a special category not a real category?
* better special category handling
* if user cancels category list form, it's the same as choosing new - maybe don't do that?
* remove unused categories from category list - don't so at the moment which means they still appear in the category list form 
* bring main form to front on start (doesn't do this when launched from vscode)
* resizeable main form?
* UI polish
* make this list issues on github
* ~~Add a drop-down list of current categories + "new category"~~
* ~~make example category "Favourites" instead of Red, Green, Blue~~

