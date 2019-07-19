# IterateAllSubSites
The SharePoint Client Side Object Model is extremely powerful. 

In larger enterprise SharePoint sites I have found it useful to be able to itterate through every site to automate actions.

This simple script can be run from Chrome Web Inspector. I have used it to bulk install apps, activate features, create site maps, enable workflows ect.

*usage*  
Update script and put automated process in the bottom inline function.

itterate((web, clientContext) => {
    //add action here, this will be run on
    //every site within the site collection
    //you run it on (including the current)
	console.log(web.get_url());
})


execute in Web Inspector
