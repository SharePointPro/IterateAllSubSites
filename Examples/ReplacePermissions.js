/** UPDATE PERMISSION FUNCTIONS */


let NEW_PERMISSION = 'Full Control - No Subsite Creation';
//any group with this permission, update with new permission
let REPLACE_PERMISSION = 'Full Control';
//Dont replace permission if its the below group name
let EXCEPT_PERMISSION = 'BQ Administrators';

//itterate through all the roles on the site, and update the permissions if they are currently REPLACE_PERMISSION
//them with the above function 
const updateRolesWithPermission = async (clientContext) => {
    let web = clientContext.get_web();
    let roles = web.get_roleAssignments();
    clientContext.load(roles);
    await processRoleAssignments(clientContext, roles);
}

//Iterate through all the role assignments. Role Assignments have Role Defintion Bindings which hold the permission level
//and a member which is either the group or user that the permissions relate to
//More information: https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff409736%28v%3doffice.14%29
const processRoleAssignments = async (clientContext, roles) => {
    return new Promise((resolve, reject) => {
        clientContext.executeQueryAsync(async () => {
            let roleEnumerator = roles.getEnumerator();
            while (roleEnumerator.moveNext()) {
                let currentRole = roleEnumerator.get_current();
                let roleDefinitionBindings = currentRole.get_roleDefinitionBindings()
                let member = currentRole.get_member();
                clientContext.load(roleDefinitionBindings);
                clientContext.load(member);
                await processBindings(clientContext, roleDefinitionBindings, member);
            }
            resolve();
        }, (err, msg) => {
            reject();
            console.log("error", msg);
        });
    });
}

//Bindings hold the actual permission level (ie "Full Control")
const processBindings = async (clientContext, roleDefinitionBindings, member) => {
    return new Promise((resolve, reject) => {
        console.log("processBindings");
        clientContext.executeQueryAsync(async () => {
            let bindingEnumerator = roleDefinitionBindings.getEnumerator();
            while (bindingEnumerator.moveNext()) {
                let roleDefBinding = bindingEnumerator.get_current();            
                if (roleDefBinding.get_name() === REPLACE_PERMISSION) {
                    if (member.get_title() !== EXCEPT_PERMISSION) {
                        console.log("updating permission");
                        await updatePermission(member.get_id(), 
                        NEW_PERMISSION, 
                        clientContext);
                        console.log("member updated: ", member.get_title());
                    }
                }
            }
            resolve();
        }, (err, msg) => {
            reject();
            console.log("error", msg)
        });
    })
}


//Update the permission of the group
const updatePermission = (groupMembershipID,
    permissionLevelName,
    clientContext) => {
    return new Promise((resolve, reject) => {
        console.log("update Permission");
        let web = clientContext.get_web();
        //Get Role Definition
        let roleDef = web.get_roleDefinitions().getByName(permissionLevelName);
        let roleDefBinding = SP.RoleDefinitionBindingCollection.newObject(clientContext);
        // Add the role to the role definiiton binding.
        roleDefBinding.add(roleDef);
        // Get the RoleAssignmentCollection for the web.
        let assignments = web.get_roleAssignments();
        //Get the role assignment for the group using Group  's membership id
        let groupRoleAssignment = assignments.getByPrincipalId(groupMembershipID);
        groupRoleAssignment.importRoleDefinitionBindings(roleDefBinding);
        groupRoleAssignment.update();
        clientContext.executeQueryAsync(function () {
            resolve();
        },
            function (sender, args) {
                console.log(args.get_message());
                reject();
            });
    })

}

/* END UPDATE PERMISSION FUNCTION */


/* ITERATION FUNCTIONS - SEE HERE https://github.com/SharePointPro/IterateAllSubSites/ */
const iterate = async (functionToCall) => {
	let clientContext;

	//Recursive function, gets all the webs in the supplied web
	const getWebs = async function (web) {
		return new Promise(async function (resolve, reject) {
			await getWeb(web);
			let subWebs = web.get_webs();
			clientContext.load(subWebs);
			clientContext.executeQueryAsync(async function () {
				let webEnumerator = subWebs.getEnumerator();
				while (webEnumerator.moveNext()) {
					let currentWeb = webEnumerator.get_current();
					await getWebs(currentWeb);
				}
				resolve();
			}, function () {
				console.log("error");
			});
		});
	}

	//Loads the web
	const getWeb = async function (web) {
		return new Promise(async function (resolve, reject) {
			clientContext.load(web);
			clientContext.executeQueryAsync(async function () {
				await executeAction(web);
				resolve();
			}, function () {
				console.log("error");
			});

		});
	}

	//Action is done here
	const executeAction = async function (web) {
		await functionToCall(web, clientContext);
	}

	clientContext = SP.ClientContext.get_current();
    await getWebs(clientContext.get_web());
    console.log("finished");
}




//ENTRY POINT
iterate(async (web, clientContext) => {
    await updateRolesWithPermission(clientContext)
	console.log(web.get_url());
})

