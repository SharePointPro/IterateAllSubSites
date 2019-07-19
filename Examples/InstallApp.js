
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
		functionToCall(web, clientContext);
	}

	//Entry Point
	clientContext = SP.ClientContext.get_current();
	await getWebs(clientContext.get_web());
}

/* End Recursive iteration function */

function inject(u, i) {
	var d = document;
	if (!d.getElementById(i)) {
		var s = d.createElement('script');
		s.src = u;
		s.id = i;
		d.body.appendChild(s);
	}
}


//Install App, in this case it is in the SiteCollection App Catalog, but may also be in tennant App Catalog
//I have also predetermined the Apps ID
//More Information here: https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
function installApp(url) {
    const app_id = 'cc99611f-ebfc-4c26-bb19-be0125b334f8';
	return new Promise((resolve, reject) => {
		$.ajax({  
			url: `${url}/_api/web/SiteCollectionAppCatalog/AvailableApps/GetById('${app_id}')/Install`,
			 method: "POST" ,
			 headers:  
				 {  
					 "Accept": "application/json;odata=nometadata",    
					 "X-RequestDigest": $("#__REQUESTDIGEST").val(),
				 },
			 cache: false,  
			 success: function(data)   
			 {  
				 resolve();
				  console.log("App installed successfully");
			 },  
			 error: function(data)  
			 {  
				 resolve();
				 console.log(data);
			 }  
		 });  
	  
	});
} 



//Inject Jquery into Page
//Makes REST Calls easier
inject('https://code.jquery.com/jquery-3.2.1.min.js', 'jquery');


//I wrap the entry point in a settimeout to give jquery time to download
setTimeout(() => {
    iterate((web, clientContext) => {
    await installApp(web.get_url);
    })
},2000);


