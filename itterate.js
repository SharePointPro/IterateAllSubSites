
const itterate = async (functionToCall) => {
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

itterate((web, clientContext) => {
    //add action here, this will be run on
    //every site within the site collection
    //you run it on (including the current)
	console.log(web.get_url());
})


