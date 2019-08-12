/* Iterate through every subsite and check all document libraries have versioning turned on, if they do not, turn it on */

const iterate = async (functionToCall) => {
    let clientContext;
    let counted = 0;
    let totalWithoutVersioning = 0;

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
        var countedObject = await functionToCall(web, clientContext);
        counted += countedObject.counted;
        totalWithoutVersioning += countedObject.totalWithoutVersioning;
    }

    //Entry Point
    clientContext = SP.ClientContext.get_current();
    await getWebs(clientContext.get_web());
    console.log("finished");
    console.log(`${counted} counted`);
    console.log(`${totalWithoutVersioning} without versioning`);
}

/* Above is itteration functions */


iterate((web, clientContext) => {
    return new Promise((resolve, reject) => {
        const excludeList = ["Drop Off Library", "Site Assets"];
        let counted = 0;
        let totalWithoutVersioning = 0;
        var lists = web.get_lists();
        clientContext.load(lists);
        clientContext.executeQueryAsync(async () => {
            var listEnumerator = lists.getEnumerator();
            while (listEnumerator.moveNext()) {
                var currentList = listEnumerator.get_current();
                if (currentList.get_baseType() === SP.BaseType.documentLibrary && !currentList.get_hidden() && excludeList.indexOf(currentList.get_title()) === -1) {
                    counted++;
                    if (!currentList.get_enableVersioning()) {
                        totalWithoutVersioning++;
                        currentList.set_enableVersioning(true);
                        currentList.update()
                        await new Promise((innerResolve, innerReject) => {
                            console.log(currentList.get_parentWebUrl() + "/" + currentList.get_title());
                            clientContext.executeQueryAsync(() => {
                                innerResolve();
                            }, (err, msg) => {
                                console.log("There was an error:", msg);
                                reject(innerReject);
                            })
                        });
                    }
                }
            }
            resolve({ counted, totalWithoutVersioning });
        },
            (err, msg) => {
                console.log("There was an error:", msg);
                reject(msg);
            }
        );
    });
});
