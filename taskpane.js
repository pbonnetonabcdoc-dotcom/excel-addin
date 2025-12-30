// Attendre que le DOM soit chargé

let apiKey = ""; // Récupérée lors de la connexion
let apiString = ""; // Récupérée lors de la connexion
let isConnected = false;

console.log("Le script taskpane.js est chargé !");

async function connectionKSL(user, password) {
	showMessage("Connexion en cours...", false); // Message de chargement

	const urlConnection =
	  "https://demo-2.saas.naelan-software.com:8343/ksl_office83_demo83/webservices/runService?" +
	  `kslUser=${encodeURIComponent(user)}&` +
	  `kslPassword=${encodeURIComponent(password)}&` +
	  "kslProjectName=PROPOSAL_MANAGER_DEMO&" +
	  "kslProjectVersion=1&" +
	  "kslProjectState=Developpement&" +
	  "kslService=SC_ExcelConnection";

    try {
        // Afficher une notification ou un message de chargement
        console.log("Connexion en cours...");

        // Effectuer la requête HTTP
		/* debug       const response = await fetch(urlConnection, {
            method: "POST",
            headers: {
                "Content-Type": "application/xml",
            },
        }); */
		response.ok = false // debug
			
        if (response.ok) {
            const xmlResponse = await response.text();
			showMessage("Retour de connexion OK !", false); // Message de succès
            console.log("Réponse XML reçue :", xmlResponse);
			
            // Analyser la réponse XML
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, "text/xml");

            // Récupérer les valeurs des balises <APIKEY> et <APISTRING>
            const apiKeyNode = xmlDoc.querySelector("APIKEY");
            const apiStringNode = xmlDoc.querySelector("APISTRING");

            if (apiKeyNode && apiStringNode) {
                apiKey = apiKeyNode.textContent;
                apiString = apiStringNode.textContent;
                isConnected = true;
                console.log("Connexion réussie ! API Key :", apiKey);
                return { success: true, apiKey, apiString };
            } else {
                console.error("Erreur - Impossible de récupérer les éléments de connexion à l'IA");
                return { success: false, message: "Impossible de récupérer les éléments de connexion à l'IA" };
			}
			
        } else {
			showMessage(`Erreur : ${response.status} - ${response.statusText}`, true); // Message d'erreur
            console.error(`Erreur lors de l'appel au web service. Statut : ${response.status} - ${response.statusText}`);
			
			// debug
			apiKey = "FVOsUbRvbi0uzFMTkPUjP947Bfte15iE"
			apiString = "https://api.mistral.ai/v1/chat/completions"
			isConnected = true
			return { success: true, message: `Debug` };
		
			// not debug
			return { success: false, message: `Erreur lors de l'appel au web service. Statut : ${response.status} - ${response.statusText}` };
        }
    } catch (error) {
		showMessage(`Erreur inattendue : ${error.message}`, true); // Message d'erreur
        console.error("Erreur lors de la connexion :", error);
		
		// debug
		apiKey = "FVOsUbRvbi0uzFMTkPUjP947Bfte15iE"
		apiString = "https://api.mistral.ai/v1/chat/completions"
		isConnected = true
		return { success: true, message: `Debug` };
		
		// not debug
        return { success: false, message: `Erreur lors de la connexion : ${error.message}` };
    }
}

function showMessage(message, isError = false) {
    // Supprimer les messages précédents
    const existingMessages = document.querySelectorAll(".message");
    existingMessages.forEach(msg => msg.remove());

    // Créer un conteneur pour le message
    const messageElement = document.createElement("div");
    messageElement.className = "message";
    messageElement.textContent = message;
    messageElement.style.padding = "10px";
    messageElement.style.margin = "10px 0";
    messageElement.style.borderRadius = "4px";
    messageElement.style.backgroundColor = isError ? "#ffdddd" : "#ddffdd"; // Fond rouge clair ou vert clair
    messageElement.style.color = isError ? "#d8000c" : "#4f8a10"; // Texte rouge ou vert
    messageElement.style.textAlign = "center";

    // Ajouter le message en haut du formulaire
    const form = document.getElementById("loginForm");
    form.insertBefore(messageElement, form.firstChild);
}

function fermerTaskPane() {
    if (Office.addin && Office.addin.hide) {
        Office.addin.hide();
    } else {
        console.error("Impossible de fermer la TaskPane : Office.addin.hide n'est pas disponible.");
    }
}

async function ecrireDansCellule(ligne, colonne, valeur) {
    await Excel.run(async (context) => {
        // Récupère la feuille active
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Cible la cellule à la ligne et colonne spécifiées
        // Note : Les indices de ligne et colonne commencent à 0
        const cell = sheet.getCell(ligne, colonne);

        // Définit la valeur de la cellule
        cell.values = [[valeur]];
		
		
        // Applique une mise en forme (optionnel)
        cell.format.wrapText = true;
		cell.format.fill.color = "#E8F0FE";
        cell.format.verticalAlignment = "Top"
    });
}

async function demanderReponse() {
    try {
        if (!isConnected) {
            showMessage("Veuillez vous connecter d'abord.", true);
            return;
        }

	await Excel.run(async (context) => {

		const ranges = context.workbook.getSelectedRanges();
		ranges.areas.load("items");

		const responseCell = context.workbook.getActiveCell();
		responseCell.load(["rowIndex", "columnIndex"]);

		await context.sync();

		// Charger les propriétés nécessaires
		for (const area of ranges.areas.items) {
			area.load(["rowCount", "columnCount", "values", "rowIndex", "columnIndex"]);
		}
		await context.sync();

		let questionValue = null;

		for (const area of ranges.areas.items) {
			for (let r = 0; r < area.rowCount; r++) {
				for (let c = 0; c < area.columnCount; c++) {

					const absoluteRow = area.rowIndex + r;
					const absoluteCol = area.columnIndex + c;

					// Ignorer la cellule active (réponse)
					if (
						absoluteRow === responseCell.rowIndex &&
						absoluteCol === responseCell.columnIndex
					) {
						continue;
					}

					// Première autre cellule = question
					if (questionValue === null) {
						questionValue = area.values[r][c];
					}
				}
			}
		}

		if (questionValue === null) {
			showMessage(
				"Veuillez sélectionner au moins 2 cellules.\n" +
				"La dernière cellule cliquée sera la réponse.",
				true
			);
			return;
		}

		// Effacer la cellule réponse
		responseCell.values = [[""]];
		await context.sync();

		console.log("Question :", questionValue);
		console.log("Réponse :", responseCell.rowIndex, responseCell.columnIndex);
		console.log("apiKey :", apiKey);
		console.log("apiString :", apiString);
		
		// Construction du payload avec instructions pour une réponse concise et professionnelle
		const payload = {
			model: "mistral-small-latest",
			messages: [
				{
					role: "system",
					content: "Réponds en moins de 50 mots, avec un style professionnel, clair et concis. Utilise un ton neutre et évite les formules de politesse superflues."
				},
				{
					role: "user",
					content: questionValue // `question` est la variable contenant la question de l'utilisateur
				}
			]
		};

		// Configuration de la requête HTTP
		try {
			const response = await fetch(apiString, {
				method: "POST",
				headers: {
					"Content-Type": "application/json",
					"Authorization": `Bearer ${apiKey}`
				},
				body: JSON.stringify(payload) // Convertit l'objet payload en chaîne JSON
			});

			if (!response.ok) {
				throw new Error(`Erreur HTTP : ${response.status} - ${response.statusText}`);
			}

			// Récupération de la réponse
			const responseData = await response.json();
			console.log("Réponse reçue :", responseData);

			// Traite la réponse ici (par exemple, l'afficher dans Excel)
			if (responseData.choices && responseData.choices.length > 0) {
				const messageContent = responseData.choices[0].message.content;
				console.log("Réponse de l'API :", messageContent);
				ecrireDansCellule(responseCell.rowIndex, responseCell.columnIndex, messageContent)
			} else {
				throw new Error("Aucune réponse valide dans 'choices'.");
			}
			
		} catch (error) {
			console.error("Erreur lors de la requête :", error);
			throw error; // Propage l'erreur pour une gestion ultérieure
		}
	});

    } catch (error) {
        console.error(error);
        showMessage("Erreur Excel : " + error.message, true);
    }
}

Office.onReady(() => {

    const btn = document.getElementById("btnDemanderReponse");

    btn.addEventListener("click", async () => {

    const login = document.getElementById("login").value;
    const password = document.getElementById("password").value;
		
	if (!isConnected) {
		const result = await connectionKSL(login, password);

		if (!result.success) {
			showMessage("Échec de la connexion : " + result.message, true);
			return;
		}
		showMessage("Connexion réussie !", false);
		
		fermerTaskPane();
		}
		
		// Exécute demanderReponse() en arrière-plan
        demanderReponse().catch(error => {
			console.error("Erreur dans demanderReponse :", error);
        });
    });
});

