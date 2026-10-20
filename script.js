// ==================================================
// COMPTEUR SCIERIE
// VERSION COMPLÉMENT EXCEL
// ==================================================

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        demarrerCompteurScierie();

    }

});


// ==================================================
// DÉMARRAGE
// ==================================================

function demarrerCompteurScierie() {

    // Boutons quantité

    document.getElementById("plus1")
        .addEventListener("click", () => ajouterModification(1));

    document.getElementById("moins1")
        .addEventListener("click", () => ajouterModification(-1));

    document.getElementById("plus5")
        .addEventListener("click", () => ajouterModification(5));

    document.getElementById("moins5")
        .addEventListener("click", () => ajouterModification(-5));

    document.getElementById("plus10")
        .addEventListener("click", () => ajouterModification(10));

    document.getElementById("moins10")
        .addEventListener("click", () => ajouterModification(-10));


    // FINI

    document.getElementById("fini")
        .addEventListener("click", () => basculerFini());


    // MAJ

    document.getElementById("maj")
        .addEventListener("click", () => miseAJour());


    // ANNUL

    document.getElementById("annul")
        .addEventListener("click", () => annulerDerniereAction());


    // Chargement des onglets

    verifierOnglets();


    // Vérification toutes les 30 secondes

    setInterval(
        verifierOnglets,
        30000
    );
}


// ==================================================
// VARIABLES COMPTEUR
// ==================================================

let changementEnAttente = 0;

let cumulVisuel = 0;

let timerModification = null;

let timerEffacementAffichage = null;

let ecritureEnCours = false;


// ==================================================
// MÉMOIRE DE LA DERNIÈRE ACTION
// ==================================================

let derniereAction = null;


// ==================================================
// AJOUT MODIFICATION QUANTITÉ
// ==================================================

function ajouterModification(changement) {

    changementEnAttente += changement;

    cumulVisuel += changement;

    afficherCumul();


    // ==================================================
    // CUMUL VISIBLE PENDANT 3 SECONDES
    // ==================================================

    if (timerEffacementAffichage !== null) {

        clearTimeout(
            timerEffacementAffichage
        );

    }


    timerEffacementAffichage =
        setTimeout(() => {

            cumulVisuel = 0;

            const zone =
                document.getElementById("attente");

            zone.textContent = "";

            timerEffacementAffichage = null;

        }, 3000);


    // ==================================================
    // REGROUPEMENT DES CLICS RAPIDES
    // ==================================================

    if (timerModification !== null) {

        clearTimeout(
            timerModification
        );

    }


    timerModification =
        setTimeout(() => {

            timerModification = null;

            traiterModifications();

        }, 120);
}


// ==================================================
// AFFICHAGE DU CUMUL
// ==================================================

function afficherCumul() {

    const zone =
        document.getElementById("attente");


    if (cumulVisuel > 0) {

        zone.textContent =
            "+" + cumulVisuel;

    }

    else if (cumulVisuel < 0) {

        zone.textContent =
            cumulVisuel.toString();

    }

    else {

        zone.textContent = "0";

    }
}


// ==================================================
// MODIFICATION DE LA QUANTITÉ DANS EXCEL
// ==================================================

async function traiterModifications() {

    if (ecritureEnCours) {
        return;
    }


    if (changementEnAttente === 0) {
        return;
    }


    ecritureEnCours = true;


    const changement =
        changementEnAttente;


    changementEnAttente = 0;


    try {

        await Excel.run(async (context) => {

            const cellule =
                context.workbook
                    .getSelectedRange();


            cellule.load([
                "values",
                "address"
            ]);


            const feuille =
                context.workbook
                    .worksheets
                    .getActiveWorksheet();


            feuille.load("name");


            await context.sync();


            let valeur =
                cellule.values[0][0];


            if (
                valeur === "" ||
                valeur === null
            ) {

                valeur = 0;

            }


            valeur =
                Number(valeur);


            if (!isNaN(valeur)) {


                // ==================================================
                // MÉMORISATION POUR ANNUL
                // ==================================================

                derniereAction = {

                    type: "quantite",

                    feuille:
                        feuille.name,

                    adresse:
                        cellule.address,

                    ancienneValeur:
                        valeur

                };


                // ==================================================
                // NOUVELLE VALEUR
                // ==================================================

                cellule.values = [
                    [
                        valeur + changement
                    ]
                ];


                await context.sync();

            }

        });

    }

    catch (erreur) {

        console.log(
            "Erreur compteur :",
            erreur
        );


        changementEnAttente +=
            changement;

    }

    finally {

        ecritureEnCours = false;


        if (
            changementEnAttente !== 0
        ) {

            setTimeout(
                () => {

                    traiterModifications();

                },
                20
            );

        }

    }

}


// ==================================================
// FINI
//
// BARRE B C D E
// DE LA LIGNE SÉLECTIONNÉE
// ==================================================

async function basculerFini() {

    try {

        await Excel.run(async (context) => {

            const selection =
                context.workbook
                    .getSelectedRange();


            selection.load(
                "rowIndex"
            );


            const feuille =
                context.workbook
                    .worksheets
                    .getActiveWorksheet();


            feuille.load("name");


            await context.sync();


            const numeroLigne =
                selection.rowIndex + 1;


            const adresse =
                "B" +
                numeroLigne +
                ":E" +
                numeroLigne;


            const zone =
                feuille.getRange(
                    adresse
                );


            zone.format.font.load(
                "strikethrough"
            );


            await context.sync();


            const ancienEtat =
                zone.format.font
                    .strikethrough === true;


            // ==================================================
            // MÉMORISATION POUR ANNUL
            // ==================================================

            derniereAction = {

                type: "fini",

                feuille:
                    feuille.name,

                adresse:
                    adresse,

                ancienEtat:
                    ancienEtat

            };


            // ==================================================
            // BARRE / DÉBARRE
            // ==================================================

            zone.format.font
                .strikethrough =
                    !ancienEtat;


            await context.sync();

        });

    }

    catch (erreur) {

        console.log(
            "Erreur bouton FINI :",
            erreur
        );

    }

}


// ==================================================
// MAJ
//
// EFFACE UNIQUEMENT H I J K
// DE LA LIGNE SÉLECTIONNÉE
// ==================================================

async function miseAJour() {

    try {

        await Excel.run(async (context) => {

            const selection =
                context.workbook
                    .getSelectedRange();


            selection.load(
                "rowIndex"
            );


            const feuille =
                context.workbook
                    .worksheets
                    .getActiveWorksheet();


            feuille.load("name");


            await context.sync();


            const numeroLigne =
                selection.rowIndex + 1;


            // ==================================================
            // H À K DE LA LIGNE SÉLECTIONNÉE
            // ==================================================

            const adresse =
                "H" +
                numeroLigne +
                ":K" +
                numeroLigne;


            const zone =
                feuille.getRange(
                    adresse
                );


            zone.load("values");


            await context.sync();


            const anciennesValeurs =
                zone.values.map(
                    ligne =>
                        ligne.slice()
                );


            // ==================================================
            // MÉMORISATION POUR ANNUL
            // ==================================================

            derniereAction = {

                type: "maj",

                feuille:
                    feuille.name,

                adresse:
                    adresse,

                anciennesValeurs:
                    anciennesValeurs

            };


            // ==================================================
            // EFFACE LE CONTENU
            // CONSERVE LA MISE EN FORME
            // ==================================================

            zone.clear(
                "Contents"
            );


            await context.sync();

        });

    }

    catch (erreur) {

        console.log(
            "Erreur bouton MAJ :",
            erreur
        );

    }

}


// ==================================================
// ANNUL
//
// ANNULE LA DERNIÈRE ACTION
// DU PANNEAU
// ==================================================

async function annulerDerniereAction() {

    if (
        derniereAction === null
    ) {

        afficherMessageAnnul(
            "RIEN À ANNULER"
        );

        return;

    }


    const action =
        derniereAction;


    try {

        await Excel.run(async (context) => {


            const feuille =
                context.workbook
                    .worksheets
                    .getItem(
                        action.feuille
                    );


            // ==================================================
            // ANNUL QUANTITÉ
            // ==================================================

            if (
                action.type ===
                "quantite"
            ) {

                const cellule =
                    feuille.getRange(
                        action.adresse
                    );


                cellule.values = [
                    [
                        action.ancienneValeur
                    ]
                ];

            }


            // ==================================================
            // ANNUL FINI
            // ==================================================

            else if (
                action.type ===
                "fini"
            ) {

                const zone =
                    feuille.getRange(
                        action.adresse
                    );


                zone.format.font
                    .strikethrough =
                        action.ancienEtat;

            }


            // ==================================================
            // ANNUL MAJ
            // ==================================================

            else if (
                action.type ===
                "maj"
            ) {

                const zone =
                    feuille.getRange(
                        action.adresse
                    );


                zone.values =
                    action.anciennesValeurs;

            }


            await context.sync();


            // Une seule annulation possible

            derniereAction = null;


            afficherMessageAnnul(
                "ANNULÉ ✓"
            );

        });

    }

    catch (erreur) {

        console.log(
            "Erreur bouton ANNUL :",
            erreur
        );


        afficherMessageAnnul(
            "ERREUR"
        );

    }

}


// ==================================================
// MESSAGE ANNULATION
// ==================================================

function afficherMessageAnnul(message) {

    const zone =
        document.getElementById(
            "messageAnnul"
        );


    zone.textContent =
        message;


    setTimeout(
        () => {

            if (
                zone.textContent ===
                message
            ) {

                zone.textContent = "";

            }

        },
        2000
    );

}


// ==================================================
// ONGLETS
// ==================================================

let derniereListeOnglets = "";

let verificationEnCours = false;


// ==================================================
// VÉRIFICATION DES ONGLETS
// ==================================================

async function verifierOnglets() {

    if (verificationEnCours) {
        return;
    }


    verificationEnCours = true;


    try {

        await Excel.run(async (context) => {


            const feuilles =
                context.workbook
                    .worksheets;


            feuilles.load(
                "items/name"
            );


            await context.sync();


            const noms =
                feuilles.items.map(
                    feuille =>
                        feuille.name
                );


            const nouvelleListe =
                JSON.stringify(
                    noms
                );


            // ==================================================
            // RECRÉATION UNIQUEMENT SI CHANGEMENT
            // ==================================================

            if (
                nouvelleListe !==
                derniereListeOnglets
            ) {

                derniereListeOnglets =
                    nouvelleListe;


                const zone =
                    document.getElementById(
                        "onglets"
                    );


                zone.innerHTML = "";


                noms.forEach(
                    (nom) => {


                        const bouton =
                            document.createElement(
                                "button"
                            );


                        bouton.textContent =
                            nom;


                        bouton.className =
                            "bouton-onglet";


                        bouton.addEventListener(
                            "click",
                            async () => {

                                await ouvrirOnglet(
                                    nom
                                );

                            }
                        );


                        zone.appendChild(
                            bouton
                        );

                    }
                );

            }

        });

    }

    catch (erreur) {

        console.log(
            "Erreur vérification des onglets :",
            erreur
        );

    }

    finally {

        verificationEnCours = false;

    }

}


// ==================================================
// OUVERTURE D'UN ONGLET
// ==================================================

async function ouvrirOnglet(
    nomOnglet
) {

    try {

        await Excel.run(async (context) => {


            const feuille =
                context.workbook
                    .worksheets
                    .getItem(
                        nomOnglet
                    );


            feuille.activate();


            await context.sync();

        });

    }

    catch (erreur) {

        console.log(
            "Erreur ouverture onglet :",
            erreur
        );

    }

}
