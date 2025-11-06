function onEdit(e) {
  const feuille = e.source.getActiveSheet();
  const nomFeuille = feuille.getName();
  const ligne = e.range.getRow();
  const colonne = e.range.getColumn();

  if (ligne === 1) return;

  const statutCol = 2;
  const emailCol = 1;
  const messageCol = 50;
  const valeur = String(e.range.getValue()).trim();

  // 🎯 TRAITER_AGREMENTS : mise en forme + protection + synchronisation
  if (nomFeuille === "Traiter_Agrements") {
    const feuilleSuivi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suivre_Agrements");

    // 🔒 Protection si email saisi
    if (colonne === emailCol && valeur.includes("@")) {
      const plageLigne = feuille.getRange(ligne, 1, 1, feuille.getLastColumn());
      const protections = feuille.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      const dejaProtegee = protections.some(p => p.getRange().getRow() === ligne);
      if (!dejaProtegee) {
        const protection = plageLigne.protect();
        protection.setDescription(`Ligne réservée à ${valeur}`);
        protection.removeEditors(protection.getEditors());
        protection.addEditor(valeur);
        protection.setWarningOnly(false);
        feuille.getRange(ligne, messageCol).setValue(`🔒 Ligne occupée par ${valeur}`);
        appliquerMiseEnForme(feuille);
      }
    }

    // 🔓 Suppression si ligne vide
    const plageLigne = feuille.getRange(ligne, 1, 1, feuille.getLastColumn());
    const valeursLigne = plageLigne.getValues()[0];
    const toutesVides = valeursLigne.every(val => val === "" || val === null);
    if (toutesVides) {
      const protections = feuille.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      protections.forEach(p => {
        if (p.getRange().getRow() === ligne) {
          p.remove();
          feuille.getRange(ligne, messageCol).clearContent();
        }
      });
      appliquerMiseEnForme(feuille);
      return;
    }

    // 🎨 Mise en forme selon statut
    if (colonne === statutCol) {
      let colonnesObligatoires = [];
      let couleurFond = null;
      let couleurSuivi = null;

      if (valeur === "Traitement Agrément") {
        colonnesObligatoires = [1, 2, 3, 4, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 41, 42];
        couleurFond = "#fff8b3";
        couleurSuivi = "#ffffff";
      } else if (valeur === "Mise à jour Agrément") {
        colonnesObligatoires = [1, 2, 3, 4, 5, 7, 8, 10, 17, 41, 42];
        couleurFond = "#e6d4f7";
        couleurSuivi = "#e6d4f7";
      } else if (valeur === "Accusé de Réception") {
        colonnesObligatoires = [1, 2, 3, 4, 5, 6, 7, 8, 10, 41, 42];
        couleurFond = "#d0f0fd";
        couleurSuivi = "#d0f0fd";
      }

      const plageTotale = feuille.getRange(ligne, 1, 1, 42);
      plageTotale.setBackground(null);
      if (valeur !== "" && colonnesObligatoires.length > 0) {
        colonnesObligatoires.forEach(col => {
          feuille.getRange(ligne, col).setBackground(couleurFond);
        });
      }

      // 🎯 Mise en forme dans Suivre_Agrements
      if (feuilleSuivi) {
        const plageSuivi = feuilleSuivi.getRange(ligne, 1, 1, 21);
        plageSuivi.setBackground(null);
        if (valeur !== "") {
          plageSuivi.setBackground(couleurSuivi);
        }
      }

      // 🔄 Synchronisation vers fichier cible
      //synchroniserSuivreAgrementsVersDestination();
    }
  }

  // 🎨 SUIVRE_AGREMENTS : saisie manuelle
  if (nomFeuille === "Suivre_Agrements" && colonne === 1) {
    let couleurFond = null;
    if (valeur === "Traitement Agrément") couleurFond = "#ffffff";
    else if (valeur === "Mise à jour Agrément") couleurFond = "#e6d4f7";
    else if (valeur === "Accusé de Réception") couleurFond = "#d0f0fd";

    const plage = feuille.getRange(ligne, 1, 1, 21);
    plage.setBackground(null);
    if (valeur !== "") plage.setBackground(couleurFond);

    // 🔄 Synchronisation vers fichier cible
    synchroniserSuivreAgrementsVersDestination();
  }
}

//***************************
function synchroniserSuivreAgrementsVersDestination() {
  const idDestination = "1TPO3vc_lBqaw1KfhU0SOQwTfTl2MDCDIHkLPskXUqqg";
  const nomFeuilleDestination = "Recap Dossiers Agréments";
  const couleurFondEntete = "#fce8b2";

  try {
    const feuilleSource = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suivre_Agrements");
    const classeurDestination = SpreadsheetApp.openById(idDestination);
    let feuilleDestination = classeurDestination.getSheetByName(nomFeuilleDestination);

    const lastRowSource = feuilleSource.getLastRow();
    const lastColSource = feuilleSource.getLastColumn();

    feuilleDestination = initialiserFeuilleDestination(classeurDestination, feuilleDestination, nomFeuilleDestination);
    copierEntete(feuilleSource, feuilleDestination, lastColSource, couleurFondEntete);
    copierDonnees(feuilleSource, feuilleDestination, lastRowSource, lastColSource);
    copierDimensions(feuilleSource, feuilleDestination, lastRowSource, lastColSource);
    copierProtections(feuilleSource, feuilleDestination);
    copierFigements(feuilleSource, feuilleDestination);

    SpreadsheetApp.getActiveSpreadsheet().toast("✅ Synchronisation complète", "Info", 5);
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast("❌ Erreur : " + error.message, "Synchronisation", 10);
  }
}

function initialiserFeuilleDestination(classeur, feuille, nom) {
  if (!feuille) return classeur.insertSheet(nom);
  const lastRow = feuille.getLastRow();
  if (lastRow > 1) feuille.getRange(2, 1, lastRow - 1, feuille.getMaxColumns()).clearContent();
  return feuille;
}

function copierEntete(source, dest, cols, fond) {
  const entete = source.getRange(1, 1, 1, cols);
  const cible = dest.getRange(1, 1, 1, cols);
  cible.setValues(entete.getValues());
  SpreadsheetApp.flush(); // Sécurise l’exécution avant les styles
  cible.setFontColors(entete.getFontColors())
       .setFontWeights(entete.getFontWeights())
       .setFontStyles(entete.getFontStyles())
       .setFontSizes(entete.getFontSizes())
       .setHorizontalAlignments(entete.getHorizontalAlignments())
       .setVerticalAlignments(entete.getVerticalAlignments())
       .setWraps(entete.getWraps())
       .setBackground(fond);
}

function copierDonnees(source, dest, rows, cols) {
  if (rows <= 1) return;
  const data = source.getRange(2, 1, rows - 1, cols);
  const cible = dest.getRange(2, 1, rows - 1, cols);
  cible.setValues(data.getValues());
  SpreadsheetApp.flush();
  cible.setBackgrounds(data.getBackgrounds())
       .setFontColors(data.getFontColors())
       .setFontWeights(data.getFontWeights())
       .setFontStyles(data.getFontStyles())
       .setFontSizes(data.getFontSizes())
       .setHorizontalAlignments(data.getHorizontalAlignments())
       .setVerticalAlignments(data.getVerticalAlignments())
       .setWraps(data.getWraps());
}

function copierDimensions(source, dest, rows, cols) {
  for (let c = 1; c <= cols; c++) dest.setColumnWidth(c, source.getColumnWidth(c));
  for (let r = 1; r <= rows; r++) dest.setRowHeight(r, source.getRowHeight(r));
}

function copierProtections(source, dest) {
  const protections = source.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    try {
      const plage = p.getRange();
      const clone = dest.getRange(plage.getRow(), plage.getColumn(), plage.getNumRows(), plage.getNumColumns());
      const protection = clone.protect();
      protection.setDescription(p.getDescription());
      protection.removeEditors(protection.getEditors());
      protection.addEditors(p.getEditors());
      protection.setWarningOnly(p.isWarningOnly());
    } catch (e) {
      Logger.log("Erreur protection : " + e.message);
    }
  });
}

function copierFigements(source, dest) {
  dest.setFrozenRows(source.getFrozenRows());
  dest.setFrozenColumns(source.getFrozenColumns());
}

// *************************
function reinitialiserColonnesParProtection() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Annexe 2");
  const utilisateur = Session.getEffectiveUser().getEmail(); // clé : exécution avec autorisations du propriétaire
  const proprietaire = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
  const protections = feuille.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const plageDebut = 2;
  const plageFin = 1000;
  const ui = SpreadsheetApp.getUi();
  const colonnesAutorisees = new Set();

  function numeroVersLettreColonne(n) {
    let lettre = '';
    while (n > 0) {
      let reste = (n - 1) % 26;
      lettre = String.fromCharCode(65 + reste) + lettre;
      n = Math.floor((n - 1) / 26);
    }
    return lettre;
  }

  // Si le propriétaire exécute le script, il peut tout effacer
  if (utilisateur === proprietaire) {
    for (let col = 6; col <= 15; col++) {
      feuille.getRange(plageDebut, col, plageFin - plageDebut + 1).clearContent();
    }
    const lettres = Array.from({ length: 10 }, (_, i) => numeroVersLettreColonne(i + 6)).join(", ");
    ui.alert(`✅ Toutes les colonnes (${lettres}) ont été réinitialisées.`);
    return;
  }

  // Sinon, on vérifie les protections pour l’utilisateur actif
  protections.forEach(protection => {
    const plage = protection.getRange();
    const col = plage.getColumn();
    const nbColonnes = plage.getNumColumns();
    const row = plage.getRow();
    const hauteur = plage.getNumRows();
    const editeurs = protection.getEditors().map(e => e.getEmail());

    const ligneDebutProtection = row;
    const ligneFinProtection = row + hauteur - 1;
    const plageCouvreZone = ligneDebutProtection <= plageFin && ligneFinProtection >= plageDebut;
    const utilisateurAutorise = editeurs.includes(utilisateur);

    if (col >= 6 && col <= 15 && plageCouvreZone && utilisateurAutorise) {
      for (let i = 0; i < nbColonnes; i++) {
        colonnesAutorisees.add(col + i);
      }
    }
  });

  let message = "";

  if (colonnesAutorisees.size > 0) {
    Array.from(colonnesAutorisees).sort((a, b) => a - b).forEach(col => {
      feuille.getRange(plageDebut, col, plageFin - plageDebut + 1).clearContent();
    });
    const lettres = Array.from(colonnesAutorisees)
      .sort((a, b) => a - b)
      .map(numeroVersLettreColonne)
      .join(", ");
    message = `✅ Vos colonnes (${lettres}) ont été réinitialisées.`;
  } else {
    message = "⚠️ Vous n'avez pas l'autorisation de modifier ces colonnes.";
  }

  ui.alert(message);
}


function doGet() {
  return HtmlService.createHtmlOutput("Cette application est prête.");
}

function boutonReinitialisation() {
  reinitialiserColonnesParProtection();
}


function miseAJourFusionDansTraiter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleProd = ss.getSheetByName("Production_Agrements_Autocrat2");
  const feuilleTraiter = ss.getSheetByName("Traiter_Agrements");

  const donneesProd = feuilleProd.getDataRange().getValues();
  const donneesTraiter = feuilleTraiter.getDataRange().getValues();

  const COL_PRODUIRE_DOC = 41; // AP
  const COL_FUSION = 42;       // AQ
  const COL_DEBUT_VERIF = 43;  // AR
  const COL_FIN_VERIF = 54;    // BC

  let messages = [];
  let incoherences = [];
  let lignesIncoherentes = [];

  for (let i = 1; i < donneesProd.length; i++) {
    const ligneProd = donneesProd[i];
    const produireDoc = ligneProd[COL_PRODUIRE_DOC];
    const fusionActuelle = ligneProd[COL_FUSION];
    const plageVerif = ligneProd.slice(COL_DEBUT_VERIF, COL_FIN_VERIF + 1);

    const contientValeur = plageVerif.some(cell => cell !== "" && cell !== null);
    const contientLienDoc = plageVerif.some(cell => typeof cell === "string" && cell.includes("https://docs.google.com/document/d"));

    if (!produireDoc || produireDoc === "Non") {
      messages.push(`Ligne ${i + 1} : l'utilisateur n'a pas encore décidé de passer à la production`);
      continue;
    }

    if (fusionActuelle === "TRUE" && !contientValeur) {
      incoherences.push(`Ligne ${i + 1} : AQ = TRUE mais aucune donnée dans AR à BC`);
      lignesIncoherentes.push(i + 1);
      continue;
    }

    if (fusionActuelle === "FALSE" && contientValeur) {
      if (contientLienDoc) {
        feuilleTraiter.getRange(i + 1, COL_FUSION + 1).setValue("TRUE");
        continue;
      } else {
        incoherences.push(`Ligne ${i + 1} : AQ = FALSE mais des données sont présentes dans AR à BC sans lien de document`);
        lignesIncoherentes.push(i + 1);
        continue;
      }
    }

    const nouvelleValeur = contientValeur ? "TRUE" : "FALSE";
    if (fusionActuelle !== nouvelleValeur) {
      feuilleTraiter.getRange(i + 1, COL_FUSION + 1).setValue(nouvelleValeur);
    }

    const produireTraiter = donneesTraiter[i][COL_PRODUIRE_DOC];
    if (["FALSE", "Non"].includes(produireTraiter)) {
      messages.push(`Ligne ${i + 1} : Production non encore effectuée`);
    } else if (!["TRUE", "Oui"].includes(produireTraiter)) {
      messages.push(`Ligne ${i + 1} : Valeur de production non définie`);
    }
  }

  // Envoi d’un email si incohérences détectées
  if (incoherences.length > 0) {
    const sujet = "🛑 Alerte : incohérences détectées dans Production_Agrements";
    const corps = "Bonjour Joseph,\n\nDes incohérences ont été détectées lors du traitement automatique :\n\n" +
                  incoherences.join("\n") +
                  "\n\nLignes concernées : " + lignesIncoherentes.join(", ") +
                  "\n\nMerci de vérifier les lignes indiquées.\n\n— Copilot";
    MailApp.sendEmail("jc.ahoulou@dgh.ci", sujet, corps);
  }

  // Coloration conditionnelle dans AQ de Traiter_Agrements
  const plageFusion = feuilleTraiter.getRange(2, COL_FUSION + 1, feuilleTraiter.getLastRow() - 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TRUE")
      .setFontColor("#008000") // Vert
      .setRanges([plageFusion])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("FALSE")
      .setFontColor("#FF0000") // Rouge
      .setRanges([plageFusion])
      .build()
  ];
  feuilleTraiter.setConditionalFormatRules(rules);
}

function preparerDeclenchementAutocrat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleProd = ss.getSheetByName("Production_Agrements_Autocrat2");
  const feuilleSuivi = ss.getSheetByName("Suivre_Agrements");

  const donneesProd = feuilleProd.getDataRange().getValues();
  const donneesSuivi = feuilleSuivi.getDataRange().getValues();

  // Index des colonnes (0-based)
  const COL_EMAIL = 0;
  const COL_MOTIF = 1;
  const COL_INITIALE = 2;
  const COL_NOM_SOCIETE = 3;
  const COL_FONCTION_RESP = 4;
  const COL_SOUS_PROFIL = 5;
  const COL_PROFIL = 6;
  const COL_CATEGORIE = 7;
  const COL_DATE_ENREG = 9;
  const COL_ADRESSE = 11;
  const COL_CONTACT = 12;
  const COL_RCCM = 13;
  const COL_NCC = 14;
  const COL_NBR_ANNEES = 15;
  const COL_EXEMPLE_ACTIVITE = 16;
  const COL_FRAIS_CHIFFRE = 17;
  const COL_FRAIS_LETTRE = 18;
  const COL_ACTIONNAIRE_1 = 19;
  const COL_NATIONALITE_1 = 20;
  const COL_PART_1 = 21;
  const COL_NRO_DATE_AGREMENT = 40;
  const COL_PRODUIRE_DOC = 41;
  const COL_FUSION = 42;
  const COL_SIGNAL = 55;
  const COL_DATE_TRAITEMENT = 56;
  const COL_SUIVI_DATE = 11; // L dans Suivre_Agrements
  const COL_DEBUT_LIENS = 43; // AR
  const COL_FIN_LIENS = 54;   // BC

  const estNonVide = (val) => val !== "" && val !== null;

  // Phase 1 : Déclenchement Autocrat
  for (let i = 1; i < donneesProd.length; i++) {
    const ligne = donneesProd[i];
    const fusion = ligne[COL_FUSION];
    const produire = ligne[COL_PRODUIRE_DOC];
    const motif = ligne[COL_MOTIF];

    if (fusion !== "FALSE" || produire !== "Oui") continue;

    if (
      motif === "Traitement Agrément" &&
      estNonVide(ligne[COL_EMAIL]) &&
      estNonVide(ligne[COL_INITIALE]) &&
      estNonVide(ligne[COL_NOM_SOCIETE]) &&
      estNonVide(ligne[COL_SOUS_PROFIL]) &&
      estNonVide(ligne[COL_PROFIL]) &&
      estNonVide(ligne[COL_CATEGORIE]) &&
      estNonVide(ligne[COL_DATE_ENREG]) &&
      estNonVide(ligne[COL_ADRESSE]) &&
      estNonVide(ligne[COL_CONTACT]) &&
      estNonVide(ligne[COL_RCCM]) &&
      estNonVide(ligne[COL_NCC]) &&
      estNonVide(ligne[COL_NBR_ANNEES]) &&
      estNonVide(ligne[COL_EXEMPLE_ACTIVITE]) &&
      estNonVide(ligne[COL_FRAIS_CHIFFRE]) &&
      estNonVide(ligne[COL_FRAIS_LETTRE]) &&
      estNonVide(ligne[COL_ACTIONNAIRE_1]) &&
      estNonVide(ligne[COL_NATIONALITE_1]) &&
      estNonVide(ligne[COL_PART_1])
    ) {
      feuilleProd.getRange(i + 1, COL_SIGNAL + 1).setValue("Lancer_Production");
      continue;
    }

    if (
      motif === "Accusé de Réception" &&
      estNonVide(ligne[COL_EMAIL]) &&
      estNonVide(ligne[COL_INITIALE]) &&
      estNonVide(ligne[COL_NOM_SOCIETE]) &&
      estNonVide(ligne[COL_FONCTION_RESP]) &&
      estNonVide(ligne[COL_SOUS_PROFIL]) &&
      estNonVide(ligne[COL_PROFIL]) &&
      estNonVide(ligne[COL_CATEGORIE]) &&
      estNonVide(ligne[COL_DATE_ENREG])
    ) {
      feuilleProd.getRange(i + 1, COL_SIGNAL + 1).setValue("Lancer_Accusé");
      continue;
    }

    if (
      motif === "Traitement Agrément" &&
      estNonVide(ligne[COL_EMAIL]) &&
      estNonVide(ligne[COL_INITIALE]) &&
      estNonVide(ligne[COL_NOM_SOCIETE]) &&
      estNonVide(ligne[COL_FONCTION_RESP]) &&
      estNonVide(ligne[COL_PROFIL]) &&
      estNonVide(ligne[COL_CATEGORIE]) &&
      estNonVide(ligne[COL_DATE_ENREG]) &&
      estNonVide(ligne[COL_EXEMPLE_ACTIVITE]) &&
      estNonVide(ligne[COL_NRO_DATE_AGREMENT])
    ) {
      feuilleProd.getRange(i + 1, COL_SIGNAL + 1).setValue("Lancer_MiseAJour");
    }
  }

  // Phase 2 : Journalisation conditionnelle
  for (let i = 1; i < donneesProd.length; i++) {
    const ligne = donneesProd[i];
    const fusion = ligne[COL_FUSION];
    const signal = ligne[COL_SIGNAL];
    const dateBE = ligne[COL_DATE_TRAITEMENT];
    const dateSuivi = donneesSuivi[i][COL_SUIVI_DATE];
    const plageLiens = ligne.slice(COL_DEBUT_LIENS, COL_FIN_LIENS + 1);
    const contientLien = plageLiens.some(cell => typeof cell === "string" && cell.includes("https://docs.google.com/document/d"));

    // Cas 1 : FUSION = TRUE + SIGNAL ≠ vide + BE vide → journaliser
    if (fusion === "TRUE" && signal !== "" && !estNonVide(dateBE)) {
      const dateTraitement = new Date();
      feuilleProd.getRange(i + 1, COL_SIGNAL + 1).clearContent();
      feuilleProd.getRange(i + 1, COL_DATE_TRAITEMENT + 1).setValue(dateTraitement);
      if (!estNonVide(dateSuivi)) {
        feuilleSuivi.getRange(i + 1, COL_SUIVI_DATE + 1).setValue(dateTraitement);
      }
      continue;
    }

    // Cas 2 : BE vide + lien détecté dans AR:BC → journaliser
    if (!estNonVide(dateBE) && contientLien) {
      const dateTraitement = new Date();
      feuilleProd.getRange(i + 1, COL_DATE_TRAITEMENT + 1).setValue(dateTraitement);
      if (!estNonVide(dateSuivi)) {
        feuilleSuivi.getRange(i + 1, COL_SUIVI_DATE + 1).setValue(dateTraitement);
      }
    }
  }
}

function postFusionAutocrat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleProd = ss.getSheetByName("Production_Agrements_Autocrat2");
  const feuilleSuivi = ss.getSheetByName("Suivre_Agrements");

  const donneesProd = feuilleProd.getDataRange().getValues();
  const donneesSuivi = feuilleSuivi.getDataRange().getValues();

  const COL_FUSION = 42;           // AQ
  const COL_SIGNAL = 55;           // BD
  const COL_DATE_TRAITEMENT = 56;  // BE
  const COL_SUIVI_DATE = 11;        // L dans Suivre_Agrements
  const COL_DEBUT_LIENS = 43;      // AR
  const COL_FIN_LIENS = 54;        // BC

  const estNonVide = (val) => val !== "" && val !== null;

  for (let i = 1; i < donneesProd.length; i++) {
    const ligne = donneesProd[i];
    const fusion = ligne[COL_FUSION];
    const signal = ligne[COL_SIGNAL];
    const dateBE = ligne[COL_DATE_TRAITEMENT];
    const dateSuivi = donneesSuivi[i][COL_SUIVI_DATE];
    const plageLiens = ligne.slice(COL_DEBUT_LIENS, COL_FIN_LIENS + 1);
    const contientLien = plageLiens.some(cell => typeof cell === "string" && cell.includes("https://docs.google.com/document/d"));

    // Cas 1 : FUSION = TRUE + SIGNAL ≠ vide + BE vide
    if (fusion === "TRUE" && estNonVide(signal) && !estNonVide(dateBE)) {
      const dateTraitement = new Date();
      feuilleProd.getRange(i + 1, COL_SIGNAL + 1).clearContent();
      feuilleProd.getRange(i + 1, COL_DATE_TRAITEMENT + 1).setValue(dateTraitement);
      if (!estNonVide(dateSuivi)) {
        feuilleSuivi.getRange(i + 1, COL_SUIVI_DATE + 1).setValue(dateTraitement);
      }
      continue;
    }

    // Cas 2 : BE vide + lien détecté dans AR:BC
    if (!estNonVide(dateBE) && contientLien) {
      const dateTraitement = new Date();
      feuilleProd.getRange(i + 1, COL_DATE_TRAITEMENT + 1).setValue(dateTraitement);
      if (!estNonVide(dateSuivi)) {
        feuilleSuivi.getRange(i + 1, COL_SUIVI_DATE + 1).setValue(dateTraitement);
      }
    }
  }
}

function synchroniserDatesTraitement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleProd = ss.getSheetByName("Production_Agrements_Autocrat2");
  const feuilleSuivi = ss.getSheetByName("Suivre_Agrements");

  const donneesProd = feuilleProd.getDataRange().getValues();
  const donneesSuivi = feuilleSuivi.getDataRange().getValues();

  const COL_DATE_TRAITEMENT = 56; // BE
  const COL_DEBUT_VERIF = 43;     // AR
  const COL_FIN_VERIF = 54;       // BC
  const COL_SUIVI_DATE = 11;       // L

  const estNonVide = (val) => val !== "" && val !== null;

  for (let i = 1; i < donneesProd.length; i++) {
    const ligneProd = donneesProd[i];
    const dateBE = ligneProd[COL_DATE_TRAITEMENT];
    const dateSuivi = donneesSuivi[i][COL_SUIVI_DATE];

    // Cas 1 : BE contient une date et L est vide → copier
    if (estNonVide(dateBE) && !estNonVide(dateSuivi)) {
      feuilleSuivi.getRange(i + 1, COL_SUIVI_DATE + 1).setValue(dateBE);
      continue;
    }

    // Cas 2 : BE vide + lien détecté dans AR à BC → inscrire date du jour
    if (!estNonVide(dateBE)) {
      const plageVerif = ligneProd.slice(COL_DEBUT_VERIF, COL_FIN_VERIF + 1);
      const lienDocPresent = plageVerif.some(cell =>
        typeof cell === "string" && cell.includes("https://docs.google.com/document/d")
      );

      if (lienDocPresent) {
        const dateTraitement = new Date();
        feuilleProd.getRange(i + 1, COL_DATE_TRAITEMENT + 1).setValue(dateTraitement);
        if (!estNonVide(dateSuivi)) {
          feuilleSuivi.getRange(i + 1, COL_SUIVI_DATE + 1).setValue(dateTraitement);
        }
      }
    }
  }
}

function insererTableauApresArticle2() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KADJA_Annexe 2");
  const donnees = feuille.getDataRange().getValues();

  // Extraire colonnes A à D avec en-têtes, corriger format de la colonne A
  const tableau = [];
  for (let i = 0; i < donnees.length; i++) {
    const ligne = donnees[i].slice(0, 4).map((val, idx) => {
      if (idx === 0 && typeof val === "number") {
        return val % 1 === 0 ? String(val.toFixed(0)) : String(val); // évite 1.0
      }
      return val;
    });
    if (ligne.join("") !== "") {
      tableau.push(ligne);
    }
  }

  // Sélectionner un document Google Docs
  const fichiers = DriveApp.getFilesByType(MimeType.GOOGLE_DOCS);
  const docs = [];
  while (fichiers.hasNext()) {
    const fichier = fichiers.next();
    docs.push({ id: fichier.getId(), nom: fichier.getName() });
  }

  const ui = SpreadsheetApp.getUi();
  let listeDocs = docs.map((doc, i) => `${i + 1}. ${doc.nom}`).join("\n");
  let reponse = ui.prompt(`Sélectionne le document Google Docs où insérer le tableau :\n\n${listeDocs}\n\nEntre le numéro correspondant :`);
  let indexChoisi = parseInt(reponse.getResponseText(), 10) - 1;

  if (isNaN(indexChoisi) || indexChoisi < 0 || indexChoisi >= docs.length) {
    ui.alert("Sélection invalide. Opération annulée.");
    return;
  }

  const doc = DocumentApp.openById(docs[indexChoisi].id);
  const body = doc.getBody();
  const texteCible = "Article 2";
  const found = body.findText(texteCible);

  if (!found) {
    ui.alert(`Le texte "${texteCible}" n'a pas été trouvé dans le document.`);
    return;
  }

  // Trouver le paragraphe contenant "Article 2"
  const element = found.getElement();
  const paragraphe = element.getParent();
  const index = body.getChildIndex(paragraphe);

  // Insérer un espace visuel
  body.insertParagraph(index + 1, "");

  // Insérer le tableau juste après
  const table = body.insertTable(index + 2, tableau);
  table.setBorderWidth(1);

  // Mise en forme : en-tête en gras, alignement centré, taille 10
  const nbLignes = table.getNumRows();
  const nbColonnes = table.getRow(0).getNumCells();

  for (let r = 0; r < nbLignes; r++) {
    const ligne = table.getRow(r);
    for (let c = 0; c < nbColonnes; c++) {
      const cellule = ligne.getCell(c);
      if (cellule.getNumChildren() > 0 && cellule.getChild(0).getType() === DocumentApp.ElementType.TEXT) {
        const texte = cellule.getChild(0).asText();
        texte.setFontSize(10).setTextAlignment(DocumentApp.TextAlignment.CENTER);
        if (r === 0) texte.setBold(true); // en-tête
      }
    }
  }

  ui.alert("✅ Le tableau a été inséré juste après 'Article 2' avec mise en forme.");
}

function sauvegarderVersCompteSecondaire() {
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceData = sourceSpreadsheet.getSheets().map(sheet => ({
    name: sheet.getName(),
    values: sheet.getDataRange().getValues()
  }));

  const backupSpreadsheet = SpreadsheetApp.create(`Backup - ${sourceSpreadsheet.getName()}`);
  
  sourceData.forEach(sheetData => {
    const sheet = backupSpreadsheet.insertSheet(sheetData.name);
    sheet.getRange(1, 1, sheetData.values.length, sheetData.values[0].length).setValues(sheetData.values);
  });

  backupSpreadsheet.deleteSheet(backupSpreadsheet.getSheets()[0]); // Supprime la feuille par défaut
  backupSpreadsheet.addEditor("ahoujochris@gmail.com"); // Partage avec le compte secondaire
}
