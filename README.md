# scan_drives_partages![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

# 📂 Audit & cartographie des Drives Partagés (Shared Drives)

## Description
Ce projet est une solution complète d'automatisation Google Apps Script. Il permet de générer un audit exhaustif de tous les **Drives Partagés** (Shared Drives) accessibles par l'utilisateur.

Le script ne se contente pas de lister les Drives, il crée une cartographie interactive en générant un onglet dédié pour chaque Drive (listant l'arborescence racine) et récupère les métadonnées de dernière activité (qui a modifié quoi et quand).

## ✨ Fonctionnalités clés

* **Importation dynamique** : Récupère la liste complète des Drives Partagés via l'API Drive.
* **Cartographie profonde** : Génère automatiquement un onglet (Feuille) pour chaque Drive Partagé contenant la liste de ses dossiers racines.
* **Analyse d'activité** : Identifie le dernier utilisateur actif ("Modifié par"), la date de modification et le fichier concerné.
* **Navigation intuitive** : Crée des liens hypertextes bidirectionnels (Index ↔ Onglets Drives) formatés pour les tableurs en locale Française (séparateur `;`).
* **Résilience (Anti-Timeout)** : Intègre un mécanisme de sauvegarde par lots (`BATCH_SIZE`) et un système de déclencheur automatique (Trigger) pour contourner la limite des 6 minutes d'exécution de Google.
* **Formatage FR** : Gestion native des dates (dd/MM/yyyy) et tri alphabétique respectant les accents.

## ⚙️ Prérequis et installation

### 1. Création du script
1. Ouvrez un nouveau Google Sheet.
2. Allez dans `Extensions` > `Apps Script`.
3. Copiez le code fourni dans le fichier `Code.js`.

### 2. Activation du service avancé (CRITIQUE)
Ce script utilise l'API Drive avancée (`Drive.Drives.list`). Vous devez l'activer manuellement :
1. Dans l'éditeur Apps Script, à gauche, cliquez sur le **+** à côté de **Services**.
2. Sélectionnez **Drive API**.
3. Cliquez sur **Ajouter**.
   * *Note : L'identifiant doit être `Drive` (par défaut).*

### 3. Premier lancement
1. Rechargez votre Google Sheet.
2. Un menu personnalisé `⚙️ Scanner les Drives partagés` apparaîtra après quelques secondes.

## 🚀 Utilisation

Le processus se déroule en deux étapes via le menu dédié :

### Étape 1 : Importer mes Drives partagés
* **Action** : Vide le tableau actuel et interroge l'API pour lister tous les ID et Noms des Drives.
* **Résultat** : Remplit les colonnes A (ID) et D (Nom).

### Étape 2 : Lancer l'audit
* **Action** : Parcourt la liste ligne par ligne.
    * Crée/Met à jour l'onglet enfant du Drive.
    * Génère le lien de navigation.
    * Cherche la dernière activité (User/Date).
* **Comportement** : Si le script approche de la limite de temps (4.5 min), il se met en pause, sauvegarde l'état et programme une reprise automatique après 1 minute.
* **Suivi** : Des messages "Toast" en bas à droite vous informent de la progression.

## 🛠 Configuration (Optionnelle)

Vous pouvez ajuster les constantes au début du fichier `Code.js` selon vos besoins :

```javascript
const CONFIG = {
  NOM_ONGLET_PRINCIPAL: "Liste_Drives_Partages", // Nom de l'onglet index
  TEMPS_MAX_EXECUTION: 1000 * 60 * 4.5,          // Seuil de déclenchement du trigger (ms)
  BATCH_SIZE: 10                                 // Fréquence de sauvegarde intermédiaire
};
