# organisation-invitation

Petite app **HTML / CSS / JS** pour chercher un nom et afficher **TABLE X — NOM**.

## Démo en ligne

https://mande-dev.github.io/organisation-invitation/

## Lancer en local

**Double-cliquez sur `index.html`** — aucun serveur local nécessaire.

(Internet utile au premier chargement pour SheetJS via CDN.)

## Utilisation

1. Cliquez sur **Choisir le fichier Excel** (une seule fois).
2. Sélectionnez `LISTE DES INVITE.xlsx` (ou un autre plan de tables au même format).
3. Attendez « Prêt ».
4. Tapez un nom (partiel OK, casse / accents ignorés).
5. Résultat : `TABLE 1 — JOYCE`.
6. Cliquez sur un résultat pour voir **toute la table** (place + nom) ; recliquez ou « Fermer » pour masquer.

Après un **rafraîchissement** de la page, le même fichier est rechargé automatiquement. Utilisez **Changer le fichier Excel** seulement pour en prendre un autre.

## Données

- Le fichier Excel est la seule base. Il est **mémorisé dans le navigateur** (IndexedDB) après le premier choix.
- Pour mettre à jour les données après modification de l’Excel : **Changer le fichier Excel** et resélectionner le fichier.
- Format attendu : en-têtes `TABLE 1`, `TABLE 2`, … avec place + nom en dessous.
- Le fichier Excel n’est **pas** dans le dépôt (données personnelles).
