# ExtractionPS :
  - [DONE] Ajouter la valeur de l'attribut <IsGood> pour chaque ligne d'extraction (indique si valeur correcte ou digital state de la table system)
  - Flexibilité dans le format de saisie de la date en paramètre (demande Ewen)
  - Développement d'une IHM pour ce script
  - Extension Visual Studio ?
  - [DONE] Paramètre optionnel pour préciser un dossier d'output & log
  - Paramètre pour lancer ExtractionPS en plusieurs instances parallèles avec nombre d'instance à choisir
  
# InsertionPS :
  -  [DONE]Gérer l'insertion des digitals state grace à l'attribut <IsGood> (différencier l'insertion d'un DS system et d'un DS good)
  -  [DONE]Différencier insertion digital state error (3 chiffres) d'un digital state standard (1 à 3 chiffres)
  - Gérer l'insertion en UTC

# ConfigurationPS :
  -  [DONE] Ne plus prendre les paramètres obsoletes des tags (anciens ptsecurity/datasecurity etc ...)
  - Paramètre optionnel pour préciser un dossier d'output & log

# DeletionPS :
  - Paramètre pour choisir la suppression d'un tag, de la donnée ou les deux (actuellement, fait les deux)
  - Paramètres date début/date fin optionnels pour réduire la période de suppression de données (actuellement, supprime toutes les données depuis 1/1/2000)

# CreationPS :
  - Utiliser la fonction import-csv avec nommage des colonnes (actuellement découpage à la main)
