# GesADSCOL
Script de gestion d'un annuaire Active Directory dans un environnement scolaire multi-niveau.

Fonctionnalités: 

  - Ajout d'un élève dans la base selon sa classe (redirection automatique vers l'OU adapté)
  - Ajout d'un élève en masse via csv (Utilisation d'un SIS pour l'extraction personnalisée, format décris plus bas)
  - Ajout d'un adulte dans la base selon sa fonction (redirection automatique vers l'OU adapté)
  - Ajout de plusieurs adultes en masse via csv (Utilisation d'un SIS pour l'extraction personalisée, format décris plus bas)
  - Recherche d'une personne selon son SAMAccountName (nom d'utilisateur)
  - Affichage d'informations plus précises via saisi du SAMAccountName (notamment l'emplacement de la personne dans l'AD et la description de son profil)
  - Création d'un mot de passe aléatoire via une fonction appelé pendant l'exécution du script (non obligatoire et personnalisable)
  - Gestion des années scolaire par analyse de la description du compte (fonction au plus basique, n'est pas exploitée à son plein potentiel)
  - Gestion de la montée des classes chaque année (fonction encore en test)
  - Suppresion d'une personne selon son SAMAccountName
  - Suppresion de plusieurs personnes par csv (Utilisation d'un SIS pour l'extraction personnalisée, format décris plus bas, fonction en cours de test)
  - Affichage des erreurs (verbeux)
  - Génération d'un fichier de log à la fin du traitement de la plupart des fonctions
  - Retour en arrière possible lors de la montée des classes dans l'AD (fonction encore en test)
  - Réorganisation des groupes élèves automatique (selon la description apposé sur chaque élève)
  - Changement du mot de passe d'une personne selon son SAMAccountName
  - Script contenant beaucoup de commentaires pour aiguiller l'utilisateur

Informations sur l'origine du script : 

Script original coécrit par M. PERROT, enseignant de technologie et M. COSTENOBLE, administrateur systèmes et réseaux.
Script retouché par M. COSTENOBLE pour l'implémentation de nouvelles fonctionnalités
Script réorganisé, réecrit plus "proprement" et raccourci par l'usage de fonction par l'IA (Gemini par Google et ChatGPT par OpenAI)
Utilisation de la documentation Microsoft pour les commandes de gestion d'un AD par Powershell
Fonction d'ajout interactif inspiré par le forum LazyAdmin (https://lazyadmin.nl/category/powershell/)
Vérification des prérequis pour l'exécution du script inspiré par Brian O’Connell via son blog (https://lifeofbrianoc.com/2015/02/11/script-create-ad-ousgroupsusers/)

Format d'import en CSV pour les différentes fonctions : 
élève : Surname; GivenName; CodeClasse (ex: Jean; DUPONT; 4E)
personnel : GivenName; Surname; Role (ex : François; XAVIER; Extérieur)
a supprimer : Nom; Prenom (à retravailler pour gérer les homonymes d'où les termes français à la place des termes anglophone plus standard)

Description non finalisée. Encore quelques fonctionnalités à décrire et quelques instructions à ajouter pour faciliter l'utilisation du script.
