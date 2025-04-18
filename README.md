# README - Générateur de CV et Lettre de Motivation

Ce projet contient deux programmes Python qui permettent de générer facilement un CV et une lettre de motivation au format DOCX à partir de fichiers CSV.

## Contenu du projet

- cv.py - Programme pour générer un CV
- cover-letter.py - Programme pour générer une lettre de motivation
- docx_utils.py - Fonctions utilitaires pour la mise en forme des documents
- À fournir -> cv_data.csv - Exemple de fichier CSV pour le CV
- À fournir -> cover_letter.csv - Exemple de fichier CSV pour la lettre de motivation

## Prérequis

- Python 3.10 ou supérieur
- Le module `python-docx`

Pour installer les dépendances :

```bash
pip install python-docx
```

## Utilisation

Attention la mise en page n'est pas encore parfaite, il faudra encore certainement l'ajuster dans un logiciel de traitement de texte.
Le code est encore en phase de développement et peut nécessiter des ajustements pour répondre à vos besoins spécifiques.

### Générer un CV

1. Créez un fichier CSV avec les données de votre CV en suivant le format défini dans cv_data.csv.
2. Exécutez la commande suivante :

```bash
python cv.py chemin/vers/votre_cv.csv [debutant|accompli]
```

Le CV généré sera sauvegardé sous le nom `cv-test.docx` dans le répertoire courant.

### Générer une lettre de motivation

1. Créez un fichier CSV avec les données de votre lettre de motivation en suivant le format défini dans cover_letter.csv.
2. Exécutez la commande suivante :

```bash
python cover-letter.py chemin/vers/votre_lettre.csv
```

La lettre de motivation générée sera sauvegardée sous le nom `cover-letter-test.docx` dans le répertoire courant.

## Format des fichiers CSV

### Format du CV

Le fichier CSV pour le CV doit contenir les colonnes suivantes :

- `section` : type de section (title, personal, experience, education, skills)
- `title` : titre de l'élément
- `subtitle` : sous-titre ou complément d'information (optionnel)
- `description` : description détaillée
- `start_date` : date de début (optionnel)
- `end_date` : date de fin (optionnel)

### Format de la lettre de motivation

Le fichier CSV pour la lettre de motivation doit contenir les colonnes suivantes :

- `section` : type de section (généralement "cover_letter")
- `title` : titre de la section (Expéditeur, Destinataire, Objet, etc.)
- `content` : contenu de la section

## Personnalisation

Vous pouvez modifier les fichiers Python pour personnaliser la mise en forme des documents générés. Les principaux éléments de style (taille de police, couleurs, alignement) sont configurés dans les fonctions du fichier docx_utils.py.

## Exemples

Des exemples de fichiers CSV sont fournis avec le projet :

### Fichier CSV pour CV (cv_example.csv)

```csv
section,title,subtitle,description,start_date,end_date
title,Titre,,"Développeur Full Stack, recherche poste en CDI",,
personal,Nom,Prénom,"Jean Dupont",,
personal,Adresse,,"123 Rue des Développeurs, 75000 Paris",,
personal,Email,,"<jean.dupont@example.com>",,
personal,Tel,,"06 12 34 56 78",,
objectives,Debutant,,""Recherche experience significative en développement logiciel"",,
experience,Lead Developer,"TechSolutions, Paris","Développement d'applications web, gestion d'une équipe de 5 développeurs, mise en place de méthodologies agiles",2020,2025
experience,Développeur Backend,"WebInnovate","Conception et maintenance de microservices en Java et Spring Boot",2017,2020
experience,Développeur Junior,"StartupLab","Développement full stack sur diverses applications clients",2015,2017
education,Master Informatique,Université de Paris,"Spécialisation en développement logiciel",2013,2015
education,Licence Informatique,Université de Lyon,"Formation générale en informatique et algorithmique",2010,2013
skills,JavaScript,,Avancé,,
skills,React,,Avancé,,
skills,Node.js,,Intermédiaire,,
skills,Java,,Intermédiaire,,
skills,Python,,Intermédiaire,,
skills,Docker,,Intermédiaire,,
skills,AWS,,Débutant,,
skills,GraphQL,,Débutant,,
```

### Fichier CSV pour Lettre de Motivation (cover_letter_example.csv)

```csv
section,title,content
cover_letter,Date,"14 avril 2025"
cover_letter,Expéditeur,"Jean Dupont, 123 Rue des Développeurs, 75000 Paris"
cover_letter,Destinataire,"InnoTech Solutions, 45 Avenue de l'Innovation, 69000 Lyon"
cover_letter,Objet,"Candidature au poste de Développeur Full Stack"
cover_letter,Salutation,"Madame, Monsieur,"
cover_letter,Corps,"Je me permets de vous adresser ma candidature pour le poste de Développeur Full Stack que vous proposez au sein de votre entreprise. Diplômé d'un Master en Informatique et fort d'une expérience de plus de 7 ans dans le développement web, je pense disposer des compétences techniques et humaines nécessaires pour intégrer votre équipe. Au cours de mon parcours professionnel, j'ai pu développer une solide expertise dans les technologies front-end et back-end, notamment avec React, Node.js et Java. J'ai également acquis une expérience significative en gestion d'équipe et en méthodologies agiles lors de mon dernier poste chez TechSolutions. Particulièrement intéressé par les défis techniques que vous relevez dans le domaine de l'intelligence artificielle appliquée, je souhaite mettre mes compétences au service de votre entreprise reconnue pour son innovation. Je serais heureux de pouvoir échanger avec vous lors d'un entretien afin de vous présenter plus en détail mon parcours et ma motivation."
cover_letter,Formule de politesse,"Je vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées."
cover_letter,Signature,"Jean Dupont"
```

## Backlog

- [ ] Améliorer la mise en page du CV et de la lettre de motivation
- [ ] Ajouter des fonctionnalités pour personnaliser davantage le style (polices, couleurs, etc.)
