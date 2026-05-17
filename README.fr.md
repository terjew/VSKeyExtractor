# VSKeyExtractor

VSKeyExtractor est un petit outil en ligne de commande qui tente d’extraire la clé de licence utilisée pour activer une installation locale de Visual Studio.

Il parcourt les entrées connues des produits Visual Studio, lit les données de licence chiffrées dans le registre Windows, les déchiffre dans le contexte de l’utilisateur actuel et recherche une clé au format `AAAAA-BBBBB-CCCCC-DDDDD-EEEEE`.

## Fonctionnalités

- Lit les données de licence Visual Studio depuis le registre Windows
- Prend en charge plusieurs générations de produits Visual Studio
- Tente de déchiffrer les blobs de licence protégés avec le profil de l’utilisateur actuel
- Affiche dans la console toute valeur ressemblant à une clé

## Prérequis

- Windows
- Une installation de Visual Studio avec des données de licence locales présentes
- Le runtime .NET requis pour le projet que vous souhaitez exécuter

## Fonctionnement

L’outil vérifie les valeurs du registre aux emplacements de licence connus de Visual Studio, puis tente de déchiffrer les données stockées. Si une clé correspondante est trouvée dans le texte déchiffré, elle est écrite dans la sortie standard.

Si aucune clé n’est trouvée, l’outil passe simplement à l’entrée de produit suivante.

## Exécuter l’application

Vous pouvez exécuter le projet directement avec `dotnet run`.

### Cible .NET Framework / Windows

```powershell
dotnet run --project .\VSKeyExtractor.csproj
```

### Projet de style SDK

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj
```

### Exécuter un framework cible spécifique

Si vous utilisez le projet de style SDK et souhaitez cibler un framework particulier, vous pouvez utiliser :

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj -f net8.0-windows
```

D’autres frameworks cibles pris en charge peuvent inclure `net9.0-windows` ou `net10.0-windows`, selon le SDK .NET installé.

## Exemple de sortie

```text
Found key for Visual Studio 2022 Professional: ABCDE-FGHIJ-KLMNO-PQRST-UVWXY
```

## Remarques

- Exécutez l’outil sur la machine où Visual Studio a été activé.
- L’outil ne peut lire que les données disponibles pour le profil utilisateur actuel.
- Certaines entrées peuvent être indisponibles, endommagées ou protégées ; c’est normal.

## Traductions

- [English](README.md)
- [Deutsch](README.de.md)
- [Español](README.es.md)
