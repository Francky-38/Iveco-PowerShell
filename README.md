# Iveco PowerShell - Extraction de Références

Un projet PowerShell complet pour extraire et rechercher des références dans des archives PPTX organisées en arborescence.

## 🎯 Fonctionnalités

✅ **Extraction de Références**
- Parcourt automatiquement l'arborescence des affaires
- Cherche les fichiers PPTX dans les dossiers structurés (`Ligne_EG0`)
- Extrait les références (format: `[TRS]?\d{5,10}`) directement du texte
- Exporte les résultats en XML structuré

✅ **Interface de Recherche Graphique**
- Interface WinForms intuitive et professionnelle
- Recherche rapide par référence
- Affichage en tableau avec colonnes: Référence, Marché, Poste, SOP, Page
- Support des caractères accentués

✅ **Configuration Centralisée**
- Fichier de configuration unique (`config.ps1`)
- Chemins facilement modifiables
- Support de paramètres personnalisés en ligne de commande

## 📁 Structure du Projet

```
PowerShell/
├── Configs/
│   └── config.ps1                 # Configuration centralisée
├── Functions/
│   └── Helper.ps1                 # Toutes les fonctions réutilisables
├── Scripts/
│   ├── ExtracRefServeur.ps1      # Script d'extraction
│   └── SearchGui-References.ps1   # Interface de recherche GUI
├── Tests/
│   └── Test-Helper.ps1            # Tests (à développer)
├── setup.ps1                       # Script d'initialisation
├── requirements.txt                # Dépendances
└── README.md                       # Ce fichier
```

## 🚀 Installation

### Prérequis
- PowerShell 5.1 ou supérieur
- Windows (pour l'interface WinForms)

### Configuration

1. **Cloner le projet**
```powershell
git clone https://github.com/[votre-username]/Iveco-PowerShell.git
cd PowerShell
```

2. **Initialiser le projet**
```powershell
.\setup.ps1
```

3. **Configurer les chemins** (optionnel)
Éditez `Configs/config.ps1` :
```powershell
$Config.ExtractionRootPath = "D:\W\Iveco\serveur"      # Chemin racine
$Config.ExtractXmlData = "D:\W\Iveco\RefServeur.xml"   # Fichier XML de sortie
```

## 📖 Utilisation

### 1. Extraire les Références

```powershell
.\Scripts\ExtracRefServeur.ps1
```

**Options :**
```powershell
# Avec chemins personnalisés
.\Scripts\ExtracRefServeur.ps1 -RootPath "D:\autre\chemin" -OutputFile "D:\sortie.xml"
```

**Résultat :** Crée un fichier XML avec toutes les références trouvées

### 2. Rechercher les Références

```powershell
.\Scripts\SearchGui-References.ps1
```

**Options :**
```powershell
# Avec fichier XML personnalisé
.\Scripts\SearchGui-References.ps1 -XmlPath "D:\mon_fichier.xml"
```

**Interface :**
- Entrez une référence (ex: `T123456`)
- Cliquez sur "Rechercher" ou appuyez sur Entrée
- Les résultats s'affichent dans le tableau

## 📊 Format des Données

### Arborescence attendue
```
D:\W\Iveco\serveur\
├── AffaireSA/
│   └── 01-Dossiers ligne EL-EG\LIGNE EG0\
│       ├── Poste1\
│       │   └── document.pptx
│       └── Poste2\
│           └── guide.pptx
└── AffaireSB/
    └── 01-Dossiers ligne EL-EG\LIGNE EG0\
        └── Poste1\
            └── manuel.pptx
```

### Format des Références
- **Format valide :** `T123456`, `R1234`, `S12345678`
- **Format :** `[TRS]?\d{5,10}` (5 à 10 chiffres, avec optionnel préfixe T/R/S)
- **Extraction :** Les références sont recherchées dans le texte complet (pas seulement les cellules isolées)

### Fichier XML de Sortie
```xml
<References>
  <Entree>
    <Affaire>AffaireSA</Affaire>
    <Poste>Poste1</Poste>
    <SOP>document.pptx</SOP>
    <Page>slide1.xml</Page>
    <Reference>T123456</Reference>
  </Entree>
  ...
</References>
```

## 🔧 Fonctions Disponibles

### Dans `Helper.ps1`

**`Export-PptxReferencesFromTree`**
- Extrait les références depuis une arborescence complète
- Paramètres: `RootPath`, `OutputFile`

**`Show-SearchGui`**
- Interface graphique WinForms de recherche
- Paramètre: `XmlPath`

**`Show-SearchMenu`**
- Menu interactif console de recherche
- Paramètre: `XmlPath`

**`Get-WelcomeMessage`**
- Affiche un message de bienvenue personnalisé

**`Get-SystemInfo`**
- Affiche les informations système

## 📝 Configuration

Fichier: `Configs/config.ps1`

```powershell
# Paramètres d'extraction et recherche des references
$Config.ExtractionRootPath = "D:\W\Iveco\serveur"
$Config.ExtractXmlData = "D:\W\Iveco\RefServeur.xml"

# Paramètres globaux
$Config.Environment = "Development"
$Config.LogLevel = "Info"
$Config.LogFile = ".\Logs\project.log"
```

## 🐛 Dépannage

### Le script bloque lors de l'exécution depuis l'explorateur
**Solution :** Vérifiez la politique d'exécution
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Caractères accentués non affichés
**Solution :** Les caractères sont encodés avec `[char]233` (é). Assurez-vous que votre terminal supporte UTF-8.

### Aucun fichier PPTX trouvé
**Solution :** Vérifiez que :
- Les dossiers d'affaires se terminent par `SA` ou `SB`
- Le chemin `01-Dossiers ligne EL-EG\LIGNE EG0` existe
- Les fichiers `.pptx` sont directement dans les dossiers de postes

## 📋 Roadmap

- [ ] Tests automatisés
- [ ] Support des autres formats (DOCX, etc.)
- [ ] Export en CSV/JSON
- [ ] Historique des recherches
- [ ] Mode batch avec rapports

## 👨‍💻 Auteur

Créé pour le projet Iveco

## 📄 Licence

MIT License

## 🤝 Contribution

Les contributions sont bienvenues ! N'hésitez pas à :
- Signaler des bugs
- Proposer des améliorations
- Soumettre des pull requests

## 📞 Support

Pour toute question ou problème, veuillez ouvrir une issue sur GitHub.

---

**Version:** 1.0.0  
**Date:** 2026-01-29  
**Status:** ✅ Stable
"# Iveco-PowerShell" 
