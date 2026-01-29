# Changelog

Tous les changements notables de ce projet seront documentés dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
et ce projet respecte [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-01-29

### ✨ Ajout
- **Extraction de références** : Parcourt l'arborescence complète et extrait les références des fichiers PPTX
- **Interface GUI WinForms** : Interface graphique professionnelle pour rechercher les références
- **Menu interactif** : Menu console pour rechercher avec plusieurs critères (référence, affaire, poste)
- **Configuration centralisée** : Fichier `config.ps1` pour gérer les chemins et paramètres
- **Support des références imbriquées** : Extrait les références même si elles sont au milieu d'une chaîne de texte
- **Gestion des namespaces XML** : Traitement correct des fichiers XML PPTX avec namespaces
- **Documentation complète** : README détaillé avec exemples et dépannage

### 🔧 Fonctionnalités
- Recherche par référence (format: `[TRS]?\d{5,10}`)
- Affichage structuré avec colonnes: Référence, Marché, Poste, SOP, Page
- Support des caractères accentués français
- Export en XML structure
- Interface responsive avec DataGrid

### 🚀 Performance
- Itération directe sur les fichiers trouvés (pas de tableau intermédiaire)
- Gestion efficace des ressources temporaires
- Support des archives PPTX volumineuses

### 📦 Structure
- `setup.ps1` : Initialisation du projet
- `Configs/config.ps1` : Configuration centralisée
- `Functions/Helper.ps1` : Toutes les fonctions réutilisables
- `Scripts/ExtracRefServeur.ps1` : Script d'extraction
- `Scripts/SearchGui-References.ps1` : Interface de recherche

### 🛠️ Améliorations futures
- Tests automatisés
- Support CSV/JSON
- Historique des recherches
- Mode batch avec rapports

---

## Notes de version

### Version 1.0.0
**Statut:** ✅ Stable et prêt pour production

**Améliorations par rapport aux versions de développement:**
- Code optimisé et nettoyé
- Documentation complète
- Configuration externalisée
- Tests fonctionnels validés
