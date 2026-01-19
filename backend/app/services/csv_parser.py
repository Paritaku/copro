import csv
import re
from typing import List, Optional
from app.models.models import Lot, Floor, ImportedData


class CSVParser:
    """Parser pour fichiers CSV de titres fonciers"""
    
    def __init__(self, delimiter=";"):
        self.delimiter = delimiter
    
    def parse_file(self, file_path: str) -> ImportedData:
        """Parse un fichier CSV de titre foncier"""
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f, delimiter=self.delimiter)
            rows = [row for row in reader]
        
        return self._parse_rows(rows)
    
    def parse_content(self, content: str) -> ImportedData:
        """Parse le contenu CSV en string"""
        rows = []
        for line in content.split('\n'):
            rows.append(line.split(self.delimiter))
        
        return self._parse_rows(rows)
    
    def _parse_rows(self, rows: List[List[str]]) -> ImportedData:
        """Parse les lignes du CSV"""
        titre_foncier = ""
        etages: List[Floor] = []
        
        i = 0
        # Extraire le titre foncier
        while i < len(rows):
            row = rows[i]
            if len(row) > 1 and "Titre foncier" in row[1]:
                # Ex: "Modification successives du Titre foncier  :154311 /05"
                titre_foncier = row[1].split(":")[-1].strip()
                i += 1
                break
            i += 1
        
        # Sauter jusqu'aux données réelles
        while i < len(rows):
            row = rows[i]
            if len(row) > 1 and row[0] and "Propriété dite" in row[0]:
                    propriete_nom = row[1].strip()
                    i += 1
                    break
            i += 1
        
        # Parser les étages et lots
        while i < len(rows):
            row = rows[i]
            
            # Détection d'un nouvel étage
            if len(row) > 0 and row[0] and ":" in row[0] and any(c.isalpha() for c in row[0]):
                parts = row[0].split(":")
                etage_name = parts[0].strip()  # "Rez-de-chaussée"
                cotes = parts[1].strip() if len(parts) > 1 else ""  # "Des côtes +0.10m et +1,10m à la côte 4,10m"
                    
                lots: List[Lot] = []
                i += 1
                
                # Parser les lots de cet étage
                total_surf_int = 0
                total_surf_surp = 0
                
                while i < len(rows):
                    row = rows[i]
                    
                    # Arrêt si nouvel étage
                    if len(row) > 0 and row[0] and ":" in row[0] and any(c.isalpha() for c in row[0]):
                        break
                    
                    # Arrêt si fin des données
                    if not row or not any(cell.strip() for cell in row):
                        i += 1
                        continue
                    
                    # Total row
                    if len(row) > 1 and "Total" in row[1]:
                        i += 1
                        break
                    
                    # Parser un lot
                    if len(row) > 3 and row[0] and row[4]:  # Propriété et Surface interieure du titre
                        lot = self._parse_lot(row)
                        if lot:
                            lots.append(lot)
                            total_surf_int += lot.surface_interieure or 0
                            total_surf_surp += lot.surface_avec_surplomb or 0
                    
                    i += 1
                
                if lots:
                    floor = Floor(
                        nom=etage_name,
                        cotes=cotes,
                        lots=lots,
                        total_surface_interieure=total_surf_int,
                        total_surface_avec_surplomb=total_surf_surp
                    )
                    etages.append(floor)
            else:
                i += 1
        
        return ImportedData(
            titre_foncier=titre_foncier,
            propriete_nom=propriete_nom,
            etages=etages
        )
    
    def _parse_lot(self, row: List[str]) -> Optional[Lot]:
        """Parse une ligne représentant un lot"""
        try:
            propriete = row[0].strip() if len(row) > 0 else ""
            titre_num = row[1].strip() if len(row) > 1 else ""

            indice_privative = row[3].strip() if len(row) > 3 else None
            indice_commune = row[4].strip() if len(row) > 4 else None

            surface_interieure = self._parse_float(row[5]) if len(row) > 5 else 0
            surface_avec_surplomb = self._parse_float(row[6]) if len(row) > 6 else 0
            
            consistance = row[7].strip() if len(row) > 7 else ""
            observations = row[8].strip() if len(row) > 8 else None
            
            # Valider que c'est un lot valide (au moins propriete et indice)
            if propriete and (indice_privative or indice_commune):
                return Lot(
                    propriete=propriete,
                    titre_num=titre_num,
                    indice_privative=surface_privative,
                    indice_commune=surface_commune,
                    surface_interieure=surface_interieure,
                    surface_avec_surplomb=surface_avec_surplomb,
                    consistance=consistance,
                    observations=observations
                )
        except Exception as e:
            print(f"Erreur parsing lot: {e}")
        
        return None
    
    def _parse_float(self, value: str) -> Optional[float]:
        """Parse une valeur float en gérant les formats différents"""
        if not value or not value.strip():
            return None
        value = value.strip()
        try:
            return float(value)
        except ValueError:
            return None
