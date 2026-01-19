from pydantic import BaseModel
from typing import Optional, List

class Lot(BaseModel):
    """Représente un lot dans un étage"""
    propriete: str  # Propriété dite
    titre_num: str  # Titre N°
    indice_privative: str
    indice_commune: str
    surface_interieure: float  # Surface intérieure du titre
    surface_avec_surplomb: float  # Surface avec surplomb
    consistance: str  # Type (Appartement, Local Commercial, etc.)
    observations: Optional[str] = None


class Floor(BaseModel):
    """Représente un étage"""
    nom: str  # Ex: "Rez-de-chaussée", "Premier Etage"
    cotes: str  # Ex: "De la cote 4,30m à la cote 7,30m"
    lots: List[Lot]
    total_surface_interieure: Optional[float] = None
    total_surface_avec_surplomb: Optional[float] = None


class ImportedData(BaseModel):
    """Structure complète des données importées"""
    titre_foncier: str  # Ex: "154311 /05"
    etages: List[Floor]
