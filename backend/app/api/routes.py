from fastapi import APIRouter, UploadFile, File, HTTPException
import tempfile
import os
from app.services.csv_parser import CSVParser
from app.models import ImportedData

router = APIRouter(prefix="/api", tags=["data"])

# Stockage en mémoire
current_data: ImportedData = None


@router.post("/upload")
async def upload_csv(file: UploadFile = File(...)):
    """Upload et parse un fichier CSV"""
    global current_data
    
    if not file.filename.endswith('.csv'):
        raise HTTPException(status_code=400, detail="Le fichier doit être un CSV")
    
    try:
        # Sauvegarder temporairement le fichier
        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name
        
        # Parser le fichier
        parser = CSVParser()
        current_data = parser.parse_file(tmp_path)
        
        # Nettoyer
        os.unlink(tmp_path)
        
        return {
            "success": True,
            "titre_foncier": current_data.titre_foncier,
            "propriete_nom": current_data.propriete_nom,
            "nb_etages": len(current_data.etages)
        }
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/data")
def get_data():
    """Retourne les données parsées"""
    global current_data
    
    if current_data is None:
        raise HTTPException(status_code=400, detail="Aucune donnée. Uploadez d'abord un fichier CSV.")
    
    return current_data