from fastapi import APIRouter, UploadFile, File, HTTPException
import tempfile
import os
from app.services.csv_parser import CSVParser
from app.models.models import ImportedData
from fastapi.responses import StreamingResponse

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
            "nb_etages": len(current_data.etages),
            "etages" : current_data.etages
        }
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate-xslx-voix")
async def generate_xslx_voix():
    """Génère un fichier XLSX pour les voix"""
    global current_data
    
    if current_data is None:
        raise HTTPException(status_code=400, detail="Aucune donnée. Uploadez d'abord un fichier CSV.")
    
    try:
        parser = CSVParser()
        file_stream = parser.generer_xlxs_voix(current_data)  
        
        headers = {
            "Content-Disposition": 'attachment; filename="Voix.xlsx"'
        }
        
        return StreamingResponse(
            file_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers = headers)  
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    
@router.post("/generate-xslx-quot")
async def generate_xslx_voix():
    """Génère un fichier XLSX pour les Quot P CH2"""
    global current_data
    
    if current_data is None:
        raise HTTPException(status_code=400, detail="Aucune donnée. Uploadez d'abord un fichier CSV.")
    
    try:
        parser = CSVParser()
        file_stream = parser.generer_xlxs_quotation(current_data)  
        
        headers = {
            "Content-Disposition": 'attachment; filename="Quot_P_CH2.xlsx"'
        }
        
        return StreamingResponse(
            file_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers = headers)  
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate-xslx-ta")
async def generate_xslx_ta():
    """Génère un fichier XLSX pour le tableau A des contenances"""
    global current_data
    
    if current_data is None:
        raise HTTPException(status_code=400, detail="Aucune donnée. Uploadez d'abord un fichier CSV.")

    try:
        parser = CSVParser()
        file_stream = parser.generer_xlxs_ta(current_data)
        
        headers = {
            "Content-Disposition": 'attachment; filename="TA.xlsx"'
        }
        
        return StreamingResponse(
            file_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers = headers)  
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate-xslx-tr-n")
async def generate_xslx_tn_r():
    """Génère un fichier XLSX pour le tableau TR-N des contenances"""
    global current_data
    
    if current_data is None:
        raise HTTPException(status_code=400, detail="Aucune donnée. Uploadez d'abord un fichier CSV.")

    try:
        parser = CSVParser()
        file_stream = parser.generer_excel_tr_n(current_data)
        
        headers = {
            "Content-Disposition": 'attachment; filename="TA.xlsx"'
        }
        
        return StreamingResponse(
            file_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers = headers)  
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
@router.get("/data")
def get_data():
    """Retourne les données parsées"""
    global current_data
    
    if current_data is None:
        raise HTTPException(status_code=400, detail="Aucune donnée. Uploadez d'abord un fichier CSV.")
    
    return current_data