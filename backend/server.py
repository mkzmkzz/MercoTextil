from fastapi import FastAPI, APIRouter, HTTPException, Depends
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field
from typing import List, Optional
import uuid
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import jwt
import hashlib
from passlib.context import CryptContext

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# MongoDB connection
mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

# Create the main app without a prefix
app = FastAPI()

# Create a router with the /api prefix
api_router = APIRouter(prefix="/api")

# Security
security = HTTPBearer()
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
JWT_SECRET = os.getenv("JWT_SECRET", "fusosmanager_secret_key_2024")

# Brazil timezone usando zoneinfo (padrão Python 3.9+)
BRAZIL_TZ = ZoneInfo("America/Sao_Paulo")

def get_utc_now():
    """Get current UTC time - MongoDB best practice"""
    return datetime.now(timezone.utc)

def get_brazil_time():
    """Get current time in Brazil timezone for display purposes"""
    utc_now = datetime.now(timezone.utc)
    return utc_now.astimezone(BRAZIL_TZ)

def convert_utc_to_brazil(utc_datetime):
    """Convert UTC datetime to Brazil timezone"""
    if utc_datetime is None:
        return None
    if isinstance(utc_datetime, str):
        utc_datetime = datetime.fromisoformat(utc_datetime.replace('Z', '+00:00'))
    if utc_datetime.tzinfo is None:
        utc_datetime = utc_datetime.replace(tzinfo=timezone.utc)
    return utc_datetime.astimezone(BRAZIL_TZ)

# Helper function to convert MongoDB documents
def serialize_doc(doc):
    """Convert MongoDB document to JSON serializable format"""
    if doc is None:
        return None
    
    # Convert ObjectId to string
    if '_id' in doc:
        doc['_id'] = str(doc['_id'])
    
    # Convert datetime objects to ISO format
    for key, value in doc.items():
        if isinstance(value, datetime):
            doc[key] = value.isoformat()
    
    return doc

def serialize_docs(docs):
    """Convert list of MongoDB documents to JSON serializable format"""
    if not docs:
        return []
    return [serialize_doc(doc) for doc in docs]

# Models
class User(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    username: str
    email: str
    role: str  # admin, operador_interno, operador_externo
    active: bool = True
    # Permissões por aba
    permissions: dict = Field(default_factory=lambda: {
        "dashboard": True,
        "producao": True,
        "ordem_producao": True,
        "relatorios": True,
        "espulagem": True,
        "manutencao": True,
        "banco_dados": True,
        "administracao": False
    })
    created_at: datetime = Field(default_factory=get_utc_now)

class UserCreate(BaseModel):
    username: str
    email: str
    password: str
    role: str
    permissions: Optional[dict] = None

class UserUpdate(BaseModel):
    username: Optional[str] = None
    email: Optional[str] = None
    password: Optional[str] = None
    role: Optional[str] = None
    permissions: Optional[dict] = None
    active: Optional[bool] = None

class UserLogin(BaseModel):
    username: str
    password: str

class LoginResponse(BaseModel):
    token: str
    user: User

class Machine(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    code: str  # CD1, CD2, CI1, F1, etc.
    position: str  # unique position identifier
    status: str = "verde"  # verde, amarelo, vermelho, azul (manutenção), desativada
    layout_type: str  # 16_fusos or 32_fusos
    active: bool = True  # True = ativa, False = desativada
    updated_at: datetime = Field(default_factory=get_utc_now)

class Maintenance(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    machine_id: str
    machine_code: str
    motivo: str
    status: str = "em_manutencao"  # em_manutencao, finalizada
    created_by: str
    created_at: datetime = Field(default_factory=get_utc_now)
    finished_at: Optional[datetime] = None
    finished_by: Optional[str] = None

class MaintenanceCreate(BaseModel):
    machine_id: str
    motivo: str

class Order(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    machine_id: str
    machine_code: str
    layout_type: str
    cliente: str
    artigo: str
    cor: str
    quantidade: str
    observacao: str = ""
    status: str = "pendente"  # pendente, em_producao, finalizado
    created_by: str
    created_at: datetime = Field(default_factory=get_utc_now)
    started_at: Optional[datetime] = None
    finished_at: Optional[datetime] = None
    observacao_liberacao: str = ""
    laudo_final: str = ""
    # New fields for machine queue management
    espulagem_id: Optional[str] = None  # Link to espulagem if created from espulagem
    ordem_producao_id: Optional[str] = None  # Link to ordem de producao
    numero_os: Optional[str] = None  # OS number
    origem: str = "manual"  # manual, espulagem, ordem
    queue_position: int = 0  # Position in machine queue

class OrderCreate(BaseModel):
    machine_id: str
    cliente: str
    artigo: str
    cor: str
    quantidade: str
    observacao: str = ""

class OrderUpdate(BaseModel):
    status: str
    observacao_liberacao: str = ""
    laudo_final: str = ""

class OrdemProducao(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    numero_os: str  # Sequential number: 0001, 0002, etc.
    cliente: str
    artigo: str
    cor: str
    metragem: str
    data_entrega: str  # ISO date string
    observacao: str = ""
    # Novos campos do banco de dados
    engrenagem: str = ""
    fios: str = ""
    maquinas: str = ""
    status: str = "pendente"  # pendente, em_producao, finalizado
    criado_em: datetime = Field(default_factory=get_utc_now)
    iniciado_em: Optional[datetime] = None
    finalizado_em: Optional[datetime] = None
    criado_por: str
    created_at: datetime = Field(default_factory=get_utc_now)
    updated_at: datetime = Field(default_factory=get_utc_now)
    # Dados temporários para salvamento parcial no relatórios
    dados_temporarios_maquinas: list = Field(default_factory=list)  # Machine allocations temporários
    espula_data_temp: dict = Field(default_factory=dict)  # Dados do formulário de espulagem temporários
    editado_por: Optional[str] = None
    editado_em: Optional[datetime] = None

class OrdemProducaoCreate(BaseModel):
    cliente: str
    artigo: str
    cor: str
    metragem: str
    data_entrega: str
    observacao: str = ""
    engrenagem: str = ""
    fios: str = ""
    maquinas: str = ""

class OrdemProducaoUpdate(BaseModel):
    status: str

class OrdemProducaoSaveTemp(BaseModel):
    dados_temporarios_maquinas: list  # Machine allocations temporários
    espula_data: dict  # Outros dados do formulário de espulagem

# Banco de Dados - Cadastro de Artigos
class ArtigoBancoDados(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    artigo: str  # Nome/código do artigo
    engrenagem: str = ""  # Engrenagem
    fios: str = ""  # Quantidade de fios
    maquinas: str = ""  # Máquinas recomendadas
    ciclos: str = ""  # Ciclos - obrigatório na criação, mas default para suportar dados antigos
    carga: str = ""  # Carga - obrigatório na criação, mas default para suportar dados antigos
    created_at: datetime = Field(default_factory=get_utc_now)
    updated_at: datetime = Field(default_factory=get_utc_now)

class ArtigoBancoDadosCreate(BaseModel):
    artigo: str
    engrenagem: str = ""
    fios: str = ""
    maquinas: str = ""
    ciclos: str
    carga: str

class ArtigoBancoDadosUpdate(BaseModel):
    artigo: Optional[str] = None
    engrenagem: Optional[str] = None
    fios: Optional[str] = None
    maquinas: Optional[str] = None
    ciclos: Optional[str] = None
    carga: Optional[str] = None

class MachineAllocation(BaseModel):
    machine_code: str
    machine_id: str
    layout_type: str
    quantidade: str

class Espula(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    # New fields
    numero_os: Optional[str] = ""  # OS number from ordem de producao
    ordem_producao_id: Optional[str] = None  # Link to ordem de producao
    maquina: Optional[str] = ""  # Machine code
    mat_prima: Optional[str] = ""  # Raw material
    qtde_fios: Optional[str] = ""  # Number of threads
    # Machine allocations - up to 5 machines with quantities
    machine_allocations: list[MachineAllocation] = Field(default_factory=list)
    # Cargas e fração - dynamic list (replaces the 5 fixed fields)
    cargas_fracoes: list[str] = Field(default_factory=list)
    # Existing fields
    cliente: str
    artigo: str
    cor: str
    quantidade_metros: str  # Changed from quantidade to quantidade_metros
    carga: str  # New field for carga (letters and numbers allowed)
    observacoes: Optional[str] = ""
    status: str = "pendente"  # pendente, em_producao_aguardando, producao, finalizado
    data_lancamento: datetime = Field(default_factory=get_utc_now)  # Save in UTC
    data_prevista_entrega: str  # ISO date string
    created_by: str
    created_at: datetime = Field(default_factory=get_utc_now)  # Save in UTC
    updated_at: datetime = Field(default_factory=get_utc_now)  # Save in UTC
    # History fields
    iniciado_em: Optional[datetime] = None  # When moved to em_producao_aguardando
    finalizado_em: Optional[datetime] = None  # When moved to finalizado

class EspulaCreate(BaseModel):
    # New optional fields
    numero_os: Optional[str] = ""
    ordem_producao_id: Optional[str] = None
    maquina: Optional[str] = ""
    mat_prima: Optional[str] = ""
    qtde_fios: Optional[str] = ""
    machine_allocations: list[MachineAllocation] = Field(default_factory=list)
    cargas_fracoes: list[str] = Field(default_factory=list)
    # Existing fields
    cliente: str
    artigo: str
    cor: str
    quantidade_metros: str  # Changed from quantidade to quantidade_metros
    carga: str  # New field for carga (letters and numbers allowed)
    observacoes: str = ""
    data_prevista_entrega: str  # Will be converted to datetime

class EspulaUpdate(BaseModel):
    status: str

class StatusHistory(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    machine_id: str
    machine_code: str
    layout_type: str
    old_status: str
    new_status: str
    changed_by: str
    changed_at: datetime = Field(default_factory=get_utc_now)
    order_id: Optional[str] = None
    maintenance_id: Optional[str] = None

# Helper functions
def hash_password(password: str) -> str:
    return pwd_context.hash(password)

def verify_password(plain_password: str, hashed_password: str) -> bool:
    return pwd_context.verify(plain_password, hashed_password)

def create_access_token(user_id: str, username: str, role: str) -> str:
    payload = {
        "user_id": user_id,
        "username": username,
        "role": role,
        "exp": get_brazil_time().timestamp() + (24 * 60 * 60)  # 24 hours
    }
    return jwt.encode(payload, JWT_SECRET, algorithm="HS256")

async def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    try:
        payload = jwt.decode(credentials.credentials, JWT_SECRET, algorithms=["HS256"])
        user = await db.users.find_one({"id": payload["user_id"]})
        if not user:
            raise HTTPException(status_code=401, detail="User not found")
        return User(**user)
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except jwt.JWTError:
        raise HTTPException(status_code=401, detail="Invalid token")

# Initialize data
async def init_data():
    # Clear existing data except users
    await db.orders.delete_many({})
    await db.espulas.delete_many({})
    await db.ordens_producao.delete_many({})
    await db.maintenance.delete_many({})
    await db.status_history.delete_many({})
    await db.machines.delete_many({})
    await db.banco_dados.delete_many({})  # Reset banco de dados de artigos
    
    # Create default users if they don't exist
    admin_exists = await db.users.find_one({"username": "admin"})
    if not admin_exists:
        admin_user = User(
            username="admin",
            email="admin@mercotextil.com",
            role="admin",
            permissions={
                "dashboard": True,
                "producao": True,
                "ordem_producao": True,
                "relatorios": True,
                "espulagem": True,
                "manutencao": True,
                "banco_dados": True,
                "administracao": True
            }
        )
        admin_dict = admin_user.dict()
        admin_dict["password"] = hash_password("admin123")
        await db.users.insert_one(admin_dict)

    interno_exists = await db.users.find_one({"username": "interno"})
    if not interno_exists:
        interno_user = User(
            username="interno",
            email="interno@mercotextil.com",
            role="operador_interno",
            permissions={
                "dashboard": True,
                "producao": True,
                "ordem_producao": True,
                "relatorios": True,
                "espulagem": True,
                "manutencao": True,
                "banco_dados": True,
                "administracao": False
            }
        )
        interno_dict = interno_user.dict()
        interno_dict["password"] = hash_password("interno123")
        await db.users.insert_one(interno_dict)

    externo_exists = await db.users.find_one({"username": "externo"})
    if not externo_exists:
        externo_user = User(
            username="externo",
            email="externo@mercotextil.com",
            role="operador_externo"
        )
        externo_dict = externo_user.dict()
        externo_dict["password"] = hash_password("externo123")
        await db.users.insert_one(externo_dict)

    # Initialize machines for 16 fusos layout - EXACT as per user image
    layout_16_machines = [
        # Top row blocks CD1-CD4, CD5-CD8 (2x2 each), CD17-CD20 (1x4)
        {"code": "CD1", "position": "block1-1"},
        {"code": "CD2", "position": "block1-2"},
        {"code": "CD3", "position": "block1-3"},
        {"code": "CD4", "position": "block1-4"},
        {"code": "CD5", "position": "block2-1"},
        {"code": "CD6", "position": "block2-2"},
        {"code": "CD7", "position": "block2-3"},
        {"code": "CD8", "position": "block2-4"},
        {"code": "CD17", "position": "block3-1"},
        {"code": "CD18", "position": "block3-2"},
        {"code": "CD19", "position": "block3-3"},
        {"code": "CD20", "position": "block3-4"},
        # Middle blocks CD9-CD12, CD13-CD16 (2x2 each), CD21-CD24 (1x4)
        {"code": "CD9", "position": "block4-1"},
        {"code": "CD10", "position": "block4-2"},
        {"code": "CD11", "position": "block4-3"},
        {"code": "CD12", "position": "block4-4"},
        {"code": "CD13", "position": "block5-1"},
        {"code": "CD14", "position": "block5-2"},
        {"code": "CD15", "position": "block5-3"},
        {"code": "CD16", "position": "block5-4"},
        {"code": "CD21", "position": "block6-1"},
        {"code": "CD22", "position": "block6-2"},
        {"code": "CD23", "position": "block6-3"},
        {"code": "CD24", "position": "block6-4"},
        # CI block (1x4) - labeled as "17 FUSOS" in image
        {"code": "CI1", "position": "ci-1"},
        {"code": "CI2", "position": "ci-2"},
        {"code": "CI3", "position": "ci-3"},
        {"code": "CI4", "position": "ci-4"},
        # F blocks (bottom section) F1-F24
        {"code": "F1", "position": "f-1"}, {"code": "F2", "position": "f-2"},
        {"code": "F3", "position": "f-3"}, {"code": "F4", "position": "f-4"},
        {"code": "F5", "position": "f-5"}, {"code": "F6", "position": "f-6"},
        {"code": "F7", "position": "f-7"}, {"code": "F8", "position": "f-8"},
        {"code": "F9", "position": "f-9"}, {"code": "F10", "position": "f-10"},
        {"code": "F11", "position": "f-11"}, {"code": "F12", "position": "f-12"},
        {"code": "F13", "position": "f-13"}, {"code": "F14", "position": "f-14"},
        {"code": "F15", "position": "f-15"}, {"code": "F16", "position": "f-16"},
        {"code": "F17", "position": "f-17"}, {"code": "F18", "position": "f-18"},
        {"code": "F19", "position": "f-19"}, {"code": "F20", "position": "f-20"},
        {"code": "F21", "position": "f-21"}, {"code": "F22", "position": "f-22"},
        {"code": "F23", "position": "f-23"}, {"code": "F24", "position": "f-24"}
    ]

    # Create 16 fusos machines
    for i, machine_data in enumerate(layout_16_machines):
        machine = Machine(
            code=machine_data["code"],
            position=machine_data["position"],
            layout_type="16_fusos"
        )
        machine_dict = machine.dict()
        machine_dict["id"] = f"16_fusos_{machine_data['code']}_{str(uuid.uuid4())[:8]}"
        await db.machines.insert_one(machine_dict)

    # Initialize machines for 32 fusos layout - EXACT as per user image
    layout_32_machines = [
        # Top row CT1-CT24
        {"code": "CT1", "position": "ct-1"}, {"code": "CT2", "position": "ct-2"},
        {"code": "CT3", "position": "ct-3"}, {"code": "CT4", "position": "ct-4"},
        {"code": "CT5", "position": "ct-5"}, {"code": "CT6", "position": "ct-6"},
        {"code": "CT7", "position": "ct-7"}, {"code": "CT8", "position": "ct-8"},
        {"code": "CT9", "position": "ct-9"}, {"code": "CT10", "position": "ct-10"},
        {"code": "CT11", "position": "ct-11"}, {"code": "CT12", "position": "ct-12"},
        {"code": "CT13", "position": "ct-13"}, {"code": "CT14", "position": "ct-14"},
        {"code": "CT15", "position": "ct-15"}, {"code": "CT16", "position": "ct-16"},
        {"code": "CT17", "position": "ct-17"}, {"code": "CT18", "position": "ct-18"},
        {"code": "CT19", "position": "ct-19"}, {"code": "CT20", "position": "ct-20"},
        {"code": "CT21", "position": "ct-21"}, {"code": "CT22", "position": "ct-22"},
        {"code": "CT23", "position": "ct-23"}, {"code": "CT24", "position": "ct-24"},
        # U groups (3 columns of 10 machines each)
        {"code": "U1", "position": "u1-1"}, {"code": "U2", "position": "u1-2"},
        {"code": "U3", "position": "u1-3"}, {"code": "U4", "position": "u1-4"},
        {"code": "U5", "position": "u1-5"}, {"code": "U6", "position": "u1-6"},
        {"code": "U7", "position": "u1-7"}, {"code": "U8", "position": "u1-8"},
        {"code": "U9", "position": "u1-9"}, {"code": "U10", "position": "u1-10"},
        {"code": "U11", "position": "u2-1"}, {"code": "U12", "position": "u2-2"},
        {"code": "U13", "position": "u2-3"}, {"code": "U14", "position": "u2-4"},
        {"code": "U15", "position": "u2-5"}, {"code": "U16", "position": "u2-6"},
        {"code": "U17", "position": "u2-7"}, {"code": "U18", "position": "u2-8"},
        {"code": "U19", "position": "u2-9"}, {"code": "U20", "position": "u2-10"},
        {"code": "U21", "position": "u3-1"}, {"code": "U22", "position": "u3-2"},
        {"code": "U23", "position": "u3-3"}, {"code": "U24", "position": "u3-4"},
        {"code": "U25", "position": "u3-5"}, {"code": "U26", "position": "u3-6"},
        {"code": "U27", "position": "u3-7"}, {"code": "U28", "position": "u3-8"},
        {"code": "U29", "position": "u3-9"}, {"code": "U30", "position": "u3-10"},
        # N row (N1-N10)
        {"code": "N1", "position": "n-1"}, {"code": "N2", "position": "n-2"},
        {"code": "N3", "position": "n-3"}, {"code": "N4", "position": "n-4"},
        {"code": "N5", "position": "n-5"}, {"code": "N6", "position": "n-6"},
        {"code": "N7", "position": "n-7"}, {"code": "N8", "position": "n-8"},
        {"code": "N9", "position": "n-9"}, {"code": "N10", "position": "n-10"},
        # Additional U machines from image (U31-U33)
        {"code": "U31", "position": "u4-1"}, {"code": "U32", "position": "u4-2"},
        {"code": "U33", "position": "u4-3"}
    ]

    # Create 32 fusos machines
    for i, machine_data in enumerate(layout_32_machines):
        machine = Machine(
            code=machine_data["code"],
            position=machine_data["position"],
            layout_type="32_fusos"
        )
        machine_dict = machine.dict()
        machine_dict["id"] = f"32_fusos_{machine_data['code']}_{str(uuid.uuid4())[:8]}"
        await db.machines.insert_one(machine_dict)

@api_router.post("/reset-database")
async def reset_database(current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    await init_data()
    return {"message": "Database reset successfully, keeping only users"}

# Auth routes
@api_router.post("/auth/login", response_model=LoginResponse)
async def login(login_data: UserLogin):
    user = await db.users.find_one({"username": login_data.username})
    if not user or not verify_password(login_data.password, user["password"]):
        raise HTTPException(status_code=401, detail="Invalid credentials")
    
    if not user["active"]:
        raise HTTPException(status_code=401, detail="User inactive")
    
    token = create_access_token(user["id"], user["username"], user["role"])
    user_obj = User(**user)
    return LoginResponse(token=token, user=user_obj)

@api_router.get("/auth/me", response_model=User)
async def get_me(current_user: User = Depends(get_current_user)):
    return current_user

# User management routes (admin only)
@api_router.post("/users", response_model=User)
async def create_user(user_data: UserCreate, current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    # Check if user already exists
    existing_user = await db.users.find_one({"username": user_data.username})
    if existing_user:
        raise HTTPException(status_code=400, detail="Username already exists")
    
    # Default permissions if not provided
    permissions = user_data.permissions if user_data.permissions else {
        "dashboard": True,
        "producao": True,
        "ordem_producao": True,
        "relatorios": True,
        "espulagem": True,
        "manutencao": True,
        "administracao": user_data.role == "admin"
    }
    
    user = User(
        username=user_data.username,
        email=user_data.email,
        role=user_data.role,
        permissions=permissions
    )
    user_dict = user.dict()
    user_dict["password"] = hash_password(user_data.password)
    await db.users.insert_one(user_dict)
    return user

@api_router.get("/users", response_model=List[User])
async def get_users(current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    users = await db.users.find().to_list(1000)
    return [User(**user) for user in users]

@api_router.put("/users/{user_id}", response_model=User)
async def update_user(user_id: str, user_data: UserUpdate, current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    # Check if user exists
    existing_user = await db.users.find_one({"id": user_id})
    if not existing_user:
        raise HTTPException(status_code=404, detail="User not found")
    
    # Don't allow admin to remove their own admin role
    if user_id == current_user.id and user_data.role and user_data.role != "admin":
        raise HTTPException(status_code=400, detail="Cannot change your own admin role")
    
    # Build update dict with only provided fields
    update_data = {"updated_at": get_utc_now()}
    
    if user_data.username is not None:
        # Check if username is already taken by another user
        username_exists = await db.users.find_one({"username": user_data.username, "id": {"$ne": user_id}})
        if username_exists:
            raise HTTPException(status_code=400, detail="Username already exists")
        update_data["username"] = user_data.username
    
    if user_data.email is not None:
        update_data["email"] = user_data.email
    
    if user_data.password is not None and user_data.password != "":
        update_data["password"] = hash_password(user_data.password)
    
    if user_data.role is not None:
        update_data["role"] = user_data.role
    
    if user_data.permissions is not None:
        update_data["permissions"] = user_data.permissions
    
    if user_data.active is not None:
        update_data["active"] = user_data.active
    
    await db.users.update_one({"id": user_id}, {"$set": update_data})
    
    # Return updated user
    updated_user = await db.users.find_one({"id": user_id})
    return User(**updated_user)

@api_router.delete("/users/{user_id}")
async def delete_user(user_id: str, current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    # Don't allow admin to delete themselves
    if user_id == current_user.id:
        raise HTTPException(status_code=400, detail="Cannot delete your own account")
    
    result = await db.users.delete_one({"id": user_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="User not found")
    
    return {"message": "User deleted successfully"}

# Machine routes
@api_router.get("/machines", response_model=List[Machine])
async def get_all_machines(current_user: User = Depends(get_current_user)):
    """Get all machines (all layout types)"""
    machines = await db.machines.find({}).to_list(1000)
    return [Machine(**machine) for machine in machines]

@api_router.get("/machines/{layout_type}", response_model=List[Machine])
async def get_machines(layout_type: str, current_user: User = Depends(get_current_user)):
    machines = await db.machines.find({"layout_type": layout_type}).to_list(1000)
    return [Machine(**machine) for machine in machines]

# Maintenance routes
@api_router.post("/maintenance", response_model=Maintenance)
async def create_maintenance(maintenance_data: MaintenanceCreate, current_user: User = Depends(get_current_user)):
    # Find machine by ID
    machine = await db.machines.find_one({"id": maintenance_data.machine_id})
    if not machine:
        raise HTTPException(status_code=404, detail="Machine not found")
    
    # Allow maintenance for any status except already in maintenance
    if machine["status"] == "azul":
        raise HTTPException(status_code=400, detail="Máquina já está em manutenção")
    
    # Save the previous status to restore later
    previous_status = machine["status"]
    
    maintenance = Maintenance(
        machine_id=maintenance_data.machine_id,
        machine_code=machine["code"],
        motivo=maintenance_data.motivo,
        created_by=current_user.username
    )
    
    maintenance_dict = maintenance.dict()
    maintenance_dict["previous_status"] = previous_status  # Save previous status
    
    await db.maintenance.insert_one(maintenance_dict)
    
    # Update machine status to azul (maintenance)
    await db.machines.update_one(
        {"id": maintenance_data.machine_id},
        {"$set": {"status": "azul", "updated_at": get_utc_now()}}
    )
    
    return maintenance

@api_router.get("/maintenance", response_model=List[Maintenance])
async def get_maintenance(current_user: User = Depends(get_current_user)):
    maintenances = await db.maintenance.find().sort("created_at", -1).to_list(1000)
    return [Maintenance(**m) for m in maintenances]

@api_router.put("/maintenance/{maintenance_id}/finish")
async def finish_maintenance(maintenance_id: str, current_user: User = Depends(get_current_user)):
    maintenance = await db.maintenance.find_one({"id": maintenance_id})
    if not maintenance:
        raise HTTPException(status_code=404, detail="Maintenance not found")
    
    if maintenance["status"] == "finalizada":
        raise HTTPException(status_code=400, detail="Maintenance already finished")
    
    # Update maintenance status
    await db.maintenance.update_one(
        {"id": maintenance_id},
        {
            "$set": {
                "status": "finalizada",
                "finished_at": get_utc_now(),
                "finished_by": current_user.username
            }
        }
    )
    
    # Restore previous status or default to verde
    previous_status = maintenance.get("previous_status", "verde")
    
    # Check if there are pending orders for this machine - if so, set to amarelo
    pending_orders = await db.orders.find({
        "machine_id": maintenance["machine_id"],
        "status": "pendente"
    }).to_list(1)
    
    # Check if there's an order in production for this machine
    in_production = await db.orders.find_one({
        "machine_id": maintenance["machine_id"],
        "status": "em_producao"
    })
    
    if in_production:
        # If there was an order in production, return to vermelho
        restore_status = "vermelho"
    elif pending_orders:
        # If there are pending orders, set to amarelo
        restore_status = "amarelo"
    else:
        # Otherwise, return to verde
        restore_status = "verde"
    
    await db.machines.update_one(
        {"id": maintenance["machine_id"]},
        {"$set": {"status": restore_status, "updated_at": get_utc_now()}}
    )
    
    return {"message": "Maintenance finished successfully"}

# Admin-only endpoint to activate/deactivate machines
@api_router.put("/machines/{machine_id}/toggle-active")
async def toggle_machine_active(machine_id: str, current_user: User = Depends(get_current_user)):
    # Verify admin user
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Only admins can activate/deactivate machines")
    
    machine = await db.machines.find_one({"id": machine_id})
    if not machine:
        raise HTTPException(status_code=404, detail="Machine not found")
    
    # Toggle active status
    new_active_status = not machine.get("active", True)
    new_status = "desativada" if not new_active_status else "verde"
    
    await db.machines.update_one(
        {"id": machine_id},
        {
            "$set": {
                "active": new_active_status,
                "status": new_status,
                "updated_at": get_utc_now()
            }
        }
    )
    
    return {
        "message": f"Machine {'deactivated' if not new_active_status else 'activated'} successfully",
        "machine_id": machine_id,
        "active": new_active_status,
        "status": new_status
    }

# Order routes
@api_router.post("/orders", response_model=Order)
async def create_order(order_data: OrderCreate, current_user: User = Depends(get_current_user)):
    if current_user.role not in ["admin", "operador_interno"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    
    # Check if machine exists and is available
    machine = await db.machines.find_one({"id": order_data.machine_id})
    if not machine:
        raise HTTPException(status_code=404, detail="Machine not found")
    
    if machine["status"] != "verde":
        raise HTTPException(status_code=400, detail="Machine not available")
    
    order = Order(
        machine_id=order_data.machine_id,
        machine_code=machine["code"],
        layout_type=machine["layout_type"],
        cliente=order_data.cliente,
        artigo=order_data.artigo,
        cor=order_data.cor,
        quantidade=order_data.quantidade,
        observacao=order_data.observacao,
        created_by=current_user.username
    )
    
    await db.orders.insert_one(order.dict())
    
    # Update machine status to amarelo (pending)
    await db.machines.update_one(
        {"id": order_data.machine_id},
        {"$set": {"status": "amarelo", "updated_at": get_utc_now()}}
    )
    
    return order

@api_router.get("/orders", response_model=List[Order])
async def get_orders(current_user: User = Depends(get_current_user)):
    orders = await db.orders.find().sort("created_at", -1).to_list(1000)
    return [Order(**order) for order in orders]

@api_router.put("/orders/{order_id}")
async def update_order(
    order_id: str, 
    order_update: OrderUpdate, 
    current_user: User = Depends(get_current_user)
):
    if current_user.role not in ["admin", "operador_externo"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    
    order = await db.orders.find_one({"id": order_id})
    if not order:
        raise HTTPException(status_code=404, detail="Order not found")
    
    update_data = {
        "observacao_liberacao": order_update.observacao_liberacao,
        "laudo_final": order_update.laudo_final
    }
    machine_status = "amarelo"
    
    if order_update.status == "em_producao":
        update_data["status"] = "em_producao"
        update_data["started_at"] = get_utc_now()
        machine_status = "vermelho"
    elif order_update.status == "finalizado":
        update_data["status"] = "finalizado"
        update_data["finished_at"] = get_utc_now()
        
        # Check if there are other pending orders for this machine
        pending_orders = await db.orders.find({
            "machine_id": order["machine_id"],
            "status": "pendente",
            "id": {"$ne": order_id}
        }).to_list(1)
        
        # If there are pending orders, keep machine yellow, otherwise green
        machine_status = "amarelo" if pending_orders else "verde"
    
    await db.orders.update_one({"id": order_id}, {"$set": update_data})
    
    # Update machine status
    await db.machines.update_one(
        {"id": order["machine_id"]},
        {"$set": {"status": machine_status, "updated_at": get_utc_now()}}
    )
    
    return {"message": "Order updated successfully"}

# Machine-specific order routes
@api_router.get("/machines/{machine_code}/orders", response_model=List[Order])
async def get_machine_orders(machine_code: str, current_user: User = Depends(get_current_user)):
    """Get all orders for a specific machine, sorted by most recent first"""
    orders = await db.orders.find({"machine_code": machine_code}).sort("created_at", -1).to_list(1000)
    return [Order(**order) for order in orders]

@api_router.post("/machines/{machine_code}/orders", response_model=Order)
async def create_machine_order(
    machine_code: str,
    order_data: OrderCreate,
    current_user: User = Depends(get_current_user)
):
    """Create order directly on machine (manual)"""
    if current_user.role not in ["admin", "operador_interno"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    
    # Get machine
    machine = await db.machines.find_one({"code": machine_code})
    if not machine:
        raise HTTPException(status_code=404, detail="Machine not found")
    
    # Get next queue position for this machine
    existing_orders = await db.orders.find({"machine_code": machine_code}).sort("queue_position", -1).limit(1).to_list(1)
    next_position = (existing_orders[0]["queue_position"] + 1) if existing_orders else 1
    
    order = Order(
        machine_id=machine["id"],
        machine_code=machine_code,
        layout_type=machine["layout_type"],
        cliente=order_data.cliente,
        artigo=order_data.artigo,
        cor=order_data.cor,
        quantidade=order_data.quantidade,
        observacao=order_data.observacao,
        created_by=current_user.username,
        origem="manual",
        queue_position=next_position
    )
    
    await db.orders.insert_one(order.dict())
    
    # Update machine status to amarelo (has pending order)
    await db.machines.update_one(
        {"id": machine["id"]},
        {"$set": {"status": "amarelo", "updated_at": get_utc_now()}}
    )
    
    return order

@api_router.put("/machines/{machine_code}/orders/{order_id}/start")
async def start_machine_order(
    machine_code: str,
    order_id: str,
    current_user: User = Depends(get_current_user)
):
    """Start production of a specific order in machine queue"""
    if current_user.role not in ["admin", "operador_externo", "operador_interno"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    
    # Check if machine already has an order in production
    existing_production = await db.orders.find_one({
        "machine_code": machine_code,
        "status": "em_producao"
    })
    
    if existing_production:
        raise HTTPException(status_code=400, detail="Machine already has an order in production")
    
    # Get the order
    order = await db.orders.find_one({"id": order_id, "machine_code": machine_code})
    if not order:
        raise HTTPException(status_code=404, detail="Order not found")
    
    if order["status"] != "pendente":
        raise HTTPException(status_code=400, detail="Order is not pending")
    
    # Update order to em_producao
    await db.orders.update_one(
        {"id": order_id},
        {"$set": {
            "status": "em_producao",
            "started_at": get_utc_now()
        }}
    )
    
    # Update machine status to vermelho (in production)
    await db.machines.update_one(
        {"code": machine_code},
        {"$set": {"status": "vermelho", "updated_at": get_utc_now()}}
    )
    
    return {"message": "Order production started successfully"}

@api_router.put("/machines/{machine_code}/orders/{order_id}/finish")
async def finish_machine_order(
    machine_code: str,
    order_id: str,
    current_user: User = Depends(get_current_user)
):
    """Finish order production and update machine status"""
    # Find the order
    order = await db.orders.find_one({"id": order_id, "machine_code": machine_code})
    if not order:
        raise HTTPException(status_code=404, detail="Order not found")
    
    if order["status"] != "em_producao":
        raise HTTPException(status_code=400, detail="Order is not in production")
    
    # Update order status to finalizado
    await db.orders.update_one(
        {"id": order_id},
        {
            "$set": {
                "status": "finalizado",
                "finished_at": get_utc_now(),
                "finished_by": current_user.username,
                "updated_at": get_utc_now()
            }
        }
    )
    
    # Check for remaining pending orders for this machine
    pending_orders = await db.orders.find({
        "machine_code": machine_code,
        "status": "pendente"
    }).to_list(100)
    
    # Check for any order still in production
    in_production = await db.orders.find_one({
        "machine_code": machine_code,
        "status": "em_producao"
    })
    
    # Update machine status
    if in_production:
        # Still has orders in production
        new_status = "vermelho"
    elif pending_orders:
        # Has pending orders
        new_status = "amarelo"
    else:
        # No more orders
        new_status = "verde"
    
    await db.machines.update_one(
        {"code": machine_code},
        {"$set": {"status": new_status, "updated_at": get_utc_now()}}
    )
    
    return {"message": "Order finished successfully"}

@api_router.delete("/orders/{order_id}")
async def delete_order(order_id: str, current_user: User = Depends(get_current_user)):
    """Delete order from machine queue"""
    try:
        order = await db.orders.find_one({"id": order_id})
        if not order:
            raise HTTPException(status_code=404, detail="Order not found")
        
        # Deletar ordem
        await db.orders.delete_one({"id": order_id})
        
        # Atualizar status da máquina se não tiver mais pedidos
        remaining_orders = await db.orders.find({"machine_code": order["machine_code"]}).to_list(100)
        
        if len(remaining_orders) == 0:
            # Sem pedidos - voltar para verde
            await db.machines.update_one(
                {"code": order["machine_code"]},
                {"$set": {"status": "verde", "updated_at": get_utc_now()}}
            )
        elif not any(o["status"] == "em_producao" for o in remaining_orders):
            # Tem pedidos mas nenhum em produção - amarelo
            await db.machines.update_one(
                {"code": order["machine_code"]},
                {"$set": {"status": "amarelo", "updated_at": get_utc_now()}}
            )
        
        return {"message": "Order deleted successfully"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error deleting order: {str(e)}")

# Ordem de Producao routes
@api_router.get("/ordens-producao/next-number")
async def get_next_ordem_number(current_user: User = Depends(get_current_user)):
    """Get the next sequential OS number"""
    try:
        # Find the highest numero_os
        ordens = await db.ordens_producao.find().sort("numero_os", -1).limit(1).to_list(1)
        
        if not ordens:
            return {"numero_os": "0001"}
        
        last_number = int(ordens[0]["numero_os"])
        next_number = last_number + 1
        
        # Format with leading zeros (minimum 4 digits)
        next_os = str(next_number).zfill(max(4, len(str(next_number))))
        
        return {"numero_os": next_os}
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error getting next number: {str(e)}")

@api_router.post("/ordens-producao", response_model=OrdemProducao)
async def create_ordem_producao(ordem_data: OrdemProducaoCreate, current_user: User = Depends(get_current_user)):
    try:
        # Get next OS number
        next_number_response = await get_next_ordem_number(current_user)
        numero_os = next_number_response["numero_os"]
        
        ordem = OrdemProducao(
            numero_os=numero_os,
            cliente=ordem_data.cliente,
            artigo=ordem_data.artigo,
            cor=ordem_data.cor,
            metragem=ordem_data.metragem,
            data_entrega=ordem_data.data_entrega,
            observacao=ordem_data.observacao,
            engrenagem=ordem_data.engrenagem,
            fios=ordem_data.fios,
            maquinas=ordem_data.maquinas,
            criado_por=current_user.username
        )
        
        await db.ordens_producao.insert_one(ordem.dict())
        return ordem
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error creating ordem de producao: {str(e)}")

@api_router.get("/ordens-producao", response_model=List[OrdemProducao])
async def get_ordens_producao(current_user: User = Depends(get_current_user)):
    ordens = await db.ordens_producao.find().sort("criado_em", -1).to_list(1000)
    return [OrdemProducao(**ordem) for ordem in ordens]

@api_router.get("/ordens-producao/pendentes", response_model=List[OrdemProducao])
async def get_ordens_producao_pendentes(current_user: User = Depends(get_current_user)):
    """Get only pending ordens de producao for Relatorios tab"""
    ordens = await db.ordens_producao.find({"status": "pendente"}).sort("criado_em", -1).to_list(1000)
    return [OrdemProducao(**ordem) for ordem in ordens]

@api_router.get("/ordens-producao/{ordem_id}", response_model=OrdemProducao)
async def get_ordem_producao(ordem_id: str, current_user: User = Depends(get_current_user)):
    ordem = await db.ordens_producao.find_one({"id": ordem_id})
    if not ordem:
        raise HTTPException(status_code=404, detail="Ordem de producao not found")
    return OrdemProducao(**ordem)

@api_router.put("/ordens-producao/{ordem_id}")
async def update_ordem_producao(
    ordem_id: str,
    ordem_update: OrdemProducaoUpdate,
    current_user: User = Depends(get_current_user)
):
    ordem = await db.ordens_producao.find_one({"id": ordem_id})
    if not ordem:
        raise HTTPException(status_code=404, detail="Ordem de producao not found")
    
    update_data = {"status": ordem_update.status, "updated_at": get_utc_now()}
    
    if ordem_update.status == "em_producao" and not ordem.get("iniciado_em"):
        update_data["iniciado_em"] = get_utc_now()
    elif ordem_update.status == "finalizado":
        update_data["finalizado_em"] = get_utc_now()
    
    await db.ordens_producao.update_one({"id": ordem_id}, {"$set": update_data})
    
    return {"message": "Ordem de producao updated successfully"}

@api_router.put("/ordens-producao/{ordem_id}/salvar-temporarios")
async def salvar_temporarios_ordem_producao(
    ordem_id: str,
    temp_data: OrdemProducaoSaveTemp,
    current_user: User = Depends(get_current_user)
):
    """Save temporary data for ordem de producao without creating espulagem"""
    ordem = await db.ordens_producao.find_one({"id": ordem_id})
    if not ordem:
        raise HTTPException(status_code=404, detail="Ordem de producao not found")
    
    update_data = {
        "dados_temporarios_maquinas": temp_data.dados_temporarios_maquinas,
        "espula_data_temp": temp_data.espula_data,
        "editado_por": current_user.username,
        "editado_em": get_utc_now(),
        "updated_at": get_utc_now()
    }
    
    await db.ordens_producao.update_one({"id": ordem_id}, {"$set": update_data})
    
    return {"message": "Dados temporários salvos com sucesso"}


@api_router.delete("/ordens-producao/{ordem_id}")
async def delete_ordem_producao(ordem_id: str, current_user: User = Depends(get_current_user)):
    """Delete ordem de producao"""
    try:
        result = await db.ordens_producao.delete_one({"id": ordem_id})
        if result.deleted_count == 0:
            raise HTTPException(status_code=404, detail="Ordem de producao not found")
        
        return {"message": "Ordem de producao deleted successfully"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error deleting ordem: {str(e)}")


# Banco de Dados - Artigos Routes
@api_router.post("/banco-dados", response_model=ArtigoBancoDados)
async def create_artigo_banco_dados(
    artigo_data: ArtigoBancoDadosCreate,
    current_user: User = Depends(get_current_user)
):
    """Create new artigo in banco de dados"""
    try:
        artigo = ArtigoBancoDados(
            artigo=artigo_data.artigo,
            engrenagem=artigo_data.engrenagem,
            fios=artigo_data.fios,
            maquinas=artigo_data.maquinas,
            ciclos=artigo_data.ciclos,
            carga=artigo_data.carga
        )
        
        await db.banco_dados.insert_one(artigo.dict())
        return artigo
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error creating artigo: {str(e)}")

@api_router.get("/banco-dados", response_model=List[ArtigoBancoDados])
async def get_artigos_banco_dados(current_user: User = Depends(get_current_user)):
    """Get all artigos from banco de dados"""
    artigos = await db.banco_dados.find().sort("created_at", -1).to_list(1000)
    return [ArtigoBancoDados(**artigo) for artigo in artigos]

@api_router.get("/banco-dados/search")
async def search_artigos_banco_dados(
    q: str,
    current_user: User = Depends(get_current_user)
):
    """Search artigos by name (autocomplete)"""
    try:
        # Case-insensitive regex search
        artigos = await db.banco_dados.find({
            "artigo": {"$regex": q, "$options": "i"}
        }).limit(10).to_list(10)
        
        return [ArtigoBancoDados(**artigo) for artigo in artigos]
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error searching artigos: {str(e)}")

@api_router.put("/banco-dados/{artigo_id}", response_model=ArtigoBancoDados)
async def update_artigo_banco_dados(
    artigo_id: str,
    artigo_update: ArtigoBancoDadosUpdate,
    current_user: User = Depends(get_current_user)
):
    """Update artigo in banco de dados"""
    artigo = await db.banco_dados.find_one({"id": artigo_id})
    if not artigo:
        raise HTTPException(status_code=404, detail="Artigo not found")
    
    update_data = {}
    if artigo_update.artigo is not None:
        update_data["artigo"] = artigo_update.artigo
    if artigo_update.engrenagem is not None:
        update_data["engrenagem"] = artigo_update.engrenagem
    if artigo_update.fios is not None:
        update_data["fios"] = artigo_update.fios
    if artigo_update.maquinas is not None:
        update_data["maquinas"] = artigo_update.maquinas
    if artigo_update.ciclos is not None:
        update_data["ciclos"] = artigo_update.ciclos
    if artigo_update.carga is not None:
        update_data["carga"] = artigo_update.carga
    
    update_data["updated_at"] = get_utc_now()
    
    await db.banco_dados.update_one({"id": artigo_id}, {"$set": update_data})
    
    updated_artigo = await db.banco_dados.find_one({"id": artigo_id})
    return ArtigoBancoDados(**updated_artigo)

@api_router.delete("/banco-dados/{artigo_id}")
async def delete_artigo_banco_dados(
    artigo_id: str,
    current_user: User = Depends(get_current_user)
):
    """Delete artigo from banco de dados"""
    result = await db.banco_dados.delete_one({"id": artigo_id})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="Artigo not found")
    
    return {"message": "Artigo deleted successfully"}


# Espulas routes
@api_router.post("/espulas", response_model=Espula)
async def create_espula(espula_data: EspulaCreate, current_user: User = Depends(get_current_user)):
    try:
        # If espula has numero_os but no ordem_producao_id, create ordem to reserve the number
        if espula_data.numero_os and not espula_data.ordem_producao_id:
            # Create ordem de producao to reserve this OS number
            ordem = OrdemProducao(
                numero_os=espula_data.numero_os,
                cliente=espula_data.cliente,
                artigo=espula_data.artigo,
                cor=espula_data.cor,
                metragem=espula_data.quantidade_metros,
                data_entrega=espula_data.data_prevista_entrega,
                observacao=espula_data.observacoes,
                status="em_producao",  # Already in production since espula is being created
                iniciado_em=get_utc_now(),
                criado_por=current_user.username
            )
            await db.ordens_producao.insert_one(ordem.dict())
        
        espula = Espula(
            # New fields
            numero_os=espula_data.numero_os,
            ordem_producao_id=espula_data.ordem_producao_id,
            maquina=espula_data.maquina,
            mat_prima=espula_data.mat_prima,
            qtde_fios=espula_data.qtde_fios,
            machine_allocations=espula_data.machine_allocations,
            cargas_fracoes=espula_data.cargas_fracoes,
            # Existing fields
            cliente=espula_data.cliente,
            artigo=espula_data.artigo,
            cor=espula_data.cor,
            quantidade_metros=espula_data.quantidade_metros,
            carga=espula_data.carga,
            observacoes=espula_data.observacoes,
            data_prevista_entrega=espula_data.data_prevista_entrega,
            created_by=current_user.username
        )
        
        await db.espulas.insert_one(espula.dict())
        
        # If espula is linked to an existing ordem de producao, update the ordem status
        if espula_data.ordem_producao_id:
            await db.ordens_producao.update_one(
                {"id": espula_data.ordem_producao_id},
                {"$set": {
                    "status": "em_producao",
                    "iniciado_em": get_utc_now(),
                    "updated_at": get_utc_now()
                }}
            )
        
        return espula
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error creating espula: {str(e)}")

@api_router.get("/espulas", response_model=List[Espula])
async def get_espulas(current_user: User = Depends(get_current_user)):
    # Get ALL espulas (including finished), sorted by delivery date
    espulas = await db.espulas.find().sort("data_prevista_entrega", 1).to_list(1000)
    return [Espula(**espula) for espula in espulas]

@api_router.put("/espulas/{espula_id}")
async def update_espula(
    espula_id: str, 
    espula_update: EspulaUpdate, 
    current_user: User = Depends(get_current_user)
):
    espula = await db.espulas.find_one({"id": espula_id})
    if not espula:
        raise HTTPException(status_code=404, detail="Espula not found")
    
    update_data = {"status": espula_update.status, "updated_at": get_utc_now()}
    
    if espula_update.status == "em_producao_aguardando" and not espula.get("iniciado_em"):
        update_data["iniciado_em"] = get_utc_now()
    elif espula_update.status == "finalizado":
        update_data["finalizado_em"] = get_utc_now()
        
        # If espula is linked to an ordem de producao, update the ordem status to finalizado
        if espula.get("ordem_producao_id"):
            await db.ordens_producao.update_one(
                {"id": espula["ordem_producao_id"]},
                {"$set": {
                    "status": "finalizado",
                    "finalizado_em": get_utc_now(),
                    "updated_at": get_utc_now()
                }}
            )
    
    await db.espulas.update_one({"id": espula_id}, {"$set": update_data})
    
    return {"message": "Espula updated successfully"}


class MachineAllocationsUpdate(BaseModel):
    machine_allocations: List[MachineAllocation]


@api_router.put("/espulas/{espula_id}/machine-allocations")
async def update_espula_machine_allocations(
    espula_id: str,
    update_data: MachineAllocationsUpdate,
    current_user: User = Depends(get_current_user)
):
    """Update machine allocations for an espula"""
    espula = await db.espulas.find_one({"id": espula_id})
    if not espula:
        raise HTTPException(status_code=404, detail="Espula not found")
    
    # Convert machine allocations to dict format
    allocations_dict = [alloc.dict() for alloc in update_data.machine_allocations]
    
    await db.espulas.update_one(
        {"id": espula_id}, 
        {"$set": {
            "machine_allocations": allocations_dict,
            "updated_at": get_utc_now()
        }}
    )
    
    return {"message": "Machine allocations updated successfully"}


@api_router.post("/espulas/{espula_id}/finalize-with-machines")
async def finalize_espula_with_machines(
    espula_id: str,
    current_user: User = Depends(get_current_user)
):
    """Finalize espulagem and create orders on selected machines"""
    espula = await db.espulas.find_one({"id": espula_id})
    if not espula:
        raise HTTPException(status_code=404, detail="Espula not found")
    
    if espula["status"] == "finalizado":
        raise HTTPException(status_code=400, detail="Espula already finalized")
    
    # Get machine allocations
    machine_allocations = espula.get("machine_allocations", [])
    
    if not machine_allocations:
        raise HTTPException(status_code=400, detail="No machines allocated")
    
    # Create orders for each machine
    created_orders = []
    for allocation in machine_allocations:
        # Get next queue position for this machine
        existing_orders = await db.orders.find({"machine_code": allocation["machine_code"]}).sort("queue_position", -1).limit(1).to_list(1)
        next_position = (existing_orders[0]["queue_position"] + 1) if existing_orders else 1
        
        order = Order(
            machine_id=allocation["machine_id"],
            machine_code=allocation["machine_code"],
            layout_type=allocation["layout_type"],
            cliente=espula["cliente"],
            artigo=espula["artigo"],
            cor=espula["cor"],
            quantidade=allocation["quantidade"],
            observacao=espula.get("observacoes", ""),
            created_by=current_user.username,
            espulagem_id=espula_id,
            ordem_producao_id=espula.get("ordem_producao_id"),
            numero_os=espula.get("numero_os"),
            origem="espulagem",
            queue_position=next_position
        )
        
        await db.orders.insert_one(order.dict())
        created_orders.append(order.id)
        
        # Update machine status to amarelo ONLY if not already in production (vermelho)
        machine = await db.machines.find_one({"id": allocation["machine_id"]})
        if machine and machine.get("status") != "vermelho":
            await db.machines.update_one(
                {"id": allocation["machine_id"]},
                {"$set": {"status": "amarelo", "updated_at": get_utc_now()}}
            )
    
    # Update espula status to finalizado
    await db.espulas.update_one(
        {"id": espula_id},
        {"$set": {
            "status": "finalizado",
            "finalizado_em": get_utc_now(),
            "updated_at": get_utc_now()
        }}
    )
    
    # If espula is linked to an ordem de producao, update the ordem status to finalizado
    if espula.get("ordem_producao_id"):
        await db.ordens_producao.update_one(
            {"id": espula["ordem_producao_id"]},
            {"$set": {
                "status": "finalizado",
                "finalizado_em": get_utc_now(),
                "updated_at": get_utc_now()
            }}
        )
    
    return {
        "message": "Espula finalized and orders created successfully",
        "orders_created": len(created_orders),
        "order_ids": created_orders
    }

# Reports routes
@api_router.get("/reports/export")
async def export_report(layout_type: str, current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    try:
        # Get all data for the report
        orders = await db.orders.find({"layout_type": layout_type}).to_list(1000)
        history = await db.status_history.find({"layout_type": layout_type}).to_list(1000)
        
        # Serialize the data to make it JSON compatible
        serialized_orders = serialize_docs(orders)
        serialized_history = serialize_docs(history)
        
        return {
            "orders": serialized_orders,
            "status_history": serialized_history,
            "generated_at": get_brazil_time().isoformat(),
            "layout_type": layout_type
        }
    except Exception as e:
        logger.error(f"Error exporting report: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error generating report: {str(e)}")

@api_router.get("/espulas/report")
async def get_espulas_report(current_user: User = Depends(get_current_user)):
    if current_user.role != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    
    try:
        espulas = await db.espulas.find({"status": "finalizado"}).to_list(1000)
        serialized_espulas = serialize_docs(espulas)
        
        return {
            "espulas": serialized_espulas,
            "generated_at": get_brazil_time().isoformat()
        }
    except Exception as e:
        logger.error(f"Error exporting espulas report: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error generating espulas report: {str(e)}")

@api_router.get("/reports/mapa-trancadeiras")
async def get_mapa_trancadeiras(layout_type: str, current_user: User = Depends(get_current_user)):
    if current_user.role not in ["admin", "operador_interno"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    try:
        machines = await db.machines.find({"layout_type": layout_type}).to_list(1000)
        orders = await db.orders.find({
            "layout_type": layout_type,
            "status": {"$in": ["em_producao", "pendente"]}
        }).to_list(1000)

        # Build order map: machine_code -> best order (em_producao > first pendente by queue_position)
        order_map = {}
        for order in orders:
            code = order.get("machine_code")
            if code and order.get("status") == "em_producao":
                order_map[code] = order

        pendente_orders = sorted(
            [o for o in orders if o.get("status") == "pendente"],
            key=lambda o: o.get("queue_position", 0)
        )
        for order in pendente_orders:
            code = order.get("machine_code")
            if code and code not in order_map:
                order_map[code] = order

        result_machines = []
        for machine in machines:
            machine_code = machine.get("code")
            order = order_map.get(machine_code)
            order_data = None
            if order:
                order_data = {
                    "cliente": order.get("cliente", ""),
                    "artigo": order.get("artigo", ""),
                    "cor": order.get("cor", ""),
                    "quantidade": order.get("quantidade", ""),
                    "status": order.get("status", "")
                }
            result_machines.append({
                "code": machine_code,
                "status": machine.get("status", "verde"),
                "active": machine.get("active", True),
                "order": order_data
            })

        return {
            "machines": result_machines,
            "generated_at": get_brazil_time().isoformat(),
            "layout_type": layout_type
        }
    except Exception as e:
        logger.error(f"Error generating mapa trancadeiras report: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error generating report: {str(e)}")

# Include the router in the main app
app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("startup")
async def startup_event():
    machines_count = await db.machines.count_documents({})
    if machines_count == 0:
        await init_data()

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()
