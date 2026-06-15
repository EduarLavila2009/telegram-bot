import json
import logging
from datetime import date, datetime
from pathlib import Path

logger = logging.getLogger(__name__)

data_volume = Path("/data")
DB_PATH = (data_volume / "usuarios.json") if data_volume.is_dir() else (Path(__file__).resolve().parent / "usuarios.json")


def load_users() -> dict:
    if not DB_PATH.exists():
        return {"users": {}}
    try:
        with DB_PATH.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Error cargando usuarios.json: {e}")
        return {"users": {}}


def save_users(data: dict) -> None:
    try:
        with DB_PATH.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        logger.error(f"Error guardando usuarios.json: {e}")


def get_user(user_id: int | str) -> dict | None:
    data = load_users()
    return data.get("users", {}).get(str(user_id))


def is_subscription_active(user_id: int | str) -> bool:
    user = get_user(user_id)
    if not user:
        return False
    
    if user.get("role") == "admin":
        return True
        
    if user.get("status") != "active":
        return False
        
    exp_str = user.get("expiration_date")
    if exp_str == "never":
        return True
        
    try:
        exp_date = datetime.strptime(exp_str, "%Y-%m-%d").date()
        return date.today() <= exp_date
    except Exception:
        return False


def has_quota(user_id: int | str) -> bool:
    user = get_user(user_id)
    if not user:
        return False
        
    if user.get("role") == "admin":
        return True
        
    limit = user.get("limit_ops", -1)
    if limit == -1:
        return True
        
    consumed = user.get("consumed_ops", 0)
    return consumed < limit


def increment_usage(user_id: int | str) -> bool:
    data = load_users()
    users = data.get("users", {})
    uid_str = str(user_id)
    if uid_str not in users:
        return False
        
    # El admin no tiene límite y no necesita contar operaciones
    if users[uid_str].get("role") == "admin":
        return True
        
    users[uid_str]["consumed_ops"] = users[uid_str].get("consumed_ops", 0) + 1
    save_users(data)
    return True


def register_user(
    user_id: int | str,
    name: str,
    role: str,
    expiration_date: str,
    limit_ops: int,
) -> None:
    data = load_users()
    if "users" not in data:
        data["users"] = {}
        
    uid_str = str(user_id)
    existing_user = data["users"].get(uid_str, {})
    data["users"][uid_str] = {
        "name": name,
        "role": role,
        "status": "active",
        "expiration_date": expiration_date,
        "limit_ops": limit_ops,
        "consumed_ops": existing_user.get("consumed_ops", 0)
    }
    
    # Inicializar campos de empresa por defecto si el rol es nueva_empresa
    if role == "nueva_empresa":
        data["users"][uid_str]["company_name"] = existing_user.get("company_name", "FlashTax")
        data["users"][uid_str]["company_rif"] = existing_user.get("company_rif", "J-00000000-0")
        data["users"][uid_str]["company_type"] = existing_user.get("company_type", "Especial")
        data["users"][uid_str]["company_email"] = existing_user.get("company_email", "")
        data["users"][uid_str]["company_phone"] = existing_user.get("company_phone", "")
        data["users"][uid_str]["company_address"] = existing_user.get("company_address", "")
        
    save_users(data)


def update_user_status(user_id: int | str, status: str) -> bool:
    data = load_users()
    users = data.get("users", {})
    uid_str = str(user_id)
    if uid_str not in users:
        return False
        
    users[uid_str]["status"] = status
    save_users(data)
    return True


def update_user_field(user_id: int | str, field: str, value: any) -> bool:
    data = load_users()
    users = data.get("users", {})
    uid_str = str(user_id)
    if uid_str not in users:
        return False
        
    users[uid_str][field] = value
    save_users(data)
    return True
