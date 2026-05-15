"""
diagnostico_bd.py
Script para diagnosticar problemas con la BD y el modelo Predio
"""

import os
import sys
from database.models import db, Predio
from app import app

print("\n" + "="*70)
print("🔍 DIAGNÓSTICO DE BD Y MODELO PREDIO")
print("="*70 + "\n")

# 1. Verificar si BD existe
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, 'database', 'sic.db')

print(f"📁 Ruta BD: {DB_PATH}")
print(f"✓ BD existe: {os.path.exists(DB_PATH)}\n")

# 2. Verificar estructura del modelo Predio
print("📋 COLUMNAS DEL MODELO Predio:")
print("-" * 70)
try:
    for column in Predio.__table__.columns:
        print(f"  • {column.name:30} | Tipo: {column.type} | Nullable: {column.nullable}")
except Exception as e:
    print(f"❌ Error al leer columnas: {e}")

print("\n")

# 3. Intentar conectar a la BD
with app.app_context():
    try:
        # Verificar conexión
        print("🔌 Verificando conexión a BD...")
        result = db.session.execute(db.text("SELECT name FROM sqlite_master WHERE type='table'"))
        tablas = [row[0] for row in result]
        print(f"✓ Conexión OK")
        print(f"📊 Tablas en BD: {tablas}\n")
        
        # Si existe tabla 'predio', ver su estructura
        if 'predio' in tablas:
            print("🔍 COLUMNAS EN TABLA 'predio' (BD REAL):")
            print("-" * 70)
            result = db.session.execute(db.text("PRAGMA table_info(predio)"))
            for row in result:
                col_id, col_name, col_type, col_notnull, col_default, col_pk = row
                print(f"  • {col_name:30} | Tipo: {col_type:15} | PK: {col_pk} | NotNull: {col_notnull}")
            print("\n")
        else:
            print("⚠️  Tabla 'predio' NO existe en la BD\n")
        
        # 4. Intentar crear un predio de prueba
        print("🧪 INTENTANDO CREAR PREDIO DE PRUEBA...")
        print("-" * 70)
        
        try:
            nuevo_predio = Predio(
                id_predio='TEST_001',
                fmi='000-000-0001',
                propietario_vur='Test Usuario',
                nombre_predio_vur='Predio Test',
                area_vur=1000.0
            )
            db.session.add(nuevo_predio)
            db.session.commit()
            print("✓ Predio de prueba creado exitosamente")
            print(f"  ID: {nuevo_predio.id_predio}")
            print(f"  FMI: {nuevo_predio.fmi}\n")
            
            # Eliminar el predio de prueba
            db.session.delete(nuevo_predio)
            db.session.commit()
            print("✓ Predio de prueba eliminado\n")
        
        except Exception as e:
            print(f"❌ Error al crear predio de prueba: {e}")
            import traceback
            traceback.print_exc()
            print()
    
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        import traceback
        traceback.print_exc()

print("="*70)
print("FIN DEL DIAGNÓSTICO")
print("="*70 + "\n")